"""Запись результатов в .xlsx / .csv."""
from __future__ import annotations

import csv
import logging
import time
from pathlib import Path

from .config import (
    Config,
    _coerce_columns,
    normalize_column,
)

log = logging.getLogger("invoice_renamer")


class TableError(Exception):
    pass


def _sniff_delimiter(first_line: str) -> str:
    counts = {d: first_line.count(d) for d in (";", ",", "\t")}
    best = max(counts, key=counts.get)
    return best if counts[best] else ";"


def _clean_cell(value) -> str:
    """Убирает BOM и пробелы по краям ячейки."""
    return str(value).lstrip("\ufeff").strip()


def _read_csv(path: Path) -> tuple[str, list[str], list[list[str]]]:
    """Читает CSV: (delimiter, заголовок, строки данных).

    Автоматически:
      * снимает BOM с ячеек (артефакт старой записи через utf-8-sig в режиме "a");
      * выбрасывает пустые строки;
      * выкидывает повторные строки-заголовки (артефакт того же бага).
    """
    if not path.exists() or path.stat().st_size == 0:
        return ";", [], []

    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        first = fh.readline()
    if not first:
        return ";", [], []
    delimiter = _sniff_delimiter(first.rstrip("\r\n"))

    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        raw_rows = [
            [_clean_cell(c) for c in row]
            for row in csv.reader(fh, delimiter=delimiter)
        ]

    raw_rows = [r for r in raw_rows if any(c != "" for c in r)]
    if not raw_rows:
        return delimiter, [], []

    header = raw_rows[0][:]
    while header and header[-1] == "":
        header.pop()

    data: list[list[str]] = []
    for row in raw_rows[1:]:
        # Повторная шапка — пропускаем
        if row[: len(header)] == header:
            continue
        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))
        data.append(row[: len(header)])

    return delimiter, header, data


def _existing_headers(path: Path) -> list[str]:
    """Заголовки существующей таблицы ([] если файл пустой/отсутствует)."""
    if not path.exists():
        return []
    ext = path.suffix.lower()
    if ext == ".xlsx":
        from openpyxl import load_workbook
        try:
            wb = load_workbook(path)
        except Exception as e:
            raise TableError(f"Cannot open {path.name}: {e}")
        try:
            ws = wb.active
            if ws.max_row < 1:
                return []
            headers = [c.value for c in ws[1]]
            while headers and headers[-1] in (None, ""):
                headers.pop()
            return [str(h) if h is not None else "" for h in headers]
        finally:
            wb.close()
    if ext == ".csv":
        _, header, _ = _read_csv(path)
        return header
    raise TableError("Table must be .xlsx or .csv")


def table_preflight(cfg: Config) -> None:
    if not cfg.table_enabled:
        return
    path = Path(cfg.table_path)
    if path.suffix.lower() not in (".xlsx", ".csv"):
        raise TableError("Table must be .xlsx or .csv")
    if not path.exists():
        if not cfg.table_create_if_missing:
            raise TableError(f"Table file does not exist: {path}")
        return
    existing = _existing_headers(path)
    if not existing:
        return  # пустая — будет создана при первой записи
    cols = _coerce_columns(cfg.table_columns)
    # id управляется программой — его не требуем в уже созданной таблице
    required = [c["header"] for c in cols if c["key"] != "id"]
    missing = [h for h in required if h not in existing]
    if missing:
        raise TableError(
            f"Table “{path.name}” is missing required columns: "
            + ", ".join(repr(m) for m in missing)
            + f". Existing columns: {existing}"
        )


def _format_value(val, typ: str):
    if val is None or val == "":
        return ""
    if typ == "null":
        return ""
    if typ == "integer":
        try:
            return int(float(str(val)))
        except (ValueError, TypeError):
            return str(val)
    if typ == "number":
        try:
            return float(str(val))
        except (ValueError, TypeError):
            return str(val)
    if typ == "boolean":
        if isinstance(val, bool):
            return val
        return str(val).strip().lower() in ("1", "true", "yes", "да")
    return str(val)


def _next_id(path: Path, id_header: str) -> int:
    """Следующий ID: max(id) + 1 по колонке с заголовком id_header."""
    if not path.exists() or path.stat().st_size == 0:
        return 1
    ext = path.suffix.lower()
    if ext == ".xlsx":
        from openpyxl import load_workbook
        try:
            wb = load_workbook(path)
        except Exception:
            return 1
        try:
            ws = wb.active
            if ws.max_row < 1:
                return 1
            headers = [str(c.value) if c.value is not None else "" for c in ws[1]]
            try:
                col = headers.index(id_header) + 1
            except ValueError:
                return 1
            max_id = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if col - 1 < len(row):
                    try:
                        n = int(row[col - 1])
                        if n > max_id:
                            max_id = n
                    except (ValueError, TypeError):
                        pass
            return max_id + 1
        finally:
            wb.close()
    if ext == ".csv":
        _, header, data = _read_csv(path)
        if not header:
            return 1
        try:
            col = header.index(id_header)
        except ValueError:
            return 1
        max_id = 0
        for row in data:
            if col < len(row):
                try:
                    n = int(str(row[col]).strip())
                    if n > max_id:
                        max_id = n
                except (ValueError, TypeError):
                    pass
        return max_id + 1
    return 1


def _align_row(existing: list[str], headers: list[str], values: list) -> list:
    """Раскладываем значения по фактическому порядку колонок в таблице."""
    lookup = {h: v for h, v in zip(headers, values)}
    return [lookup.get(str(h), "") for h in existing]


def table_append(cfg: Config, row: dict) -> None:
    if not cfg.table_enabled:
        return
    path = Path(cfg.table_path)
    cols = _coerce_columns(cfg.table_columns)
    existing = _existing_headers(path)

    # Авто-ID: для новой таблицы добавляем колонку, для существующей —
    # только если в ней уже есть подходящий столбец (ID / № / номер).
    if cfg.table_auto_id and not any(c["key"] == "id" for c in cols):
        if not existing:
            cols = [normalize_column({"key": "id", "header": "ID"})] + cols
        else:
            id_h = next(
                (h for h in existing
                 if str(h).strip().lower() in ("id", "№", "номер", "no")),
                None,
            )
            if id_h:
                cols = [normalize_column({"key": "id", "header": str(id_h)})] + cols

    has_id = any(c["key"] == "id" for c in cols)
    if has_id and "id" not in row:
        id_col = next(c for c in cols if c["key"] == "id")
        # Пишем id только если столбец реально есть/будет в таблице
        if not existing or id_col["header"] in existing:
            row["id"] = _next_id(path, id_col["header"])

    headers = [c["header"] for c in cols]
    values = [_format_value(row.get(c["key"], ""), c.get("type", "string")) for c in cols]

    path.parent.mkdir(parents=True, exist_ok=True)
    last_error: Exception | None = None
    for _ in range(5):
        try:
            if path.suffix.lower() == ".xlsx":
                _append_xlsx(path, headers, values)
            else:
                _append_csv(path, headers, values)
            return
        except PermissionError as e:
            last_error = e
            time.sleep(2)
    raise TableError(f"Cannot write {path.name} (is it open in Excel?): {last_error}")


def _append_xlsx(path: Path, headers: list, values: list) -> None:
    from openpyxl import Workbook, load_workbook

    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Счета"

    existing: list[str] = []
    if ws.max_row >= 1:
        candidate = [c.value for c in ws[1]]
        while candidate and candidate[-1] in (None, ""):
            candidate.pop()
        existing = [str(h) if h is not None else "" for h in candidate]

    if not existing:
        ws.append(headers)
        ws.append(values)
        ws.freeze_panes = "A2"
        for i, h in enumerate(headers, start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = max(
                14, len(str(h)) + 4
            )
    else:
        ws.append(_align_row(existing, headers, values))

    wb.save(path)
    wb.close()


def _append_csv(path: Path, headers: list, values: list) -> None:
    delimiter, existing, data = _read_csv(path)

    # Новый файл — пишем шапку и первую строку. BOM уместен только здесь.
    if not existing:
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh, delimiter=delimiter)
            w.writerow(headers)
            w.writerow(values)
        return

    new_row = _align_row(existing, headers, values)

    # Проверяем, "чистый" ли файл: непустых строк должно быть ровно
    # len(data) + 1 (шапка + данные). Если больше — были дубли шапок
    # или другие артефакты — перезаписываем файл целиком.
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        raw_count = sum(
            1 for r in csv.reader(fh, delimiter=delimiter)
            if any(_clean_cell(c) != "" for c in r)
        )
    is_clean = (raw_count == len(data) + 1)

    if not is_clean:
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh, delimiter=delimiter)
            w.writerow(existing)
            for r in data:
                w.writerow(r)
            w.writerow(new_row)
        log.warning(
            "Table %s contained duplicate headers/blank rows — rewritten cleanly.",
            path.name,
        )
        return

    # Обычная дозапись — БЕЗ BOM (иначе Excel-friendly BOM попадёт
    # внутрь данных и сломает и заголовки, и ID).
    with path.open("rb") as fh:
        fh.seek(-1, 2)
        needs_newline = fh.read(1) not in (b"\n", b"\r")

    with path.open("a", encoding="utf-8", newline="") as fh:
        if needs_newline:
            fh.write("\r\n")
        csv.writer(fh, delimiter=delimiter).writerow(new_row)