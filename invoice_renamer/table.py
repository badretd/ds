"""Запись результатов в .xlsx / .csv."""
from __future__ import annotations

import csv
import logging
import time
from pathlib import Path

from .config import Config, _coerce_columns

log = logging.getLogger("invoice_renamer")


class TableError(Exception):
    pass


def _sniff_delimiter(first_line: str) -> str:
    counts = {d: first_line.count(d) for d in (";", ",", "\t")}
    best = max(counts, key=counts.get)
    return best if counts[best] else ";"


def _xlsx_header_count(ws) -> int:
    header = [c.value for c in ws[1]]
    while header and header[-1] in (None, ""):
        header.pop()
    return len(header)


def _existing_column_count(path: Path) -> int:
    from openpyxl import load_workbook

    ext = path.suffix.lower()
    if ext == ".xlsx":
        try:
            wb = load_workbook(path)
        except Exception as e:
            raise TableError(f"Cannot open {path.name}: {e}")
        try:
            return _xlsx_header_count(wb.active)
        finally:
            wb.close()
    if ext == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            first = fh.readline().rstrip("\r\n")
        if not first:
            return 0
        d = _sniff_delimiter(first)
        return len(next(csv.reader([first], delimiter=d)))
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
    have = _existing_column_count(path)
    need = len(_coerce_columns(cfg.table_columns))
    if 0 < have < need:
        raise TableError(
            f"Table “{path.name}” has {have} column(s), but {need} are required. "
            "Add the missing columns (extra columns are fine)."
        )


def table_append(cfg: Config, row: dict) -> None:
    if not cfg.table_enabled:
        return
    path = Path(cfg.table_path)
    cols = _coerce_columns(cfg.table_columns)
    headers = [c["header"] for c in cols]
    values = [row.get(c["key"], "") for c in cols]
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
        if _xlsx_header_count(ws) == 0:
            ws.append(headers)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Счета"
        ws.append(headers)
        ws.freeze_panes = "A2"
        for i, h in enumerate(headers, start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = max(14, len(h) + 4)
    ws.append(values)
    wb.save(path)
    wb.close()


def _append_csv(path: Path, headers: list, values: list) -> None:
    delimiter, have, needs_newline = ";", 0, False
    if path.exists() and path.stat().st_size > 0:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            first = fh.readline().rstrip("\r\n")
        delimiter = _sniff_delimiter(first)
        have = len(next(csv.reader([first], delimiter=delimiter)))
        with path.open("rb") as fh:
            fh.seek(-1, 2)
            needs_newline = fh.read(1) not in (b"\n", b"\r")
    with path.open("a", encoding="utf-8-sig", newline="") as fh:
        if needs_newline:
            fh.write("\r\n")
        w = csv.writer(fh, delimiter=delimiter)
        if have == 0:
            w.writerow(headers)
        w.writerow(values + [""] * max(0, have - len(values)))
