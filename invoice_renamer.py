#!/usr/bin/env python3
"""
Invoice Renamer
===============
Watches a folder for PDF invoices (Russian language), renders pages to images,
OCRs them with Tesseract, asks a local LLM (Ollama) to extract
date / supplier / invoice number / main material, renames the file, e.g.

    260924 АРГО Сч № КА-16339 металл.pdf

and appends a record to an Excel (.xlsx) or CSV table.

Run with the settings UI:      python invoice_renamer.py
Run headless (no window):      python invoice_renamer.py --headless
Use another config file:       python invoice_renamer.py --config my.json
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import logging.handlers
import queue
import re
import shutil
import sys
import threading
import time
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF – renders PDF pages to images, no external poppler needed
import pytesseract
import requests
from openpyxl import Workbook, load_workbook
from PIL import Image
from pytesseract import TesseractNotFoundError

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
LOG_PATH = APP_DIR / "invoice_renamer.log"

REQUIRED_TOKEN = "Сч"  # must be present in the filename template
TEMPLATE_FIELDS = ("date", "supplier", "invoice", "material")

# Table column keys -> header titles
TABLE_HEADERS = {
    "date": "Дата",
    "supplier": "Поставщик",
    "invoice": "Номер счёта",
    "material": "Материал",
    "original_name": "Исходное имя файла",
    "new_name": "Новое имя файла",
    "processed_at": "Обработано",
}

log = logging.getLogger("invoice_renamer")


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    # Folders / scheduling
    watch_dir: str = ""
    output_dir: str = ""              # empty = rename in place
    recursive: bool = False
    poll_seconds: int = 10
    stable_seconds: int = 3           # file must be untouched this long
    move_failed: bool = True          # move unrecognised files to failed_dir_name
    failed_dir_name: str = "_failed"

    # OCR
    tesseract_cmd: str = ""           # empty = use PATH
    ocr_lang: str = "rus+eng"
    ocr_dpi: int = 300
    max_pages: int = 3
    tesseract_config: str = "--oem 1 --psm 6"
    use_embedded_text: bool = False   # use PDF text layer when present (skips OCR)

    # LLM
    ollama_url: str = "http://localhost:11434"
    model: str = "qwen2.5:7b"
    num_ctx: int = 8192
    timeout: int = 180
    max_chars: int = 12000
    keep_alive: str = "30m"           # how long Ollama keeps the model in RAM

    # Naming
    filename_template: str = "{date} {supplier} Сч № {invoice} {material}"
    date_format: str = "%d%m%y"
    invoice_cyrillic: bool = True     # K A-123 -> КА-123 (Latin look-alikes -> Cyrillic)
    max_name_len: int = 150

    # Table
    table_enabled: bool = True
    table_path: str = ""
    table_create_if_missing: bool = True
    table_columns: list = field(default_factory=lambda: list(TABLE_HEADERS))


def load_config(path: Path | None = None) -> Config:
    path = path or CONFIG_PATH
    cfg = Config()
    if path.exists():
        try:
            data = json.loads(path.read_text("utf-8"))
        except (OSError, ValueError):
            data = {}
        for f in fields(Config):
            if f.name in data:
                setattr(cfg, f.name, data[f.name])
    return cfg


def save_config(cfg: Config, path: Path | None = None) -> None:
    path = path or CONFIG_PATH
    path.write_text(json.dumps(asdict(cfg), ensure_ascii=False, indent=2), "utf-8")


class ConfigError(Exception):
    pass


def validate_config(cfg: Config) -> None:
    if REQUIRED_TOKEN not in cfg.filename_template:
        raise ConfigError(f"The filename template must contain “{REQUIRED_TOKEN}”.")
    try:
        cfg.filename_template.format(**{k: "x" for k in TEMPLATE_FIELDS})
    except (KeyError, IndexError, ValueError) as e:
        raise ConfigError(
            f"Bad filename template ({e!r}). Allowed placeholders: "
            + ", ".join("{%s}" % k for k in TEMPLATE_FIELDS)
        )
    try:
        datetime(2024, 9, 26).strftime(cfg.date_format)
    except ValueError as e:
        raise ConfigError(f"Bad date format: {e}")
    bad = [c for c in cfg.table_columns if c not in TABLE_HEADERS]
    if bad or not cfg.table_columns:
        raise ConfigError(
            "Table columns must be a non-empty list of: " + ", ".join(TABLE_HEADERS)
        )
    for name in ("poll_seconds", "ocr_dpi", "max_pages", "num_ctx", "timeout", "max_chars"):
        if int(getattr(cfg, name)) <= 0:
            raise ConfigError(f"“{name}” must be a positive number.")
    if cfg.table_enabled and not cfg.table_path.strip():
        raise ConfigError("Table logging is enabled but the table path is empty.")


# --------------------------------------------------------------------------- #
# Text extraction (PDF -> image -> Tesseract)
# --------------------------------------------------------------------------- #
class ExtractionError(Exception):
    """Document-specific problem: the file is moved to the failed folder."""


class ServiceUnavailable(Exception):
    """Environment problem (Ollama / Tesseract down): the file is retried later."""


def extract_text(cfg: Config, path: Path) -> str:
    parts: list[str] = []
    try:
        doc = fitz.open(path)
    except Exception as e:  # corrupted / not a PDF
        raise ExtractionError(f"Cannot open PDF: {e}")
    with doc:
        if doc.needs_pass:
            raise ExtractionError("PDF is password protected")
        for i, page in enumerate(doc):
            if i >= cfg.max_pages:
                break
            text = page.get_text().strip() if cfg.use_embedded_text else ""
            if len(text) < 40:
                pix = page.get_pixmap(dpi=cfg.ocr_dpi)  # rendered to a PNG-like raster
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
                try:
                    text = pytesseract.image_to_string(
                        img, lang=cfg.ocr_lang, config=cfg.tesseract_config
                    )
                except TesseractNotFoundError as e:
                    raise ServiceUnavailable(f"Tesseract not found: {e}")
                except pytesseract.TesseractError as e:
                    raise ServiceUnavailable(
                        f"Tesseract failed (is the language “{cfg.ocr_lang}” installed?): {e}"
                    )
            parts.append(text)
    full = "\n\n".join(parts).strip()
    if len(full) < 20:
        raise ExtractionError("No text recognised")
    return full


# --------------------------------------------------------------------------- #
# Deterministic hints: invoice number / date from the document header
# --------------------------------------------------------------------------- #
MONTHS_RU = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5,
    "июня": 6, "июля": 7, "августа": 8, "сентября": 9, "октября": 10,
    "ноября": 11, "декабря": 12,
}

# "Счёт на оплату № КА-16339"  /  "Счёт на оплату № 12345"
INVOICE_HINT_RE = re.compile(
    r'Сч[её]т\s+на\s+оплату\s*[№N#:]?\s*'
    r'([A-ZА-ЯЁа-яёa-z]{0,6}[\s\-–—/_]?\d{1,15})',
    re.IGNORECASE,
)

# "Счёт на оплату № … от 26.09.2024"
DATE_NUM_HINT_RE = re.compile(
    r'Сч[её]т\s+на\s+оплату[^\n]{0,150}?'
    r'от\s+(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})',
    re.IGNORECASE,
)

# "Счёт на оплату № … от 26 сентября 2024 г."
DATE_WORD_HINT_RE = re.compile(
    r'от\s+(\d{1,2})\s+(' + '|'.join(MONTHS_RU) + r')\s+(\d{4})',
    re.IGNORECASE,
)


def extract_hints(text: str) -> dict:
    """Regex-based fallback values pulled from the invoice header.

    Anchored to "Счёт на оплату № … от …" — the most reliable place in a
    Russian invoice.  If the anchor isn't found, we return {} and let the
    LLM do the work.
    """
    hints: dict[str, str] = {}
    if not text:
        return hints

    m = INVOICE_HINT_RE.search(text)
    if m:
        hints["invoice"] = m.group(1).strip()

    m = DATE_NUM_HINT_RE.search(text)
    if m:
        hints["date"] = m.group(1)
    else:
        m = DATE_WORD_HINT_RE.search(text)
        if m:
            day, month_name, year = m.groups()
            month = MONTHS_RU.get(month_name.lower())
            if month:
                hints["date"] = f"{int(day):02d}.{month:02d}.{year}"

    return hints


# --------------------------------------------------------------------------- #
# LLM (Ollama)
# --------------------------------------------------------------------------- #
SYSTEM_PROMPT = """Ты — помощник бухгалтера. Тебе дают текст счёта на русском языке, \
полученный через OCR (возможны ошибки распознавания). Извлеки поля и верни ТОЛЬКО JSON.

Поля:

- date: дата выставления счёта в формате YYYY-MM-DD.
  В российском счёте она почти всегда стоит в САМОЙ ВЕРХНЕЙ части документа,
  сразу после номера счёта, во фразе вида:
      «Счёт на оплату № КА-16339 от 26.09.2024»
      «Счёт на оплату № 12345 от 26 сентября 2024 г.»
  Бери именно эту дату — ту, что идёт после слова «от» в строке с «Счёт на оплату».
  НЕ путай её с:
    * сроком оплаты («Оплатить до …», «Срок оплаты …», «Счёт действителен до …»),
    * датой отгрузки / поставки / приёмки,
    * датой договора или спецификации,
    * датой в счёте-фактуре / УПД / акте, если они приложены ниже.
  Если дата записана словами («26 сентября 2024 г.») — переведи в числа.
  Если в шапке стоит только день и месяц без года — верни пустую строку.

- supplier: поставщик (продавец, получатель платежа, «Исполнитель», «Поставщик»), \
НЕ покупатель и НЕ плательщик. Верни короткое название без организационно-правовой формы \
и без кавычек, например: АРГО.

- invoice_number: номер счёта ровно как напечатан.
  Ключевой маркер в шапке документа: «Счёт на оплату № …» — бери то, что стоит \
СРАЗУ ПОСЛЕ него (до слова «от» или до конца строки).
  Номер может быть:
    * с буквами: КА-16339, АБ123, С-42, AB-12;
    * только цифрами: 105770, 1234, 00000123.
  Если в номере есть буквы — приводи к виду «буквы-цифры» (КА-16339).
  Если букв нет — оставь только цифры, без ведущих нулей и лишних пробелов.
  НЕ бери в качестве номера счёта: ИНН, КПП, ОГРН, ОКПО, номера банковских счетов, \
БИК, номера договоров/спецификаций, номера счетов-фактур и УПД (если они отличаются \
от номера самого счёта на оплату).

- material: основной материал или товар заказа (на который приходится основная часть суммы), \
одним-двумя словами, строчными буквами, в именительном падеже, общим словом: \
например «металл», «бетон», «кабель», «пластик».

Если поле найти невозможно — верни пустую строку."""

SCHEMA = {
    "type": "object",
    "properties": {
        "date": {"type": "string"},
        "supplier": {"type": "string"},
        "invoice_number": {"type": "string"},
        "material": {"type": "string"},
    },
    "required": ["date", "supplier", "invoice_number", "material"],
}


def ask_llm(cfg: Config, text: str) -> dict:
    payload = {
        "model": cfg.model,
        "stream": False,
        "format": SCHEMA,
        "keep_alive": cfg.keep_alive,
        "options": {
            "temperature": 0,
            "num_ctx": cfg.num_ctx,
            "num_predict": 256,          # ответ — короткий JSON, больше не нужно
        },
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": 'Текст документа:\n"""\n' + text[: cfg.max_chars] + '\n"""',
            },
        ],
    }
    url = cfg.ollama_url.rstrip("/") + "/api/chat"
    try:
        r = requests.post(url, json=payload, timeout=cfg.timeout)
    except requests.ConnectionError as e:
        raise ServiceUnavailable(f"Ollama is not reachable at {cfg.ollama_url}: {e}")
    except requests.Timeout:
        # A slow model should NOT send the file to _failed — retry next scan.
        raise ServiceUnavailable(f"Ollama timed out after {cfg.timeout}s")
    if r.status_code == 404:
        raise ServiceUnavailable(
            f"Model “{cfg.model}” not found in Ollama. Run: ollama pull {cfg.model}"
        )
    if not r.ok:
        raise ExtractionError(f"Ollama error {r.status_code}: {r.text[:200]}")
    try:
        return json.loads(r.json()["message"]["content"])
    except (ValueError, KeyError, TypeError) as e:
        raise ExtractionError(f"LLM returned invalid JSON: {e}")


def test_ollama(cfg: Config) -> str:
    """Returns a human-readable status; raises on failure."""
    r = requests.get(cfg.ollama_url.rstrip("/") + "/api/tags", timeout=10)
    r.raise_for_status()
    names = [m.get("name", "") for m in r.json().get("models", [])]
    wanted = cfg.model if ":" in cfg.model else cfg.model + ":latest"
    if wanted in names or cfg.model in names:
        return f"Ollama OK. Model “{cfg.model}” is available."
    return (
        f"Ollama is running, but model “{cfg.model}” is not installed.\n"
        f"Installed: {', '.join(names) or 'none'}\nRun: ollama pull {cfg.model}"
    )


# --------------------------------------------------------------------------- #
# Normalisation & naming
# --------------------------------------------------------------------------- #
LAT2CYR = str.maketrans("ABCEHKMOPTXY", "АВСЕНКМОРТХУ")
LEGAL_FORMS = re.compile(
    r"^(общество\s+с\s+ограниченной\s+ответственностью|публичное\s+акционерное\s+общество|"
    r"акционерное\s+общество|ооо|оао|зао|пао|ао|ип|тоо|llc|ltd|gmbh)\b[\s,.]*",
    re.IGNORECASE,
)
FORBIDDEN = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def parse_date(value: str) -> datetime | None:
    value = (value or "").strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            d = datetime.strptime(value, fmt)
        except ValueError:
            continue
        if 2000 <= d.year <= 2100:
            return d
    return None


def normalize_invoice(value: str, to_cyrillic: bool) -> str:
    """Return a clean invoice number.

    Accepts:
      * "КА-16339" / "KA-16339" / "АБ 123" / "С42"  → "КА-16339" (letters-digits)
      * "12345" / "№ 12345" / "105770,18"           → "12345"    (digits only)
    """
    s = (value or "").strip()
    s = s.replace("№", " ").replace("#", " ")
    s = s.upper()
    # letters followed (optionally) by a separator and digits
    m = re.search(r'([A-ZА-ЯЁ]{1,6})\s*[-–—/_]?\s*(\d{1,15})', s)
    if m:
        result = f"{m.group(1)}-{m.group(2)}"
    else:
        digits = re.search(r'\d{1,15}', s)
        if not digits:
            return ""
        result = digits.group(0)
    return result.translate(LAT2CYR) if to_cyrillic else result


def normalize_supplier(value: str) -> str:
    s = (value or "").strip()
    s = re.sub(r"[«»“”\"']", "", s).strip()
    s = LEGAL_FORMS.sub("", s).strip(" ,.-")
    return s


def normalize_fields(cfg: Config, data: dict, hints: dict | None = None) -> dict:
    """Combine the LLM output with deterministic hints.

    Priority:  regex hint (literal text from the header) > LLM value > empty.
    Mismatches are logged so you can spot prompt regressions.
    """
    hints = hints or {}

    date_src = (hints.get("date") or str(data.get("date", ""))).strip()
    invoice_src = (hints.get("invoice") or str(data.get("invoice_number", ""))).strip()

    # Log disagreements — useful when tuning the prompt.
    llm_inv = normalize_invoice(str(data.get("invoice_number", "")), cfg.invoice_cyrillic)
    hint_inv = normalize_invoice(hints.get("invoice", ""), cfg.invoice_cyrillic)
    if hint_inv and llm_inv and hint_inv != llm_inv:
        log.warning("Invoice mismatch: LLM=%r, regex=%r → using regex", llm_inv, hint_inv)

    d = parse_date(date_src)
    fields_ = {
        "date": d.strftime(cfg.date_format) if d else "",
        "supplier": normalize_supplier(str(data.get("supplier", ""))),
        "invoice": normalize_invoice(invoice_src, cfg.invoice_cyrillic),
        "material": str(data.get("material", "")).strip().lower().strip(".,;"),
    }
    missing = [k for k in ("date", "supplier", "invoice") if not fields_[k]]
    if missing:
        raise ExtractionError(
            f"Could not extract: {', '.join(missing)} "
            f"(LLM={data!r}, hints={hints!r})"
        )
    return fields_


def build_filename(cfg: Config, fields_: dict) -> str:
    stem = cfg.filename_template.format(**fields_)
    stem = FORBIDDEN.sub("", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .")
    stem = stem[: cfg.max_name_len].rstrip(" .")
    return stem + ".pdf"


def unique_target(directory: Path, name: str, source: Path) -> Path:
    target = directory / name
    n = 2
    while target.exists() and target.resolve() != source.resolve():
        target = directory / f"{Path(name).stem} ({n}).pdf"
        n += 1
    return target


# --------------------------------------------------------------------------- #
# Table (xlsx / csv)
# --------------------------------------------------------------------------- #
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
    """Fail early (before any file is touched) if the table cannot be used."""
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
    need = len(cfg.table_columns)
    if 0 < have < need:
        raise TableError(
            f"Table “{path.name}” has {have} column(s), but {need} are required. "
            "Add the missing columns (extra columns are fine)."
        )


def table_append(cfg: Config, row: dict) -> None:
    if not cfg.table_enabled:
        return
    path = Path(cfg.table_path)
    headers = [TABLE_HEADERS[k] for k in cfg.table_columns]
    values = [row.get(k, "") for k in cfg.table_columns]
    path.parent.mkdir(parents=True, exist_ok=True)
    last_error: Exception | None = None
    for _ in range(5):  # the file may be open in Excel
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
    ws.append(values)  # extra existing columns stay blank
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


# --------------------------------------------------------------------------- #
# Worker
# --------------------------------------------------------------------------- #
ALREADY_NAMED = re.compile(r"^\d{4,8} .+ Сч\s*№", re.IGNORECASE)


class Worker(threading.Thread):
    def __init__(self, cfg: Config):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.stop_event = threading.Event()
        self.seen: set[tuple] = set()
        if cfg.tesseract_cmd.strip():
            pytesseract.pytesseract.tesseract_cmd = cfg.tesseract_cmd.strip()

    # -- lifecycle ----------------------------------------------------------
    def run(self) -> None:
        log.info("Started. Watching: %s", self.cfg.watch_dir)
        try:
            if "rus" in self.cfg.ocr_lang and "rus" not in pytesseract.get_languages():
                log.warning("Tesseract language data “rus” is not installed!")
        except Exception:
            pass
        while not self.stop_event.is_set():
            try:
                self.scan_once()
            except Exception:
                log.exception("Unexpected error during scan")
            self.stop_event.wait(max(1, int(self.cfg.poll_seconds)))
        log.info("Stopped.")

    def stop(self) -> None:
        self.stop_event.set()

    # -- scanning -----------------------------------------------------------
    def _candidates(self):
        root = Path(self.cfg.watch_dir)
        failed = root / self.cfg.failed_dir_name
        out = Path(self.cfg.output_dir) if self.cfg.output_dir.strip() else None
        it = root.rglob("*") if self.cfg.recursive else root.glob("*")
        for p in sorted(it):
            if not p.is_file() or p.suffix.lower() != ".pdf":
                continue
            if failed in p.parents or (out and out in p.parents):
                continue
            if ALREADY_NAMED.match(p.name):
                continue
            yield p

    def _ready(self, p: Path):
        try:
            st = p.stat()
        except OSError:
            return None
        if time.time() - st.st_mtime < self.cfg.stable_seconds:
            return None
        key = (str(p), st.st_mtime_ns, st.st_size)
        return None if key in self.seen else key

    def scan_once(self) -> None:
        cfg = self.cfg
        root = Path(cfg.watch_dir)
        if not cfg.watch_dir.strip() or not root.is_dir():
            log.error("Watch folder does not exist: %r", cfg.watch_dir)
            return
        try:
            validate_config(cfg)
            table_preflight(cfg)
        except (ConfigError, TableError) as e:
            log.error("%s — processing paused.", e)
            return
        for pdf in self._candidates():
            if self.stop_event.is_set():
                return
            key = self._ready(pdf)
            if key is None:
                continue
            try:
                self.process(pdf)
            except ServiceUnavailable as e:
                log.error("%s — will retry on the next scan.", e)
                return
            except ExtractionError as e:
                log.error("%s: %s", pdf.name, e)
                self.seen.add(key)
                self._handle_failed(pdf)
            except Exception:
                log.exception("%s: unexpected error", pdf.name)
                self.seen.add(key)

    # -- single document ----------------------------------------------------
    def process(self, pdf: Path) -> Path:
        cfg = self.cfg
        log.info("Processing %s", pdf.name)
        text = extract_text(cfg, pdf)

        # Deterministic hints from the header — used as primary source for
        # invoice number and date, and as a fallback if the LLM misses them.
        hints = extract_hints(text)
        if hints:
            log.info("Header hints: %s", hints)

        data = ask_llm(cfg, text)
        fields_ = normalize_fields(cfg, data, hints)
        new_name = build_filename(cfg, fields_)

        dest_dir = Path(cfg.output_dir) if cfg.output_dir.strip() else pdf.parent
        dest_dir.mkdir(parents=True, exist_ok=True)
        target = unique_target(dest_dir, new_name, pdf)
        shutil.move(str(pdf), str(target))
        log.info("OK: %s  ->  %s", pdf.name, target.name)

        st = target.stat()
        self.seen.add((str(target), st.st_mtime_ns, st.st_size))
        row = {
            **fields_,
            "original_name": pdf.name,
            "new_name": target.name,
            "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        try:
            table_append(cfg, row)
        except TableError as e:
            log.error("File renamed but the table row was NOT written: %s | row=%s", e, row)
        return target

    def _handle_failed(self, pdf: Path) -> None:
        if not self.cfg.move_failed:
            return
        try:
            failed = Path(self.cfg.watch_dir) / self.cfg.failed_dir_name
            failed.mkdir(parents=True, exist_ok=True)
            target = unique_target(failed, pdf.name, pdf)
            shutil.move(str(pdf), str(target))
            log.info("Moved to %s", target)
        except OSError as e:
            log.error("Could not move %s to the failed folder: %s", pdf.name, e)


# --------------------------------------------------------------------------- #
# Logging
# --------------------------------------------------------------------------- #
class QueueHandler(logging.Handler):
    def __init__(self, q: queue.Queue):
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        self.q.put(self.format(record))


def setup_logging(q: queue.Queue | None = None) -> None:
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s", "%H:%M:%S")
    handlers: list[logging.Handler] = [
        logging.handlers.RotatingFileHandler(
            LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
    ]
    if q is None:
        handlers.append(logging.StreamHandler(sys.stdout))
    else:
        handlers.append(QueueHandler(q))
    for h in handlers:
        h.setFormatter(fmt)
        log.addHandler(h)


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
def run_ui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    class App(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title("Invoice Renamer")
            self.geometry("860x660")
            self.cfg = load_config()
            self.vars: dict[str, tk.Variable] = {}
            self.kinds: dict[str, str] = {}
            self.worker: Worker | None = None
            self.log_q: queue.Queue = queue.Queue()
            setup_logging(self.log_q)
            self._build()
            self.after(200, self._drain_log)
            self.protocol("WM_DELETE_WINDOW", self._close)

        # ---- layout --------------------------------------------------------
        def _build(self) -> None:
            bar = ttk.Frame(self, padding=(10, 8))
            bar.pack(fill="x")
            self.btn_start = ttk.Button(bar, text="▶ Start", command=self.start)
            self.btn_stop = ttk.Button(bar, text="■ Stop", command=self.stop, state="disabled")
            self.btn_once = ttk.Button(bar, text="Scan once", command=self.scan_once)
            for b in (self.btn_start, self.btn_stop, self.btn_once):
                b.pack(side="left", padx=(0, 6))
            ttk.Button(bar, text="Save settings", command=self.save).pack(side="left", padx=(12, 6))
            ttk.Button(bar, text="Test Ollama", command=self.test_llm).pack(side="left")
            self.status = ttk.Label(bar, text="Stopped", foreground="gray")
            self.status.pack(side="right")

            self.nb = ttk.Notebook(self)
            self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 6))

            t = self._tab("Folders")
            self._row(t, "watch_dir", "Watch folder", "dir")
            self._row(t, "output_dir", "Output folder", "dir", "Empty = rename in place")
            self._row(t, "recursive", "Include subfolders", "bool")
            self._row(t, "poll_seconds", "Scan interval (sec)", "int")
            self._row(t, "stable_seconds", "File must be idle for (sec)", "int",
                      "Avoids reading files that are still being copied")
            self._row(t, "move_failed", "Move unrecognised files", "bool")
            self._row(t, "failed_dir_name", "…into subfolder", "str")

            t = self._tab("OCR")
            self._row(t, "tesseract_cmd", "Tesseract executable", "file", "Empty = found via PATH")
            self._row(t, "ocr_lang", "Languages", "str", "e.g. rus+eng (needs the “rus” language data)")
            self._row(t, "ocr_dpi", "Render DPI", "int", "300 is a good default; 200 is faster")
            self._row(t, "max_pages", "Max pages to read", "int")
            self._row(t, "tesseract_config", "Tesseract options", "str")
            self._row(t, "use_embedded_text", "Use PDF text layer if present", "bool",
                      "Skips OCR for digital (non-scanned) PDFs")

            t = self._tab("LLM (Ollama)")
            self._row(t, "ollama_url", "Ollama URL", "str")
            self._row(t, "model", "Model", "str", "Suggested: qwen2.5:7b, qwen2.5:14b, llama3.1:8b")
            self._row(t, "num_ctx", "Context size (tokens)", "int")
            self._row(t, "timeout", "Timeout (sec)", "int")
            self._row(t, "max_chars", "Max characters sent", "int")
            self._row(t, "keep_alive", "Keep model in RAM for", "str",
                      "e.g. 30m, 1h — avoids reloading on every file")

            t = self._tab("Naming")
            self._row(t, "filename_template", "Filename template", "str",
                      "Placeholders: {date} {supplier} {invoice} {material}. “Сч” is required.")
            self._row(t, "date_format", "Date format", "str", "strftime: %d%m%y → 260924")
            self._row(t, "invoice_cyrillic", "Latin look-alikes → Cyrillic in invoice no.", "bool",
                      "KA-16339 → КА-16339")
            self._row(t, "max_name_len", "Max name length", "int")
            pr = ttk.Frame(t)
            pr.grid(row=t._r, column=0, columnspan=3, sticky="ew", pady=10)
            ttk.Button(pr, text="Preview", command=self.preview).pack(side="left")
            self.preview_lbl = ttk.Label(pr, text="")
            self.preview_lbl.pack(side="left", padx=8)

            t = self._tab("Table")
            self._row(t, "table_enabled", "Write results to a table", "bool")
            self._row(t, "table_path", "Table file (.xlsx / .csv)", "table")
            self._row(t, "table_create_if_missing", "Create file if it does not exist", "bool")
            self._row(t, "table_columns", "Columns (in order)", "list",
                      "Keys: " + ", ".join(TABLE_HEADERS))
            ttk.Label(
                t, foreground="gray", wraplength=640, justify="left",
                text=("If the existing table has FEWER columns than listed here, nothing is processed "
                      "and an error is logged. If it has MORE, the extra columns stay blank."),
            ).grid(row=t._r, column=0, columnspan=3, sticky="w", pady=10)

            t = ttk.Frame(self.nb, padding=4)
            self.nb.add(t, text="Log")
            self.log_box = scrolledtext.ScrolledText(t, height=10, state="disabled", wrap="word")
            self.log_box.pack(fill="both", expand=True)

        def _tab(self, title: str) -> ttk.Frame:
            f = ttk.Frame(self.nb, padding=12)
            f.columnconfigure(1, weight=1)
            f._r = 0  # next free grid row
            self.nb.add(f, text=title)
            return f

        def _row(self, tab, key: str, label: str, kind: str = "str", hint: str = "") -> None:
            r = tab._r
            tab._r += 2
            val = getattr(self.cfg, key)
            ttk.Label(tab, text=label).grid(row=r, column=0, sticky="w", padx=(0, 10), pady=(4, 0))
            if kind == "bool":
                var: tk.Variable = tk.BooleanVar(value=bool(val))
                widget = ttk.Checkbutton(tab, variable=var)
            else:
                text = ", ".join(val) if kind == "list" else str(val)
                var = tk.StringVar(value=text)
                widget = ttk.Entry(tab, textvariable=var)
            widget.grid(row=r, column=1, sticky="ew", pady=(4, 0))
            if kind in ("dir", "file", "table"):
                ttk.Button(tab, text="Browse…",
                           command=lambda v=var, k=kind: self._browse(v, k)).grid(
                    row=r, column=2, padx=(6, 0), pady=(4, 0))
            if hint:
                ttk.Label(tab, text=hint, foreground="gray").grid(row=r + 1, column=1, sticky="w")
            self.vars[key] = var
            self.kinds[key] = kind

        def _browse(self, var, kind: str) -> None:
            if kind == "dir":
                p = filedialog.askdirectory()
            elif kind == "file":
                p = filedialog.askopenfilename()
            else:
                p = filedialog.asksaveasfilename(
                    defaultextension=".xlsx", confirmoverwrite=False,
                    filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
            if p:
                var.set(p)

        # ---- config <-> UI -------------------------------------------------
        def _gather(self) -> Config:
            data = asdict(self.cfg)
            for key, var in self.vars.items():
                kind, v = self.kinds[key], var.get()
                if kind == "int":
                    try:
                        data[key] = int(str(v).strip())
                    except ValueError:
                        raise ConfigError(f"“{key}” must be a whole number.")
                elif kind == "list":
                    data[key] = [x.strip() for x in str(v).split(",") if x.strip()]
                elif kind == "bool":
                    data[key] = bool(v)
                else:
                    data[key] = str(v).strip()
            cfg = Config(**data)
            validate_config(cfg)
            return cfg

        def _apply(self) -> Config | None:
            try:
                cfg = self._gather()
            except ConfigError as e:
                messagebox.showerror("Invalid settings", str(e))
                return None
            self.cfg = cfg
            save_config(cfg)
            return cfg

        # ---- actions -------------------------------------------------------
        def save(self) -> None:
            if self._apply():
                log.info("Settings saved to %s", Path(CONFIG_PATH).name)

        def preview(self) -> None:
            cfg = self._apply()
            if not cfg:
                return
            sample = {"date": datetime(2024, 9, 26).strftime(cfg.date_format),
                      "supplier": "АРГО", "invoice": "КА-16339", "material": "металл"}
            self.preview_lbl.config(text=build_filename(cfg, sample))

        def test_llm(self) -> None:
            cfg = self._apply()
            if not cfg:
                return

            def job() -> None:
                try:
                    msg, ok = test_ollama(cfg), True
                except Exception as e:
                    msg, ok = f"Cannot reach Ollama:\n{e}", False
                self.after(0, lambda: (messagebox.showinfo if ok else messagebox.showerror)("Ollama", msg))

            threading.Thread(target=job, daemon=True).start()

        def start(self) -> None:
            cfg = self._apply()
            if not cfg:
                return
            if not Path(cfg.watch_dir).is_dir():
                messagebox.showerror("Invalid settings", "Pick an existing watch folder first.")
                return
            self.worker = Worker(cfg)
            self.worker.start()
            self.btn_start.config(state="disabled")
            self.btn_once.config(state="disabled")
            self.btn_stop.config(state="normal")
            self.status.config(text="Running", foreground="green")

        def stop(self) -> None:
            if self.worker:
                self.worker.stop()
                self.worker = None
            self.btn_start.config(state="normal")
            self.btn_once.config(state="normal")
            self.btn_stop.config(state="disabled")
            self.status.config(text="Stopped", foreground="gray")

        def scan_once(self) -> None:
            cfg = self._apply()
            if cfg:
                w = Worker(cfg)
                threading.Thread(target=w.scan_once, daemon=True).start()
                self.nb.select(self.nb.index("end") - 1)

        def _drain_log(self) -> None:
            try:
                while True:
                    line = self.log_q.get_nowait()
                    self.log_box.config(state="normal")
                    self.log_box.insert("end", line + "\n")
                    self.log_box.see("end")
                    self.log_box.config(state="disabled")
            except queue.Empty:
                pass
            self.after(200, self._drain_log)

        def _close(self) -> None:
            if self.worker:
                self.worker.stop()
            self.destroy()

    App().mainloop()


# --------------------------------------------------------------------------- #
def main() -> None:
    global CONFIG_PATH
    ap = argparse.ArgumentParser(description="Invoice Renamer")
    ap.add_argument("--headless", action="store_true", help="run without the UI")
    ap.add_argument("--config", type=Path, default=CONFIG_PATH, help="path to config.json")
    args = ap.parse_args()
    CONFIG_PATH = args.config

    if not args.headless:
        run_ui()
        return

    setup_logging()
    cfg = load_config(CONFIG_PATH)
    try:
        validate_config(cfg)
    except ConfigError as e:
        sys.exit(f"Config error: {e}")
    worker = Worker(cfg)
    worker.start()
    try:
        while worker.is_alive():
            time.sleep(0.5)
    except KeyboardInterrupt:
        worker.stop()
        worker.join(timeout=10)


if __name__ == "__main__":
    main()
