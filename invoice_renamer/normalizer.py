"""Нормализация полей и построение имени файла."""
from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path

from .config import Config

log = logging.getLogger("invoice_renamer")


class ExtractionError(Exception):
    """Проблема с документом: файл уходит в _failed."""


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
    s = (value or "").strip()
    s = s.replace("№", " ").replace("#", " ")
    s = s.upper()
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
    hints = hints or {}

    date_src = (hints.get("date") or str(data.get("date", ""))).strip()
    invoice_src = (hints.get("invoice") or str(data.get("invoice_number", ""))).strip()

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
