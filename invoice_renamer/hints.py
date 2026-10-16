"""Детерминированные подсказки из шапки счёта (regex)."""
from __future__ import annotations

import re

MONTHS_RU = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5,
    "июня": 6, "июля": 7, "августа": 8, "сентября": 9, "октября": 10,
    "ноября": 11, "декабря": 12,
}

INVOICE_HINT_RE = re.compile(
    r'Сч[её]т\s+на\s+оплату\s*[№N#:]?\s*'
    r'([A-ZА-ЯЁа-яёa-z]{0,6}[\s\-–—/_]?\d{1,15})',
    re.IGNORECASE,
)

DATE_NUM_HINT_RE = re.compile(
    r'Сч[её]т\s+на\s+оплату[^\n]{0,150}?'
    r'от\s+(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})',
    re.IGNORECASE,
)

DATE_WORD_HINT_RE = re.compile(
    r'от\s+(\d{1,2})\s+(' + '|'.join(MONTHS_RU) + r')\s+(\d{4})',
    re.IGNORECASE,
)


def extract_hints(text: str) -> dict:
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
