from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime
from pathlib import Path


# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
def _default_app_dir() -> Path:
    env = os.environ.get("INVOICE_RENAMER_HOME")
    if env:
        return Path(env).expanduser().resolve()
    # config.py лежит в <app>/invoice_renamer/config.py → <app> — родитель пакета
    return Path(__file__).resolve().parent.parent


APP_DIR = _default_app_dir()
CONFIG_PATH = APP_DIR / "config.json"
LOG_PATH = APP_DIR / "invoice_renamer.log"


# --------------------------------------------------------------------------- #
# Domain constants
# --------------------------------------------------------------------------- #
REQUIRED_TOKEN = "Сч"
TEMPLATE_FIELDS = ("date", "supplier", "invoice", "material")

# Поля, которые обязательно должен вернуть LLM
EXTRACTION_FIELDS = ("date", "supplier", "invoice", "material")

# Поля, добавляемые приложением (не из LLM)
META_FIELDS = ("original_name", "new_name", "processed_at")

# Заголовки по умолчанию для «известных» ключей
KNOWN_HEADERS = {
    "date": "Дата",
    "supplier": "Поставщик",
    "invoice": "Номер счёта",
    "material": "Материал",
    "original_name": "Исходное имя файла",
    "new_name": "Новое имя файла",
    "processed_at": "Обработано",
}


DEFAULT_SYSTEM_PROMPT = """Ты — помощник бухгалтера. Тебе дают текст счёта на русском языке, \
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


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    # Folders / scheduling
    watch_dir: str = ""
    output_dir: str = ""              # пусто = переименовывать на месте
    recursive: bool = False
    poll_seconds: int = 10
    stable_seconds: int = 3
    move_failed: bool = True
    failed_dir_name: str = "_failed"

    # OCR
    tesseract_cmd: str = ""
    ocr_lang: str = "rus+eng"
    ocr_dpi: int = 300
    max_pages: int = 3
    tesseract_config: str = "--oem 1 --psm 6"
    use_embedded_text: bool = False

    # LLM
    ollama_url: str = "http://localhost:11434"
    model: str = "qwen2.5:7b"
    num_ctx: int = 8192
    timeout: int = 180
    max_chars: int = 12000
    keep_alive: str = "30m"
    system_prompt: str = DEFAULT_SYSTEM_PROMPT

    # Naming
    filename_template: str = "{date} {supplier} Сч № {invoice} {material}"
    date_format: str = "%d%m%y"
    invoice_cyrillic: bool = True
    max_name_len: int = 150

    # Table
    table_enabled: bool = True
    table_path: str = ""
    table_create_if_missing: bool = True
    # Список словарей {"key": str, "header": str}. Ключ может быть любым:
    # известным (date/supplier/invoice/material/original_name/new_name/processed_at)
    # или произвольным — тогда значение будет пустым.
    table_columns: list = field(default_factory=lambda: [
        {"key": k, "header": v} for k, v in KNOWN_HEADERS.items()
    ])

    # Meta
    setup_completed: bool = False


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
    # Гарантируем корректный тип колонок (на случай «старого» формата)
    cfg.table_columns = _coerce_columns(cfg.table_columns)
    return cfg


def save_config(cfg: Config, path: Path | None = None) -> None:
    path = path or CONFIG_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(cfg), ensure_ascii=False, indent=2), "utf-8")


class ConfigError(Exception):
    pass


# --------------------------------------------------------------------------- #
# Columns helpers
# --------------------------------------------------------------------------- #
def _coerce_columns(value) -> list[dict]:
    """Приводим table_columns к списку {"key","header"}.

    Поддерживаем: список dict, список строк (ключи), строку «a, b, c».
    """
    if not value:
        return []
    if isinstance(value, str):
        value = [x.strip() for x in value.split(",") if x.strip()]
    out: list[dict] = []
    for item in value:
        if isinstance(item, dict):
            key = str(item.get("key", "")).strip()
            header = str(item.get("header", "")).strip() or KNOWN_HEADERS.get(key, key)
        else:
            key = str(item).strip()
            header = KNOWN_HEADERS.get(key, key)
        if key:
            out.append({"key": key, "header": header})
    return out


def parse_columns(text: str) -> list[dict]:
    """Разбор текстового представления колонок: 'key | Заголовок' на строку."""
    return _coerce_columns([
        line.strip() for line in text.splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ])


def format_columns(cols: list[dict]) -> str:
    return "\n".join(f"{c['key']} | {c['header']}" for c in cols)


# --------------------------------------------------------------------------- #
# Validation
# --------------------------------------------------------------------------- #
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

    cols = _coerce_columns(cfg.table_columns)
    if not cols:
        raise ConfigError("Table columns must not be empty.")
    for col in cols:
        if not col.get("key") or not col.get("header"):
            raise ConfigError("Each table column must have both key and header.")

    for name in ("poll_seconds", "ocr_dpi", "max_pages", "num_ctx", "timeout", "max_chars"):
        if int(getattr(cfg, name)) <= 0:
            raise ConfigError(f"“{name}” must be a positive number.")
    if cfg.table_enabled and not cfg.table_path.strip():
        raise ConfigError("Table logging is enabled but the table path is empty.")
