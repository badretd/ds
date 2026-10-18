from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime
from pathlib import Path
from string import Formatter


# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
def _default_app_dir() -> Path:
    env = os.environ.get("INVOICE_RENAMER_HOME")
    if env:
        return Path(env).expanduser().resolve()
    # PyInstaller: __file__ указывает во временную папку _MEIPASS,
    # поэтому берём каталог, где лежит сам EXE.
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Обычный запуск: config.py лежит в <app>/invoice_renamer/config.py
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
    "id": "ID",
    "date": "Дата",
    "supplier": "Поставщик",
    "invoice": "Номер счёта",
    "material": "Материал",
    "original_name": "Исходное имя файла",
    "new_name": "Новое имя файла",
    "processed_at": "Обработано",
}
KNOWN_KEYS = set(KNOWN_HEADERS.keys())

# Допустимые типы колонок.
# «null» — служебная колонка: присутствует в таблице, но не заполняется ни LLM,
# ни программой (кроме случая, когда пользователь сам впишет туда значение).
FIELD_TYPES = ("string", "integer", "number", "date", "datetime", "boolean", "null")
DEFAULT_FIELD_TYPES = {
    "id": "integer",
    "date": "string",
    "supplier": "string",
    "invoice": "string",
    "material": "string",
    "original_name": "string",
    "new_name": "string",
    "processed_at": "datetime",
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
    table_auto_id: bool = True
    # Список словарей {"key": str, "header": str, "type": str, "description": str}.
    # Ключ может быть известным (date/supplier/invoice/material/... ) или
    # произвольным — тогда это кастомное поле, которое уходит в LLM.
    # Тип «null» — колонка есть в таблице, но не заполняется автоматически.
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
    # Миграция старых форматов + гарантия корректного типа колонок
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
def normalize_column(col) -> dict:
    """Приводим одну колонку к канону: {key, header, type, description}."""
    if isinstance(col, str):
        if "|" in col:
            parts = [p.strip() for p in col.split("|")]
            key = parts[0] if parts else ""
            header = parts[1] if len(parts) > 1 else ""
            typ = parts[2] if len(parts) > 2 and parts[2] else None
            desc = parts[3] if len(parts) > 3 else ""
            col = {"key": key, "header": header, "description": desc}
            if typ:
                col["type"] = typ
        else:
            col = {"key": col.strip()}

    if not isinstance(col, dict):
        return {}

    key = str(col.get("key", "")).strip()
    header = str(col.get("header", "")).strip()

    # Миграция старого формата: «key | header» хранился прямо в ключе
    if "|" in key:
        k, _, h = key.partition("|")
        key = k.strip()
        if not header or "|" in header:
            header = h.strip()

    if not key:
        return {}
    if not header:
        header = KNOWN_HEADERS.get(key, key)

    typ = col.get("type") or DEFAULT_FIELD_TYPES.get(key, "string")
    if typ not in FIELD_TYPES:
        typ = "string"

    return {
        "key": key,
        "header": header,
        "type": typ,
        "description": str(col.get("description", "")).strip(),
    }


def _coerce_columns(value) -> list[dict]:
    """Приводим table_columns к списку канонических dict."""
    if not value:
        return []
    if isinstance(value, str):
        value = [x.strip() for x in value.split(",") if x.strip()]
    out: list[dict] = []
    for item in value:
        col = normalize_column(item)
        if col:
            out.append(col)
    return out


def parse_columns(text: str) -> list[dict]:
    """Разбор текстового представления колонок.

    Строки: ``key | Заголовок`` или ``key | Заголовок | тип | описание``.
    Пустые строки и строки, начинающиеся с ``#``, пропускаются.
    """
    out: list[dict] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.lstrip().startswith("#"):
            continue
        parts = [p.strip() for p in line.split("|")]
        while len(parts) < 4:
            parts.append("")
        col = normalize_column({
            "key": parts[0],
            "header": parts[1],
            "type": parts[2] or None,
            "description": parts[3],
        })
        if col:
            out.append(col)
    return out


def format_columns(cols: list[dict]) -> str:
    """Обратное к parse_columns: одна колонка — одна строка."""
    lines: list[str] = []
    for c in cols:
        parts = [c["key"], c["header"]]
        typ = c.get("type", "")
        desc = c.get("description", "")
        show_typ = bool(typ) and typ != DEFAULT_FIELD_TYPES.get(c["key"], "string")
        if show_typ or desc:
            parts.append(typ if show_typ else "")
        if desc:
            parts.append(desc)
        lines.append(" | ".join(parts))
    return "\n".join(lines)


def custom_fields(cfg: "Config") -> list[dict]:
    """Все колонки, которых нет в KNOWN_HEADERS."""
    return [c for c in cfg.table_columns if c["key"] not in KNOWN_KEYS]


def llm_fields(cfg: "Config") -> list[dict]:
    """Кастомные поля, которые должен заполнять LLM.

    Исключаем тип «null» — это служебные колонки, LLM их не касается.
    """
    return [c for c in custom_fields(cfg) if c.get("type") != "null"]


def build_system_prompt(cfg: "Config") -> str:
    """Дополняем системный промпт инструкциями для кастомных полей."""
    prompt = cfg.system_prompt
    custom = llm_fields(cfg)
    if not custom:
        return prompt
    lines = [
        "",
        "",
        "ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ.",
        "Извлеки их из документа и добавь в JSON-ответ.",
        "Если поле не найдено — верни пустую строку.",
    ]
    for c in custom:
        typ = c.get("type", "string")
        desc = c.get("description") or c.get("header") or c["key"]
        lines.append(f"- {c['key']} (тип: {typ}): {desc}")
    return prompt + "\n".join(lines)


# --------------------------------------------------------------------------- #
# Validation
# --------------------------------------------------------------------------- #
def _template_placeholders(template: str) -> set[str]:
    names: set[str] = set()
    for _, fieldname, _, _ in Formatter().parse(template):
        if fieldname:
            names.add(fieldname.split(".", 1)[0].split("[", 1)[0])
    return names


def validate_config(cfg: Config) -> None:
    if REQUIRED_TOKEN not in cfg.filename_template:
        raise ConfigError(f"The filename template must contain “{REQUIRED_TOKEN}”.")

    cols = _coerce_columns(cfg.table_columns)
    custom_keys = {c["key"] for c in cols} - KNOWN_KEYS
    allowed = set(TEMPLATE_FIELDS) | custom_keys

    placeholders = _template_placeholders(cfg.filename_template)

    if "id" in placeholders:
        raise ConfigError(
            "«{id}» нельзя использовать в шаблоне имени файла: ID присваивается "
            "при записи строки в таблицу — уже после переименования файла."
        )
    unknown = placeholders - allowed
    if unknown:
        raise ConfigError(
            "Unknown placeholders in filename template: "
            + ", ".join(sorted(unknown))
            + ". Allowed: "
            + ", ".join(sorted(allowed))
            + "."
        )

    try:
        cfg.filename_template.format(**{k: "x" for k in allowed})
    except (KeyError, IndexError, ValueError) as e:
        raise ConfigError(f"Bad filename template ({e!r}).")

    try:
        datetime(2024, 9, 26).strftime(cfg.date_format)
    except ValueError as e:
        raise ConfigError(f"Bad date format: {e}")

    if not cols:
        raise ConfigError("Table columns must not be empty.")
    for col in cols:
        if not col.get("key") or not col.get("header"):
            raise ConfigError("Each table column must have both key and header.")
        if col.get("type") not in FIELD_TYPES:
            raise ConfigError(
                f"Unknown column type {col.get('type')!r} for "
                f"column {col['key']!r}. Allowed: {', '.join(FIELD_TYPES)}."
            )

    for name in ("poll_seconds", "ocr_dpi", "max_pages", "num_ctx", "timeout", "max_chars"):
        if int(getattr(cfg, name)) <= 0:
            raise ConfigError(f"“{name}” must be a positive number.")
    if cfg.table_enabled and not cfg.table_path.strip():
        raise ConfigError("Table logging is enabled but the table path is empty.")