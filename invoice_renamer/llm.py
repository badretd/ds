"""Клиент Ollama: извлечение полей, проверка соединения, список моделей."""
from __future__ import annotations

import json
import logging

from .config import (
    Config,
    build_system_prompt,
    llm_fields,
)
from .extractor import ServiceUnavailable
from .normalizer import ExtractionError

log = logging.getLogger("invoice_renamer")


def build_schema(cfg: Config) -> dict:
    """Динамическая JSON-схема: базовые поля + кастомные (не «null»)."""
    props = {
        "date": {"type": "string"},
        "supplier": {"type": "string"},
        "invoice_number": {"type": "string"},
        "material": {"type": "string"},
    }
    for c in llm_fields(cfg):
        # Кастомные поля всегда запрашиваем строкой — приводить к числу/дате
        # будет normalizer. Так модель не спотыкается о строгие типы.
        props[c["key"]] = {"type": "string"}
    return {
        "type": "object",
        "properties": props,
        "required": ["date", "supplier", "invoice_number", "material"],
    }


def ask_llm(cfg: Config, text: str) -> dict:
    import requests

    payload = {
        "model": cfg.model,
        "stream": False,
        "format": build_schema(cfg),
        "keep_alive": cfg.keep_alive,
        "options": {
            "temperature": 0,
            "num_ctx": cfg.num_ctx,
            "num_predict": 512,
        },
        "messages": [
            {"role": "system", "content": build_system_prompt(cfg)},
            {"role": "user", "content": 'Текст документа:\n"""\n' + text[: cfg.max_chars] + '\n"""'},
        ],
    }
    url = cfg.ollama_url.rstrip("/") + "/api/chat"
    try:
        r = requests.post(url, json=payload, timeout=cfg.timeout)
    except requests.ConnectionError as e:
        raise ServiceUnavailable(f"Ollama is not reachable at {cfg.ollama_url}: {e}")
    except requests.Timeout:
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


def list_ollama_models(url: str, timeout: int = 10) -> list[str]:
    """Возвращает список имён моделей, установленных в Ollama."""
    import requests
    r = requests.get(url.rstrip("/") + "/api/tags", timeout=timeout)
    r.raise_for_status()
    return [m.get("name", "") for m in r.json().get("models", []) if m.get("name")]


def test_ollama(cfg: Config) -> str:
    """Человекочитаемый статус; бросает исключение при ошибке."""
    names = list_ollama_models(cfg.ollama_url, timeout=10)
    wanted = cfg.model if ":" in cfg.model else cfg.model + ":latest"
    if wanted in names or cfg.model in names:
        return f"Ollama OK. Model “{cfg.model}” is available."
    return (
        f"Ollama is running, but model “{cfg.model}” is not installed.\n"
        f"Installed: {', '.join(names) or 'none'}\nRun: ollama pull {cfg.model}"
    )