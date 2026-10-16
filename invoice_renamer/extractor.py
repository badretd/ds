"""PDF → текст (встроенный слой или OCR через Tesseract)."""
from __future__ import annotations

import logging

from .config import Config
from .normalizer import ExtractionError

log = logging.getLogger("invoice_renamer")


class ServiceUnavailable(Exception):
    """Проблема окружения (Ollama/Tesseract недоступны) — файл повторим позже."""


def extract_text(cfg: Config, path) -> str:
    # Импорты внутри функции, чтобы мастер настройки мог работать без них
    try:
        import pymupdf as fitz           # PyMuPDF >= 1.24
    except ImportError:
        import fitz                      # PyMuPDF < 1.24
    import pytesseract
    from PIL import Image
    from pytesseract import TesseractNotFoundError

    parts: list[str] = []
    try:
        doc = fitz.open(path)
    except Exception as e:
        raise ExtractionError(f"Cannot open PDF: {e}")
    with doc:
        if doc.needs_pass:
            raise ExtractionError("PDF is password protected")
        for i, page in enumerate(doc):
            if i >= cfg.max_pages:
                break
            text = page.get_text().strip() if cfg.use_embedded_text else ""
            if len(text) < 40:
                pix = page.get_pixmap(dpi=cfg.ocr_dpi)
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
