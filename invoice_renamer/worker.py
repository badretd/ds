"""Фоновый поток обработки PDF-счетов."""
from __future__ import annotations

import logging
import re
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path

from .config import Config, ConfigError, validate_config
from .extractor import ServiceUnavailable, extract_text
from .hints import extract_hints
from .llm import ask_llm
from .normalizer import (
    ExtractionError,
    build_filename,
    normalize_fields,
    unique_target,
)
from .table import TableError, table_append, table_preflight

log = logging.getLogger("invoice_renamer")

ALREADY_NAMED = re.compile(r"^\d{4,8} .+ Сч\s*№", re.IGNORECASE)


class Worker(threading.Thread):
    def __init__(self, cfg: Config):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.stop_event = threading.Event()
        self.seen: set[tuple] = set()
        if cfg.tesseract_cmd.strip():
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = cfg.tesseract_cmd.strip()

    def run(self) -> None:
        log.info("Started. Watching: %s", self.cfg.watch_dir)
        try:
            import pytesseract
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

    def process(self, pdf: Path) -> Path:
        cfg = self.cfg
        log.info("Processing %s", pdf.name)
        text = extract_text(cfg, pdf)

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
