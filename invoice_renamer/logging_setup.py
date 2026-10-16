"""Единая настройка логгера."""
from __future__ import annotations

import logging
import logging.handlers
import queue
import sys

from .config import LOG_PATH

log = logging.getLogger("invoice_renamer")


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
