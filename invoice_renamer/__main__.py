"""CLI-точка входа."""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from . import config as cfg_module
from .config import ConfigError, load_config, validate_config


def main() -> None:
    ap = argparse.ArgumentParser(description="Invoice Renamer")
    ap.add_argument("--headless", action="store_true", help="Run without the UI")
    ap.add_argument("--config", type=Path, default=None, help="Path to config.json")
    ap.add_argument("--setup", action="store_true", help="Force the setup wizard")
    args = ap.parse_args()

    if args.config:
        cfg_module.CONFIG_PATH = args.config

    if args.headless:
        _run_headless()
    else:
        _run_gui(force_setup=args.setup)


# --------------------------------------------------------------------------- #
def _run_headless() -> None:
    try:
        from .logging_setup import setup_logging
        from .worker import Worker  # noqa: F401
    except ImportError as e:
        sys.exit(
            f"Missing dependency: {e}\n"
            "Install the requirements and try again. "
            "Run without --headless for the guided setup."
        )
    setup_logging()
    cfg = load_config(cfg_module.CONFIG_PATH)
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


def _run_gui(force_setup: bool = False) -> None:
    try:
        import tkinter  # noqa: F401
    except ImportError:
        sys.exit("Tkinter is not installed. Use --headless or install python3-tk.")

    cfg = load_config(cfg_module.CONFIG_PATH)
    if force_setup or not cfg.setup_completed:
        try:
            from .setup_wizard import SetupWizard
        except ImportError as e:
            sys.exit(f"Cannot start the setup wizard: {e}")

        wiz = SetupWizard()
        wiz.mainloop()
        if not wiz.saved:
            # пользователь закрыл мастер без сохранения — выходим
            return

    try:
        from .ui import run_ui
    except ImportError as e:
        # Показываем хоть какое-то окно
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Missing dependency",
                f"Не удалось запустить приложение: {e}\n\n"
                "Запустите мастер настройки:\n"
                "    python invoice_renamer.py --setup",
            )
            root.destroy()
        except Exception:
            pass
        sys.exit(f"Missing dependency: {e}")

    run_ui()


if __name__ == "__main__":
    main()
