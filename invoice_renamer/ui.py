"""Основное окно приложения."""
from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from .config import (
    CONFIG_PATH,
    DEFAULT_SYSTEM_PROMPT,
    Config,
    ConfigError,
    format_columns,
    load_config,
    parse_columns,
    save_config,
    validate_config,
)
from .llm import list_ollama_models, test_ollama
from .logging_setup import setup_logging
from .normalizer import build_filename
from .worker import Worker

log = logging.getLogger("invoice_renamer")


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Invoice Renamer")
        self.geometry("920x720")
        self.minsize(760, 560)
        self.cfg = load_config()
        self.vars: dict[str, tk.Variable] = {}
        self.kinds: dict[str, str] = {}
        self.text_widgets: dict[str, tk.Text] = {}
        self.worker: Worker | None = None
        self.log_q: queue.Queue = queue.Queue()
        setup_logging(self.log_q)
        self._build()
        self.after(200, self._drain_log)
        self.protocol("WM_DELETE_WINDOW", self._close)

    # ------------------------------------------------------------------ #
    # Layout
    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        # ---- toolbar -----------------------------------------------------
        bar = ttk.Frame(self, padding=(10, 8))
        bar.pack(fill="x")

        self.btn_start = ttk.Button(bar, text="▶ Start", command=self.start)
        self.btn_stop = ttk.Button(bar, text="■ Stop", command=self.stop, state="disabled")
        self.btn_once = ttk.Button(bar, text="Scan once", command=self.scan_once)
        for b in (self.btn_start, self.btn_stop, self.btn_once):
            b.pack(side="left", padx=(0, 6))

        ttk.Button(bar, text="Save settings", command=self.save).pack(
            side="left", padx=(12, 6))
        ttk.Button(bar, text="Setup wizard…", command=self.open_wizard).pack(
            side="left")
        ttk.Button(bar, text="Test Ollama", command=self.test_llm).pack(
            side="left", padx=(6, 0))

        self.status = ttk.Label(bar, text="Stopped", foreground="gray")
        self.status.pack(side="right")

        # ---- notebook ----------------------------------------------------
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self._build_folders_tab()
        self._build_ocr_tab()
        self._build_llm_tab()
        self._build_prompt_tab()
        self._build_naming_tab()
        self._build_table_tab()
        self._build_log_tab()

    # ---- Folders ---------------------------------------------------------
    def _build_folders_tab(self) -> None:
        t = self._tab("Folders")
        self._row(t, "watch_dir", "Watch folder", "dir")
        self._row(t, "output_dir", "Output folder", "dir",
                  "Empty = rename in place")
        self._row(t, "recursive", "Include subfolders", "bool")
        self._row(t, "poll_seconds", "Scan interval (sec)", "int")
        self._row(t, "stable_seconds", "File must be idle for (sec)", "int",
                  "Avoids reading files that are still being copied")
        self._row(t, "move_failed", "Move unrecognised files", "bool")
        self._row(t, "failed_dir_name", "…into subfolder", "str")

    # ---- OCR -------------------------------------------------------------
    def _build_ocr_tab(self) -> None:
        t = self._tab("OCR")
        self._row(t, "tesseract_cmd", "Tesseract executable", "file",
                  "Empty = found via PATH")
        self._row(t, "ocr_lang", "Languages", "str",
                  "e.g. rus+eng (needs the “rus” language data)")
        self._row(t, "ocr_dpi", "Render DPI", "int",
                  "300 is a good default; 200 is faster")
        self._row(t, "max_pages", "Max pages to read", "int")
        self._row(t, "tesseract_config", "Tesseract options", "str")
        self._row(t, "use_embedded_text", "Use PDF text layer if present", "bool",
                  "Skips OCR for digital (non-scanned) PDFs")

    # ---- LLM (Ollama) ----------------------------------------------------
    def _build_llm_tab(self) -> None:
        t = self._tab("LLM (Ollama)")
        self._row(t, "ollama_url", "Ollama URL", "str")

        # Модель: combobox + кнопка обновления (вручную, т.к. нужен свой layout)
        r = t._r
        t._r += 2
        ttk.Label(t, text="Model").grid(
            row=r, column=0, sticky="w", padx=(0, 10), pady=(4, 0))
        self.model_var = tk.StringVar(value=self.cfg.model)
        self.model_combo = ttk.Combobox(t, textvariable=self.model_var)
        self.model_combo.grid(row=r, column=1, sticky="ew", pady=(4, 0))
        ttk.Button(t, text="↻", width=3, command=self.refresh_models).grid(
            row=r, column=2, padx=(6, 0), pady=(4, 0))
        ttk.Label(
            t, text="Нажмите ↻, чтобы получить список моделей из Ollama",
            foreground="gray",
        ).grid(row=r + 1, column=1, sticky="w")
        self.vars["model"] = self.model_var
        self.kinds["model"] = "str"

        self._row(t, "num_ctx", "Context size (tokens)", "int")
        self._row(t, "timeout", "Timeout (sec)", "int")
        self._row(t, "max_chars", "Max characters sent", "int")
        self._row(t, "keep_alive", "Keep model in RAM for", "str",
                  "e.g. 30m, 1h — avoids reloading on every file")

    # ---- Prompt ----------------------------------------------------------
    def _build_prompt_tab(self) -> None:
        t = self._tab("Prompt")
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)

        ttk.Label(
            t, wraplength=840, justify="left", foreground="gray",
            text=("Системный промпт, отправляемый LLM. Правьте осторожно: "
                  "модель должна вернуть JSON с полями date, supplier, "
                  "invoice_number, material."),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        self.prompt_text = scrolledtext.ScrolledText(t, wrap="word", height=20)
        self.prompt_text.grid(row=1, column=0, sticky="nsew")
        self.prompt_text.insert("1.0", self.cfg.system_prompt)
        self.text_widgets["system_prompt"] = self.prompt_text

        row = ttk.Frame(t)
        row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(row, text="Сбросить к стандартному",
                   command=self.reset_prompt).pack(side="left")

    # ---- Naming ----------------------------------------------------------
    def _build_naming_tab(self) -> None:
        t = self._tab("Naming")
        self._row(t, "filename_template", "Filename template", "str",
                  "Placeholders: {date} {supplier} {invoice} {material}. "
                  "“Сч” is required.")
        self._row(t, "date_format", "Date format", "str",
                  "strftime: %d%m%y → 260924")
        self._row(t, "invoice_cyrillic", "Latin look-alikes → Cyrillic", "bool",
                  "KA-16339 → КА-16339")
        self._row(t, "max_name_len", "Max name length", "int")

        pr = ttk.Frame(t)
        pr.grid(row=t._r, column=0, columnspan=3, sticky="ew", pady=10)
        ttk.Button(pr, text="Preview", command=self.preview).pack(side="left")
        self.preview_lbl = ttk.Label(pr, text="")
        self.preview_lbl.pack(side="left", padx=8)

    # ---- Table -----------------------------------------------------------
    def _build_table_tab(self) -> None:
        t = self._tab("Table")
        t.columnconfigure(0, weight=1)

        self._row(t, "table_enabled", "Write results to a table", "bool")
        self._row(t, "table_path", "Table file (.xlsx / .csv)", "table")
        self._row(t, "table_create_if_missing",
                  "Create file if it does not exist", "bool")

        info_row = t._r
        t._r += 1
        ttk.Label(
            t, wraplength=840, justify="left",
            text=("Колонки таблицы: одна строка на колонку, формат "
                  "«key | Заголовок». Известные ключи: date, supplier, "
                  "invoice, material, original_name, new_name, processed_at. "
                  "Можно добавлять свои — значения будут пустыми."),
        ).grid(row=info_row, column=0, columnspan=3, sticky="w", pady=(8, 4))

        text_row = t._r
        t.rowconfigure(text_row, weight=1)
        self.columns_text = scrolledtext.ScrolledText(t, wrap="none", height=10)
        self.columns_text.grid(row=text_row, column=0, columnspan=3, sticky="nsew")
        self.columns_text.insert("1.0", format_columns(self.cfg.table_columns))
        self.text_widgets["table_columns"] = self.columns_text

    # ---- Log -------------------------------------------------------------
    def _build_log_tab(self) -> None:
        t = ttk.Frame(self.nb, padding=4)
        t.columnconfigure(0, weight=1)
        t.rowconfigure(0, weight=1)
        self.nb.add(t, text="Log")
        self.log_box = scrolledtext.ScrolledText(
            t, height=10, state="disabled", wrap="word")
        self.log_box.grid(row=0, column=0, sticky="nsew")

    # ------------------------------------------------------------------ #
    # Helpers for building rows
    # ------------------------------------------------------------------ #
    def _tab(self, title: str) -> ttk.Frame:
        f = ttk.Frame(self.nb, padding=12)
        f.columnconfigure(1, weight=1)
        f._r = 0  # type: ignore[attr-defined]
        self.nb.add(f, text=title)
        return f

    def _row(self, tab, key: str, label: str, kind: str = "str",
             hint: str = "") -> None:
        r = tab._r
        tab._r += 2  # type: ignore[attr-defined]

        val = getattr(self.cfg, key)
        ttk.Label(tab, text=label).grid(
            row=r, column=0, sticky="w", padx=(0, 10), pady=(4, 0))

        if kind == "bool":
            var: tk.Variable = tk.BooleanVar(value=bool(val))
            widget = ttk.Checkbutton(tab, variable=var)
        else:
            var = tk.StringVar(value=str(val))
            widget = ttk.Entry(tab, textvariable=var)
        widget.grid(row=r, column=1, sticky="ew", pady=(4, 0))

        if kind in ("dir", "file", "table"):
            ttk.Button(
                tab, text="Browse…",
                command=lambda v=var, k=kind: self._browse(v, k),
            ).grid(row=r, column=2, padx=(6, 0), pady=(4, 0))

        if hint:
            ttk.Label(tab, text=hint, foreground="gray").grid(
                row=r + 1, column=1, sticky="w")

        self.vars[key] = var
        self.kinds[key] = kind

    def _browse(self, var, kind: str) -> None:
        if kind == "dir":
            p = filedialog.askdirectory()
        elif kind == "file":
            p = filedialog.askopenfilename()
        else:
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                confirmoverwrite=False,
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            )
        if p:
            var.set(p)

    # ------------------------------------------------------------------ #
    # Config <-> UI
    # ------------------------------------------------------------------ #
    def _gather(self) -> Config:
        data = asdict(self.cfg)

        for key, var in self.vars.items():
            kind, v = self.kinds[key], var.get()
            if kind == "int":
                try:
                    data[key] = int(str(v).strip())
                except ValueError:
                    raise ConfigError(f"“{key}” must be a whole number.")
            elif kind == "bool":
                data[key] = bool(v)
            else:
                data[key] = str(v).strip()

        if "system_prompt" in self.text_widgets:
            data["system_prompt"] = (
                self.text_widgets["system_prompt"].get("1.0", "end").rstrip()
            )
        if "table_columns" in self.text_widgets:
            data["table_columns"] = parse_columns(
                self.text_widgets["table_columns"].get("1.0", "end")
            )

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

    # ------------------------------------------------------------------ #
    # Actions
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        if self._apply():
            log.info("Settings saved to %s", Path(CONFIG_PATH).name)

    def preview(self) -> None:
        cfg = self._apply()
        if not cfg:
            return
        sample = {
            "date": datetime(2024, 9, 26).strftime(cfg.date_format),
            "supplier": "АРГО",
            "invoice": "КА-16339",
            "material": "металл",
        }
        self.preview_lbl.config(text=build_filename(cfg, sample))

    def reset_prompt(self) -> None:
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", DEFAULT_SYSTEM_PROMPT)

    def refresh_models(self) -> None:
        try:
            url = str(self.vars["ollama_url"].get()).strip() or "http://localhost:11434"
        except KeyError:
            url = "http://localhost:11434"

        def job() -> None:
            try:
                models = list_ollama_models(url, timeout=5)
            except Exception as e:
                self.after(0, lambda err=e: messagebox.showerror("Ollama", str(err)))
                return

            def apply_models() -> None:
                self.model_combo["values"] = models
                if models and self.model_var.get() not in models:
                    self.model_var.set(models[0])
                log.info("Ollama: найдено моделей — %d (%s)",
                         len(models), ", ".join(models))

            self.after(0, apply_models)

        threading.Thread(target=job, daemon=True).start()

    def test_llm(self) -> None:
        cfg = self._apply()
        if not cfg:
            return

        def job() -> None:
            try:
                msg, ok = test_ollama(cfg), True
            except Exception as e:
                msg, ok = f"Cannot reach Ollama:\n{e}", False

            def show() -> None:
                fn = messagebox.showinfo if ok else messagebox.showerror
                fn("Ollama", msg)

            self.after(0, show)

        threading.Thread(target=job, daemon=True).start()

    def open_wizard(self) -> None:
        try:
            from .setup_wizard import SetupWizard
        except ImportError as e:
            messagebox.showerror("Setup wizard", f"Cannot start wizard: {e}")
            return
        wiz = SetupWizard()
        wiz.mainloop()
        if getattr(wiz, "saved", False):
            self._reload_from_config()

    def _reload_from_config(self) -> None:
        """Перечитать config.json и обновить поля UI после мастера."""
        self.cfg = load_config()
        for key, var in self.vars.items():
            if hasattr(self.cfg, key):
                try:
                    var.set(getattr(self.cfg, key))
                except Exception:
                    pass
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", self.cfg.system_prompt)
        self.columns_text.delete("1.0", "end")
        self.columns_text.insert("1.0", format_columns(self.cfg.table_columns))
        self.model_var.set(self.cfg.model)

    def start(self) -> None:
        cfg = self._apply()
        if not cfg:
            return
        if not Path(cfg.watch_dir).is_dir():
            messagebox.showerror("Invalid settings",
                                 "Pick an existing watch folder first.")
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
        if not cfg:
            return
        w = Worker(cfg)
        threading.Thread(target=w.scan_once, daemon=True).start()
        self.nb.select(self.nb.index("end") - 1)

    # ------------------------------------------------------------------ #
    # Log draining
    # ------------------------------------------------------------------ #
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


def run_ui() -> None:
    App().mainloop()
