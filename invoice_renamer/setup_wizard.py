"""Мастер первичной настройки."""
from __future__ import annotations

import shutil
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext

from .config import (
    CONFIG_PATH,
    Config,
    load_config,
    save_config,
    validate_config,
    ConfigError,
)


# --------------------------------------------------------------------------- #
# Dependency check
# --------------------------------------------------------------------------- #
def _check_dependencies() -> list[dict]:
    checks: list[dict] = []

    ok = sys.version_info >= (3, 10)
    checks.append({
        "name": "Python 3.10+",
        "ok": ok,
        "detail": f"запущено: {sys.version.split()[0]}",
        "hint": "" if ok else "Обновите Python: https://www.python.org/downloads/",
    })

    packages = [
        ("fitz", "PyMuPDF", "рендеринг PDF-страниц в изображения"),
        ("pytesseract", "pytesseract", "обёртка вокруг Tesseract OCR"),
        ("PIL", "Pillow", "работа с изображениями"),
        ("requests", "requests", "HTTP-запросы к Ollama"),
        ("openpyxl", "openpyxl", "запись в Excel .xlsx"),
    ]
    for mod, pip_name, purpose in packages:
        try:
            __import__(mod)
            ok = True
        except ImportError:
            ok = False
        checks.append({
            "name": f"Python-пакет: {pip_name}",
            "ok": ok,
            "detail": purpose,
            "hint": "" if ok else f"pip install {pip_name}",
        })

    tesseract = shutil.which("tesseract")
    checks.append({
        "name": "Tesseract OCR (бинарник)",
        "ok": bool(tesseract),
        "detail": tesseract or "не найден в PATH",
        "hint": "" if tesseract else (
            "Windows: скачайте инсталлятор https://github.com/UB-Mannheim/tesseract/wiki\n"
            "    при установке отметьте галочку «Russian».\n"
            "Linux:   sudo apt install tesseract-ocr tesseract-ocr-rus\n"
            "macOS:   brew install tesseract tesseract-lang"
        ),
    })

    if tesseract:
        has_rus: bool | None
        try:
            import pytesseract
            langs = pytesseract.get_languages(config="")
            has_rus = "rus" in langs
        except Exception:
            has_rus = None
        if has_rus is not None:
            checks.append({
                "name": "Язык «rus» для Tesseract",
                "ok": has_rus,
                "detail": "установлен" if has_rus else "не найден",
                "hint": "" if has_rus else (
                    "Linux: sudo apt install tesseract-ocr-rus\n"
                    "Windows: переустановите Tesseract с галочкой Russian."
                ),
            })

    return checks


# --------------------------------------------------------------------------- #
# Wizard base
# --------------------------------------------------------------------------- #
class WizardPage(ttk.Frame):
    title = ""
    subtitle = ""

    def __init__(self, wizard: "SetupWizard"):
        super().__init__(wizard.body, padding=6)
        self.wizard = wizard
        self.cfg: Config = wizard.cfg

    def on_enter(self) -> None: ...
    def on_leave(self) -> bool:
        return True
    def apply(self) -> None: ...


class WelcomePage(WizardPage):
    title = "Добро пожаловать в Invoice Renamer"
    subtitle = "Первичная настройка займёт пару минут."

    def __init__(self, wizard: "SetupWizard"):
        super().__init__(wizard)
        text = (
            "Мастер поможет:\n\n"
            "  1. Проверить, установлены ли нужные зависимости:\n"
            "     Python-пакеты, Tesseract OCR (с русским языком), Ollama.\n"
            "  2. Подключиться к локальному серверу Ollama и выбрать модель.\n"
            "  3. Указать папку, в которой будут обрабатываться счета.\n"
            "  4. Настроить таблицу (Excel или CSV) для учёта.\n\n"
            "Все настройки сохраняются в файл:\n\n"
            f"      {CONFIG_PATH}\n\n"
            "Мастер можно вызвать заново из основного окна или командой:\n\n"
            "      python invoice_renamer.py --setup"
        )
        ttk.Label(self, text=text, justify="left", wraplength=680).pack(anchor="w", pady=8)


class DependenciesPage(WizardPage):
    title = "Проверка зависимостей"
    subtitle = "Установите всё, что отмечено красным крестиком."

    def __init__(self, wizard: "SetupWizard"):
        super().__init__(wizard)
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="Результаты проверки:").pack(side="left")
        ttk.Button(top, text="Проверить снова", command=self.refresh).pack(side="right")

        self.txt = scrolledtext.ScrolledText(self, wrap="word", height=20, state="disabled")
        self.txt.pack(fill="both", expand=True, pady=(6, 0))
        self.txt.tag_configure("ok", foreground="#1a7f1a")
        self.txt.tag_configure("fail", foreground="#b00020")
        self.txt.tag_configure("head", font=("TkDefaultFont", 10, "bold"))

    def on_enter(self) -> None:
        self.refresh()

    def refresh(self) -> None:
        checks = _check_dependencies()
        n_ok = sum(1 for c in checks if c["ok"])
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert(
            "end",
            f"Проверено: {n_ok}/{len(checks)} компонентов в порядке.\n\n",
            "head",
        )
        for c in checks:
            mark = "✓" if c["ok"] else "✗"
            tag = "ok" if c["ok"] else "fail"
            self.txt.insert("end", f"{mark}  ", tag)
            self.txt.insert("end", c["name"], "head")
            if c["detail"]:
                self.txt.insert("end", f" — {c['detail']}\n")
            else:
                self.txt.insert("end", "\n")
            if not c["ok"] and c["hint"]:
                for line in c["hint"].splitlines():
                    self.txt.insert("end", f"      {line}\n", "fail")
                self.txt.insert("end", "\n")
        self.txt.config(state="disabled")


class LlmPage(WizardPage):
    title = "Подключение к Ollama"
    subtitle = "Сервер Ollama должен быть запущен на этой машине."

    def __init__(self, wizard: "SetupWizard"):
        super().__init__(wizard)
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="URL Ollama:").grid(row=0, column=0, sticky="w", pady=4)
        self.url_var = tk.StringVar(value=self.cfg.ollama_url or "http://localhost:11434")
        ttk.Entry(self, textvariable=self.url_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(self, text="Проверить", command=self.fetch_models).grid(
            row=0, column=2, padx=(6, 0), pady=4)

        ttk.Label(self, text="Модель:").grid(row=1, column=0, sticky="w", pady=4)
        self.model_var = tk.StringVar(value=self.cfg.model)
        self.model_combo = ttk.Combobox(self, textvariable=self.model_var)
        self.model_combo.grid(row=1, column=1, columnspan=2, sticky="ew", pady=4)

        self.status = ttk.Label(self, text="", foreground="gray",
                                wraplength=680, justify="left")
        self.status.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        hint = (
            "Если Ollama ещё не установлен:\n"
            "    1. Скачайте с https://ollama.com/download и установите.\n"
            "    2. Сервер должен быть запущен (обычно сам поднимается).\n"
            "    3. Скачайте модель, например:\n"
            "           ollama pull qwen2.5:7b\n"
            "    4. Нажмите «Проверить» — модель появится в списке.\n\n"
            "Хорошо работают: qwen2.5:7b, qwen2.5:14b, llama3.1:8b.\n"
            "Модель держится в памяти (keep_alive) — это ускоряет обработку."
        )
        ttk.Label(self, text=hint, justify="left", foreground="gray",
                  wraplength=680).grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 0))

    def on_enter(self) -> None:
        self.fetch_models(silent=True)

    def fetch_models(self, silent: bool = False) -> None:
        from .llm import list_ollama_models
        url = self.url_var.get().strip() or "http://localhost:11434"
        try:
            models = list_ollama_models(url, timeout=5)
        except Exception as e:
            if not silent:
                self.status.config(text=f"Не удалось подключиться: {e}", foreground="red")
            self.model_combo["values"] = []
            return
        self.model_combo["values"] = models
        if models:
            if self.model_var.get() not in models:
                self.model_var.set(models[0])
            self.status.config(
                text=f"Найдено моделей: {len(models)} — {', '.join(models)}",
                foreground="#1a7f1a",
            )
        else:
            self.status.config(
                text="Ollama запущен, но моделей нет. Выполните: ollama pull qwen2.5:7b",
                foreground="#b06000",
            )

    def apply(self) -> None:
        self.cfg.ollama_url = self.url_var.get().strip() or "http://localhost:11434"
        self.cfg.model = self.model_var.get().strip() or "qwen2.5:7b"


class FoldersPage(WizardPage):
    title = "Папки и таблица"
    subtitle = "Где искать счета и куда складывать результаты."

    def __init__(self, wizard: "SetupWizard"):
        super().__init__(wizard)
        self.columnconfigure(1, weight=1)

        self.watch_var = tk.StringVar(value=self.cfg.watch_dir)
        self._row(0, "Папка со счетами:", self.watch_var, "dir",
                  "PDF-файлы кладутся сюда; сюда же ведётся наблюдение.")

        self.output_var = tk.StringVar(value=self.cfg.output_dir)
        self._row(2, "Куда переименовывать:", self.output_var, "dir",
                  "Оставьте пустым, чтобы переименовывать на месте.")

        self.table_var = tk.StringVar(value=self.cfg.table_path)
        self._row(4, "Файл таблицы (.xlsx/.csv):", self.table_var, "table",
                  "Сюда добавляется строка по каждому счёту.")

        self.table_enabled = tk.BooleanVar(value=self.cfg.table_enabled)
        ttk.Checkbutton(self, text="Вести учёт в таблице",
                        variable=self.table_enabled).grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _row(self, row: int, label: str, var: tk.StringVar, kind: str, hint: str) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=(6, 0))
        ttk.Entry(self, textvariable=var).grid(row=row, column=1, sticky="ew", pady=(6, 0))
        ttk.Button(self, text="Обзор…",
                   command=lambda v=var, k=kind: self._browse(v, k)).grid(
            row=row, column=2, padx=(6, 0), pady=(6, 0))
        if hint:
            ttk.Label(self, text=hint, foreground="gray").grid(
                row=row + 1, column=1, sticky="w")

    def _browse(self, var: tk.StringVar, kind: str) -> None:
        if kind == "dir":
            p = filedialog.askdirectory()
        else:
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx", confirmoverwrite=False,
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if p:
            var.set(p)

    def apply(self) -> None:
        self.cfg.watch_dir = self.watch_var.get().strip()
        self.cfg.output_dir = self.output_var.get().strip()
        self.cfg.table_path = self.table_var.get().strip()
        self.cfg.table_enabled = bool(self.table_enabled.get())


class FinishPage(WizardPage):
    title = "Готово"
    subtitle = "Нажмите «Сохранить», чтобы применить настройки."

    def __init__(self, wizard: "SetupWizard"):
        super().__init__(wizard)
        self.summary = scrolledtext.ScrolledText(self, wrap="word", height=18,
                                                 state="disabled")
        self.summary.pack(fill="both", expand=True, pady=(6, 0))

    def on_enter(self) -> None:
        self.wizard.apply_all()
        cfg = self.wizard.cfg
        lines = [
            f"Папка со счетами:  {cfg.watch_dir or '(не задана)'}",
            f"Вывод:              {cfg.output_dir or '(на месте)'}",
            f"Таблица:            {('вкл' if cfg.table_enabled else 'выкл')} "
            f"— {cfg.table_path or '(не задана)'}",
            f"Колонок в таблице:  {len(cfg.table_columns)}",
            f"Ollama:             {cfg.ollama_url}",
            f"Модель:             {cfg.model}",
            f"Языки OCR:          {cfg.ocr_lang}",
            "",
            "Файл настроек:",
            f"  {CONFIG_PATH}",
        ]
        self.summary.config(state="normal")
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", "\n".join(lines))
        self.summary.config(state="disabled")


# --------------------------------------------------------------------------- #
# Wizard window
# --------------------------------------------------------------------------- #
class SetupWizard(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Invoice Renamer — первичная настройка")
        self.geometry("800x600")
        self.minsize(720, 520)
        self.cfg: Config = load_config()
        self.saved: bool = False
        self.pages: list[WizardPage] = []
        self.index: int = 0
        self._build_chrome()
        self._build_pages()
        self._center()
        self._show(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _center(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _build_chrome(self) -> None:
        top = ttk.Frame(self, padding=(14, 12, 14, 4))
        top.pack(fill="x")
        self.title_lbl = ttk.Label(top, text="", font=("TkDefaultFont", 14, "bold"))
        self.title_lbl.pack(anchor="w")
        self.sub_lbl = ttk.Label(top, text="", foreground="gray")
        self.sub_lbl.pack(anchor="w")

        ttk.Separator(self).pack(fill="x")

        self.body = ttk.Frame(self, padding=14)
        self.body.pack(fill="both", expand=True)

        nav = ttk.Frame(self, padding=(14, 8))
        nav.pack(fill="x")
        self.btn_back = ttk.Button(nav, text="◀ Назад", command=self._back)
        self.btn_back.pack(side="left")
        self.btn_cancel = ttk.Button(nav, text="Отмена", command=self._on_close)
        self.btn_cancel.pack(side="left", padx=(6, 0))
        self.btn_next = ttk.Button(nav, text="Далее ▶", command=self._next)
        self.btn_next.pack(side="right")
        self.btn_finish = ttk.Button(nav, text="Сохранить и продолжить",
                                     command=self._finish)
        self.btn_finish.pack(side="right", padx=(0, 6))
        self.step_lbl = ttk.Label(nav, text="", foreground="gray")
        self.step_lbl.pack(side="right", padx=10)

    def _build_pages(self) -> None:
        self.pages = [
            WelcomePage(self),
            DependenciesPage(self),
            LlmPage(self),
            FoldersPage(self),
            FinishPage(self),
        ]

    def _show(self, i: int) -> None:
        if i < 0 or i >= len(self.pages):
            return
        for p in self.pages:
            p.pack_forget()
        self.index = i
        page = self.pages[i]
        self.title_lbl.config(text=page.title)
        self.sub_lbl.config(text=page.subtitle)
        page.pack(fill="both", expand=True)
        page.on_enter()
        self.btn_back.config(state="normal" if i > 0 else "disabled")
        self.btn_next.config(state="normal" if i < len(self.pages) - 1 else "disabled")
        self.btn_finish.config(state="normal" if i == len(self.pages) - 1 else "disabled")
        self.step_lbl.config(text=f"Шаг {i + 1} из {len(self.pages)}")

    def _back(self) -> None:
        if self.index > 0:
            self._show(self.index - 1)

    def _next(self) -> None:
        page = self.pages[self.index]
        try:
            page.apply()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=self)
            return
        if not page.on_leave():
            return
        self._show(self.index + 1)

    def apply_all(self) -> None:
        for p in self.pages:
            try:
                p.apply()
            except Exception:
                pass

    def _finish(self) -> None:
        # Собираем значения со всех страниц
        self.apply_all()
        try:
            validate_config(self.cfg)
        except ConfigError as e:
            messagebox.showerror("Проверьте настройки", str(e), parent=self)
            return
        self.cfg.setup_completed = True
        try:
            save_config(self.cfg)
        except OSError as e:
            messagebox.showerror("Не удалось сохранить настройки", str(e), parent=self)
            return
        self.saved = True
        messagebox.showinfo(
            "Настройка завершена",
            f"Настройки сохранены в\n{CONFIG_PATH}\n\n"
            "Сейчас откроется основное окно.",
            parent=self,
        )
        self.destroy()

    def _on_close(self) -> None:
        if not self.saved:
            if not messagebox.askyesno(
                "Выйти без сохранения?",
                "Настройки не сохранены. Выйти из мастера?", parent=self):
                return
        self.destroy()
