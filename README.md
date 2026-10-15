# Invoice Renamer

Watches a folder for PDF invoices (in Russian), reads them with **Tesseract OCR**, asks a **local LLM via Ollama** to pull out the key fields, and renames each file:

```
260924 АРГО Сч № КА-16339 металл.pdf
```

Every processed file is also logged as a new row in an Excel (`.xlsx`) or CSV table. Everything is configured from a small desktop UI.

## What it extracts

| Field | Format | Example |
|---|---|---|
| Date | `ддммгг` (configurable) | `260924` |
| Supplier | short name, no ООО/АО/quotes | `АРГО` |
| Invoice number | `буквы-цифры` | `КА-16339` |
| Main material | lowercase, nominative | `металл` |

The date is returned by the LLM as an ISO date and formatted by the script, which is more reliable than asking the model to produce `260924` directly. Latin look-alike letters in the invoice number (`KA` → `КА`) are converted to Cyrillic by default, because OCR often confuses them (switchable in the UI).

## Requirements

1. **Python 3.10+**
2. **Tesseract OCR** with the Russian language data
   - Windows: installer from <https://github.com/UB-Mannheim/tesseract/wiki> (tick *Russian* during setup)
   - Debian/Ubuntu: `sudo apt install tesseract-ocr tesseract-ocr-rus python3-tk`
   - macOS: `brew install tesseract tesseract-lang`
3. **Ollama** (<https://ollama.com>) running locally, with a model that handles Russian well:
   ```
   ollama pull qwen2.5:7b
   ```
   `qwen2.5:14b` is more accurate if you have the RAM/VRAM.

PDF pages are rendered to images with PyMuPDF, so no poppler installation is needed.

## Install

```
pip install -r requirements.txt
```

## Run

```
python invoice_renamer.py              # settings UI
python invoice_renamer.py --headless   # no window, uses config.json
python invoice_renamer.py --config D:\invoices\config.json
```

In the UI: pick the **watch folder**, check the other tabs, press **Test Ollama**, then **▶ Start**. Settings are saved to `config.json` next to the script. **Scan once** processes the current contents of the folder a single time. The **Log** tab shows progress; the same output goes to `invoice_renamer.log`.

For a permanent background service, run `--headless` from Task Scheduler (Windows), systemd, or launchd.

## Settings

| Tab | Setting | Meaning |
|---|---|---|
| Folders | Watch folder / Output folder | Output empty = rename in place |
| | Scan interval, idle time | A file is only read after it has been untouched for N seconds, so half-copied files are skipped |
| | Move unrecognised files | Failed documents go to `_failed/` and are not retried |
| OCR | Languages, DPI, max pages | `rus+eng`, 300 DPI, first 3 pages by default |
| | Use PDF text layer | Skips OCR for digital PDFs that already contain text |
| LLM | URL, model, context, timeout | Defaults: `http://localhost:11434`, `qwen2.5:7b` |
| Naming | Template | Placeholders `{date} {supplier} {invoice} {material}`; **`Сч` is mandatory** and enforced |
| | Date format | `strftime` format, default `%d%m%y` |
| Table | File, create if missing, columns | See below |

## Table behaviour

* Supported: `.xlsx` and `.csv` (CSV uses `;` and UTF-8 with BOM so Excel opens it correctly; an existing delimiter is auto-detected).
* **Create if missing** (checkbox): if the file does not exist it is created with a header row.
* **Fewer columns than required** in the existing table → an error is logged and **no files are processed** until it is fixed (checked before anything is renamed).
* **More columns than required** → fine; the new row fills the first columns and the extra ones stay blank.
* If the table is open in Excel, writing is retried for ~10 seconds; if it still fails, the file is renamed anyway and the row is printed in the log so nothing is lost.

Default columns: Дата, Поставщик, Номер счёта, Материал, Исходное имя файла, Новое имя файла, Обработано. Choose and reorder them in the UI with the keys `date, supplier, invoice, material, original_name, new_name, processed_at`.

> Note: the request said “add a new *column*”; this tool interprets that as adding a new *row* (one record per invoice). The column-count rules above apply to the table's header.

## Failure handling

* Fields the LLM cannot find (date, supplier or number) → the file is moved to `_failed/` and the reason is logged. The material is optional; if it is missing the name is built without it.
* Ollama or Tesseract unreachable / model not pulled → files stay where they are and are retried on the next scan.
* Name collisions get ` (2)`, ` (3)`… appended. Files already matching `NNNNNN … Сч № …` are skipped.

## Tips

* Poor results on scans: raise DPI to 400, or try `--psm 4` in *Tesseract options*.
* Slow: lower DPI to 200, use a smaller model, or enable *Use PDF text layer*.
* The extraction prompt is the `SYSTEM_PROMPT` constant in `invoice_renamer.py` if you want to tune it (for example, to tell the model which buyer company to ignore).
