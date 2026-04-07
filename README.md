# DELR Ledger

[繁體中文](README_zh.md)

DELR Ledger is a lightweight desktop bookkeeping app (Tkinter) for personal ledger tracking, with multi-language UI and local-file-first data storage.

This project was completed with assistance from Codex.

## Version

Current version: `v0.1.0`

## Highlights

- Desktop GUI app built with Python + Tkinter
- Multi-language UI: `繁體中文`, `English`, `Deutsch`
- Ledger format: `.delr` (CSV-compatible content)
- Open/create/export/import ledgers
- Supported import/export formats:
  - `.delr`
  - `.csv`
  - `.tsv`
  - `.xlsx`
  - `.json`
  - `.xml`
  - `.yaml` / `.yml`
- Per-currency totals (Total / Income / Expense)
- Table header filters (toggle on/off)

## Project Structure

- `delr_ledger_app.py`: Main GUI application
- `expenses.py`: Earlier CLI utility (legacy/simple flow)
- `scripts/`: Build and release scripts
- `VERSION`: App version source
- `LICENSE`: MIT License

Runtime/output folders (generated):

- `dist/`: Build output
- `build/`: Intermediate build artifacts

## Quick Start (Dev)

1. Create and activate a virtual environment.
2. Install dependencies.
3. Run the GUI app.

Example (PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl pyyaml
python .\delr_ledger_app.py
```

## Build EXE

Use your existing scripts under `scripts/` (as configured in this workspace). Typical flow:

```powershell
.\scripts\build.bat
```

or release flow:

```powershell
.\scripts\release.bat
```

Notes:

- Release output is expected under `dist/`.
- Release packaging includes versioned zip artifacts.

## Data Notes

- Default ledger storage can be configured in-app.
- `.delr` files store tabular ledger rows with CSV headers:
  - `date, amount, item, unit, payment, merchant, category`
- `settings.json` stores app-level preferences (language, last folder/file, etc.).

## Optional Dependencies

Some formats require extra packages:

- `XLSX` support: `openpyxl`
- `YAML` support: `PyYAML`

If missing, the app will show a runtime error message when that format is used.

## License

MIT License. See [LICENSE](LICENSE).


