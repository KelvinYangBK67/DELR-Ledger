# DELR Ledger Quick Start

## Build (stable on this machine)

Use onedir mode to avoid onefile extraction permission issues:

```powershell
python -m PyInstaller --noconfirm --clean --name "DELR-Ledger" .\expenses.py
```

Output:
- `dist\DELR-Ledger\DELR-Ledger.exe`

## Run

```powershell
.\dist\DELR-Ledger\DELR-Ledger.exe --help
.\dist\DELR-Ledger\DELR-Ledger.exe summary --month 2026-04
```
