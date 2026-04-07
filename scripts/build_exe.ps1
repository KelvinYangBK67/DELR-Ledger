param(
    [ValidateSet('cli','gui')]
    [string]$Mode = 'gui',
    [switch]$InstallPyInstaller
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $projectRoot

if ($InstallPyInstaller) {
    python -m pip install --upgrade pip
    python -m pip install pyinstaller
}

if ($Mode -eq 'gui') {
    python -m PyInstaller --noconfirm --clean --windowed --name "DELR-Ledger" .\delr_ledger_app.py
    Write-Host "Build complete: dist\\DELR-Ledger\\DELR-Ledger.exe"
}
else {
    python -m PyInstaller --noconfirm --clean --name "DELR-Ledger" .\expenses.py
    Write-Host "Build complete: dist\\DELR-Ledger\\DELR-Ledger.exe"
}
