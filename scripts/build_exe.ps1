param(
    [ValidateSet('cli','gui')]
    [string]$Mode = 'gui',
    [switch]$InstallPyInstaller
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $projectRoot

function Remove-PathWithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [int]$Retries = 5,
        [int]$DelaySeconds = 1
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    for ($i = 1; $i -le $Retries; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            return
        }
        catch {
            if ($i -eq $Retries) { throw }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

if ($InstallPyInstaller) {
    python -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed with exit code $LASTEXITCODE" }
    python -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) { throw "pyinstaller install failed with exit code $LASTEXITCODE" }
}

Remove-PathWithRetry -Path (Join-Path $projectRoot 'build\DELR-Ledger')
Remove-PathWithRetry -Path (Join-Path $projectRoot 'dist\DELR-Ledger')

if ($Mode -eq 'gui') {
    python -m PyInstaller --noconfirm --windowed --name "DELR-Ledger" .\delr_ledger_app.py
}
else {
    python -m PyInstaller --noconfirm --name "DELR-Ledger" .\expenses.py
}

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

Write-Host "Build complete: dist\DELR-Ledger\DELR-Ledger.exe"
