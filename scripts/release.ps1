param(
    [ValidateSet('cli','gui')]
    [string]$Mode = 'gui'
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $projectRoot

$version = (Get-Content -LiteralPath .\VERSION -Raw).Trim()
if (-not $version) {
    throw 'VERSION file is empty.'
}

powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1 -Mode $Mode

$releaseRoot = Join-Path $projectRoot 'dist'
$releaseName = "DELR-Ledger-$version"
$zipPath = Join-Path $releaseRoot ("$releaseName.zip")
$tempRoot = Join-Path $projectRoot '.release_tmp'
$tempStage = Join-Path $tempRoot $releaseName

if (Test-Path $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
New-Item -ItemType Directory -Path $tempStage -Force | Out-Null

$srcDir = Join-Path $projectRoot 'dist\DELR-Ledger'
robocopy $srcDir (Join-Path $tempStage 'DELR-Ledger') /E /R:5 /W:2 /NFL /NDL /NJH /NJS /NC /NS | Out-Null
$rc = $LASTEXITCODE
if ($rc -ge 8) { throw "robocopy failed with exit code $rc" }

Copy-Item -LiteralPath .\LICENSE -Destination (Join-Path $tempStage 'LICENSE') -Force
Copy-Item -LiteralPath .\VERSION -Destination (Join-Path $tempStage 'VERSION') -Force

if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $tempStage '*') -DestinationPath $zipPath -CompressionLevel Optimal

Remove-Item -LiteralPath $tempRoot -Recurse -Force
Write-Host "Release zip: $zipPath"
