param(
    [ValidateSet('cli','gui')]
    [string]$Mode = 'cli'
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $projectRoot

$version = (Get-Content -LiteralPath .\VERSION -Raw).Trim()
if (-not $version) {
    throw 'VERSION file is empty.'
}

powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1 -Mode $Mode

$releaseRoot = Join-Path $projectRoot 'release'
$releaseName = "DELR-Ledger-$version"
$releaseDir = Join-Path $releaseRoot $releaseName
$zipPath = Join-Path $releaseRoot ("$releaseName.zip")

if (Test-Path $releaseDir) { Remove-Item -LiteralPath $releaseDir -Recurse -Force }
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

if ($Mode -eq 'cli') {
    Copy-Item -LiteralPath .\dist\DELR-Ledger -Destination $releaseDir -Recurse -Force
} else {
    Copy-Item -LiteralPath .\dist\DELR-Ledger-GUI.exe -Destination (Join-Path $releaseDir 'DELR-Ledger-GUI.exe') -Force
}

Copy-Item -LiteralPath .\LICENSE -Destination (Join-Path $releaseDir 'LICENSE') -Force
Copy-Item -LiteralPath .\VERSION -Destination (Join-Path $releaseDir 'VERSION') -Force

if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $releaseDir '*') -DestinationPath $zipPath -CompressionLevel Optimal

Write-Host "Release folder: $releaseDir"
Write-Host "Release zip: $zipPath"
