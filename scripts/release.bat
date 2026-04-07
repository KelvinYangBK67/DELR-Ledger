@echo off
setlocal
cd /d "%~dp0\.."
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\release.ps1" -Mode cli
if errorlevel 1 (
  echo.
  echo Release build failed.
  pause
  exit /b 1
)
echo.
echo Release build succeeded.
pause
