@echo off
setlocal
cd /d "%~dp0"

rem Keep the launcher itself current before it opens.
where git >nul 2>nul
if not errorlevel 1 git pull --ff-only origin main > "%TEMP%\antzaz_investment_launcher_update.log" 2>&1

where pyw >nul 2>nul
if not errorlevel 1 (
  start "" pyw -3 "%~dp0desktop_launcher.py"
  exit /b 0
)

where pythonw >nul 2>nul
if not errorlevel 1 (
  start "" pythonw "%~dp0desktop_launcher.py"
  exit /b 0
)

where py >nul 2>nul
if not errorlevel 1 (
  start "" py -3 "%~dp0desktop_launcher.py"
  exit /b 0
)

python "%~dp0desktop_launcher.py"
