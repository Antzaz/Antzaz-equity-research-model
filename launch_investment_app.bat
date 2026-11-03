@echo off
setlocal
cd /d "%~dp0"

rem Offline-first: NEVER call GitHub when merely opening the desktop application.
rem The Research Hub has explicit Update from GitHub and Sync All Data actions.

where pyw >nul 2>nul
if not errorlevel 1 (
  start "" pyw -3 "%~dp0desktop_research_hub.py"
  exit /b 0
)

where pythonw >nul 2>nul
if not errorlevel 1 (
  start "" pythonw "%~dp0desktop_research_hub.py"
  exit /b 0
)

where py >nul 2>nul
if not errorlevel 1 (
  start "" py -3 "%~dp0desktop_research_hub.py"
  exit /b 0
)

python "%~dp0desktop_research_hub.py"
