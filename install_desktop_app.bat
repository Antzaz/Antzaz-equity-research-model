@echo off
setlocal
cd /d "%~dp0"

echo Updating project before installing desktop launcher...
git pull --ff-only origin main
if errorlevel 1 (
  echo.
  echo Git update failed. Fix the Git error above, then run this installer again.
  pause
  exit /b 1
)

for /f "delims=" %%P in ('python -c "import sys; print(sys.executable)"') do set "PYTHON=%%P"
if not defined PYTHON (
  echo Python was not found.
  pause
  exit /b 1
)

set "PYTHONW=%PYTHON:python.exe=pythonw.exe%"
if not exist "%PYTHONW%" set "PYTHONW=%PYTHON%"

set "SCRIPT=%CD%\desktop_research_hub.py"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$desktop=[Environment]::GetFolderPath('Desktop');" ^
  "$ws=New-Object -ComObject WScript.Shell;" ^
  "$lnk=$ws.CreateShortcut((Join-Path $desktop 'Antzaz Research Hub.lnk'));" ^
  "$lnk.TargetPath='%PYTHONW%';" ^
  "$lnk.Arguments='\"%SCRIPT%\"';" ^
  "$lnk.WorkingDirectory='%CD%';" ^
  "$lnk.Description='Antzaz private equity research and portfolio launcher';" ^
  "$lnk.IconLocation='%SystemRoot%\System32\shell32.dll,220';" ^
  "$lnk.Save()"

if errorlevel 1 (
  echo Failed to create the desktop shortcut.
  pause
  exit /b 1
)

echo.
echo Installed: Antzaz Research Hub
echo A shortcut has been created on your Windows Desktop.
echo Double-click it from now on. It will update from GitHub automatically.
pause
