@echo off
setlocal
cd /d "%~dp0"

echo ===============================================
echo   Antzaz Investment Research - Desktop Setup
echo ===============================================
echo.

where git >nul 2>nul
if errorlevel 1 (
  echo ERROR: Git is not installed or not on PATH.
  echo Install Git or GitHub Desktop, then run this installer again.
  pause
  exit /b 1
)

where py >nul 2>nul
if not errorlevel 1 (
  set "PY=py -3"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo ERROR: Python 3 is not installed or not on PATH.
    pause
    exit /b 1
  )
  set "PY=python"
)

echo Updating repository...
git pull --ff-only origin main
if errorlevel 1 (
  echo.
  echo WARNING: GitHub update could not fast-forward.
  echo The installer will continue with the local version.
)

echo.
echo Installing equity-research dependencies...
%PY% -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo ERROR: Equity-research dependency installation failed.
  pause
  exit /b 1
)

echo.
echo Installing portfolio-dashboard dependencies...
%PY% -m pip install -r "%~dp0institutional_research\requirements.txt"
if errorlevel 1 (
  echo ERROR: Portfolio dependency installation failed.
  pause
  exit /b 1
)

echo.
echo Creating desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $desktop=[Environment]::GetFolderPath('Desktop'); $s=$ws.CreateShortcut((Join-Path $desktop 'Antzaz Investment Research.lnk')); $s.TargetPath=$env:WINDIR+'\System32\wscript.exe'; $s.Arguments='\"%~dp0launch_investment_app.vbs\"'; $s.WorkingDirectory='%~dp0'; $s.IconLocation=$env:WINDIR+'\System32\shell32.dll,220'; $s.Description='Antzaz private equity research and portfolio launcher'; $s.Save()"
if errorlevel 1 (
  echo ERROR: Could not create the desktop shortcut.
  pause
  exit /b 1
)

echo.
echo Desktop app installed successfully.
echo You can now double-click "Antzaz Investment Research" on your desktop.
echo.
choice /M "Open the app now"
if errorlevel 2 exit /b 0
wscript.exe "%~dp0launch_investment_app.vbs"
