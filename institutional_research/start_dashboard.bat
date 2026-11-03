@echo off
setlocal
cd /d "%~dp0"

rem OPEN SAVED DATA ONLY. Never install dependencies or download market data on an
rem offline dashboard launch. Run Sync All Data in Research Hub while online to refresh.
set "ANTZAZ_OFFLINE=1"

if not exist "portfolio.csv" (
  echo ERROR: Portfolio input file is missing: %CD%\portfolio.csv
  echo Connect to GitHub and pull the latest code while online.
  pause
  exit /b 1
)
if not exist "outputs\latest\summary.json" (
  echo No saved analysis was found. The app will open Setup ^& Holdings,
  echo where you can inspect the portfolio and build the full analysis when online.
)

echo Opening saved portfolio data in offline mode.
echo No internet connection or market-data download is required.
python -m streamlit run app.py --server.port 8501 --server.headless false --browser.gatherUsageStats false
if errorlevel 1 (
  echo ERROR: Streamlit could not start. Run install_desktop_app.bat while online.
  pause
  exit /b 1
)
