@echo off
cd /d "%~dp0"
echo Refreshing offline portfolio research...
python -m pip install -r requirements.txt
python run_research.py
if errorlevel 1 (
  echo.
  echo Research refresh failed. The dashboard was not started because its data may be incomplete.
  echo Check the ERROR line above. Optional CSV inputs may contain malformed data.
  pause
  exit /b 1
)
echo.
echo Research refresh complete. Starting Streamlit...
python -m streamlit run app.py
