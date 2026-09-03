@echo off
REM Bank Reference Data — Windows launcher.
REM
REM   run.bat                              start the web app (opens in your browser)
REM   run.bat --countries US,GB,IN         export straight to Excel, no browser
REM   run.bat --sepa --out sepa.xlsx       any refdata.py arguments work here
REM
REM First run creates a private .venv in this folder and installs requirements.
REM You must be on the Nium VPN - the API is not reachable from elsewhere.
setlocal
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo Python not found. Install it from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" in the installer, then re-run.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment ^(.venv^)...
  %PY% -m venv .venv || goto :fail
)
call ".venv\Scripts\activate.bat"

if not exist ".venv\.deps-installed" (
  echo Installing requirements...
  python -m pip install --quiet --upgrade pip
  python -m pip install --quiet -r requirements.txt || goto :fail
  type nul > ".venv\.deps-installed"
)

set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

if not "%~1"=="" (
  python refdata.py %*
  goto :end
)

echo Starting the web app - press Ctrl+C to stop.
streamlit run app.py
goto :end

:fail
echo.
echo Setup failed - see the messages above.
pause
exit /b 1

:end
endlocal
