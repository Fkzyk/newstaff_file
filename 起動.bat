@echo off
rem ============================================
rem  Nyusha Shorui App Launcher (Windows)
rem  First run installs required libraries.
rem ============================================
cd /d "%~dp0"

set PY=python
where py >nul 2>nul
if %errorlevel%==0 set PY=py -3

echo Checking libraries (first run may take 1-2 minutes)...
%PY% -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo.
  echo [ERROR] Failed to install libraries.
  echo Please make sure Python is installed:
  echo   https://www.python.org/downloads/
  echo   (Check "Add python.exe to PATH" during installation)
  pause
  exit /b 1
)

echo Starting the app... A browser window will open.
%PY% -m streamlit run app.py --server.headless false
pause
