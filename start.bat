@echo off
rem ============================================
rem  Nyusha Shorui App Launcher (Windows)
rem  Uses Python 3.12 for best library support.
rem ============================================
cd /d "%~dp0"

set PY=

rem 1) Use Python 3.12 if available
py -V:3.12 -c "pass" >nul 2>nul && set PY=py -V:3.12
if defined PY goto :found

rem 2) Try to install 3.12 via the Python Install Manager
where py >nul 2>nul
if %errorlevel%==0 (
  echo Installing Python 3.12 runtime - this may take a few minutes...
  py install 3.12
  py -V:3.12 -c "pass" >nul 2>nul && set PY=py -V:3.12
)
if defined PY goto :found

rem 3) Fall back to any Python 3
py -3 -c "pass" >nul 2>nul && set PY=py -3
if defined PY goto :found
python -c "pass" >nul 2>nul && set PY=python
if defined PY goto :found

echo.
echo [ERROR] Python not found.
echo Please install Python from https://www.python.org/downloads/
echo (Check "Add python.exe to PATH" during installation)
pause
exit /b 1

:found
echo Using Python: %PY%

rem Create/refresh the Desktop shortcut (Japanese name lives in the .ps1)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_shortcut.ps1" >nul 2>nul

%PY% -m ensurepip --upgrade >nul 2>nul

echo Checking libraries (first run may take 1-2 minutes)...
%PY% -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo.
  echo [ERROR] Failed to install libraries.
  echo Please send this window's messages to the support chat.
  pause
  exit /b 1
)

rem Register pywin32 COM support so Word/Excel PDF conversion works.
rem This is optional; ignore its exit code so it never blocks startup.
%PY% -m pywin32_postinstall -install -silent >nul 2>nul
cd /d "%~dp0"

echo.
echo Starting the app... A browser window will open shortly.
echo (Keep this black window open while using the app. Do NOT press any key here.)
echo.
%PY% -m streamlit run app.py --server.headless false
echo.
echo The app has stopped. You can close this window.
pause
