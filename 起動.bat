@echo off
rem ============================================
rem  入社書類作成アプリ 起動用 (Windows)
rem  初回のみ必要なライブラリを自動インストールします
rem ============================================
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  set PY=py -3
) else (
  set PY=python
)

%PY% -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo.
  echo ライブラリのインストールに失敗しました。
  echo Python がインストールされているかご確認ください。
  echo https://www.python.org/downloads/ からインストールできます。
  echo ※インストール時に「Add python.exe to PATH」に必ずチェックを入れてください。
  pause
  exit /b 1
)

%PY% -m streamlit run app.py --server.headless false
pause
