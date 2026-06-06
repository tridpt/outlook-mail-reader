@echo off
setlocal

set "APP_DIR=%~dp0"
set "HOST=0.0.0.0"
set "PORT=8809"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"

rem Public Azure app client id used by the existing local token cache.
if not defined OUTLOOK_CLIENT_ID set "OUTLOOK_CLIENT_ID=cbf20cb1-8560-4ca3-995b-310d68f2bd1b"

cd /d "%APP_DIR%"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found in PATH.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  python -m venv .venv
)

call ".venv\Scripts\activate.bat"

echo Installing requirements...
python -X utf8 -m pip install -r requirements.txt
if errorlevel 1 (
  echo Failed to install requirements.
  pause
  exit /b 1
)

cd /d "%APP_DIR%backend"
echo Starting Outlook Mail Reader on http://127.0.0.1:%PORT%
echo Press Ctrl+C to stop the server.
python -m uvicorn main:app --host %HOST% --port %PORT%

pause
