@echo off
setlocal

set "PORT=8809"

echo Stopping Outlook Mail Reader on port %PORT%...

for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  echo Stopping PID %%P
  taskkill /PID %%P /F >nul 2>nul
)

for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  echo Failed to stop PID %%P
  pause
  exit /b 1
)

echo Server on port %PORT% stopped.
pause
