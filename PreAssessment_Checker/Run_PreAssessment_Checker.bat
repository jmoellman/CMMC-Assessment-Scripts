@echo off
setlocal
cd /d "%~dp0"
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY set "PY=python"
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo Python 3.10 or newer is required.
  echo Install it from https://www.python.org/downloads/ and check "Add python.exe to PATH".
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  echo First run: setting up the checker. This takes about a minute...
  %PY% -m venv .venv
  if errorlevel 1 goto err
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto err
)
start "" ".venv\Scripts\pythonw.exe" app.py %*
exit /b 0
:err
echo.
echo Setup failed. Check your internet connection for the one-time package install, then run this file again.
echo If it keeps failing, delete the .venv folder and try again.
pause
exit /b 1
