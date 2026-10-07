@echo off
setlocal
cd /d "%~dp0"
REM CMMC eMASS Results Checker launcher. First run creates a private Python environment (.venv).
REM Tip: drag a results .xlsx onto this file to open it directly.

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY goto :nopy

if not exist ".venv\Scripts\pythonw.exe" (
  echo First run: setting up the checker. This takes about a minute...
  %PY% -m venv .venv
  if errorlevel 1 goto :nopy
  ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt
  if errorlevel 1 goto :pipfail
  ".venv\Scripts\python.exe" emass_checker.py --selftest
  if errorlevel 1 goto :selftestfail
)
start "" ".venv\Scripts\pythonw.exe" "%~dp0emass_checker.py" %*
exit /b 0

:nopy
echo.
echo Python 3.10 or newer was not found.
echo Install it from https://www.python.org/downloads/windows/
echo During setup check "Add python.exe to PATH" and keep "tcl/tk and IDLE" selected.
pause
exit /b 1

:pipfail
echo.
echo Could not install openpyxl. If your network blocks pypi.org, ask IT to allow it,
echo or copy a .venv folder from a machine where setup succeeded.
rmdir /s /q .venv 2>nul
pause
exit /b 1

:selftestfail
echo.
echo Self-test failed. See the messages above.
rmdir /s /q .venv 2>nul
pause
exit /b 1
