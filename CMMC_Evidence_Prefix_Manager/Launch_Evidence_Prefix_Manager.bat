@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY_CMD=py -3"
) else (
  set "PY_CMD=python"
)
%PY_CMD% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing Excel support...
  %PY_CMD% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Excel support could not be installed. CSV mappings will still work.
    pause
  )
)
%PY_CMD% evidence_prefix_manager.py --gui
if errorlevel 1 pause
endlocal
