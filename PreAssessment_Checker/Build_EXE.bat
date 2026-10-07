@echo off
REM Optional: builds a standalone Windows app folder (no Python needed on the target PC).
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run Run_PreAssessment_Checker.bat once first to set up Python packages.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m pip install pyinstaller
if errorlevel 1 goto err
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --windowed --name "PreAssessment_Checker" --icon assets\app.ico --add-data "config;config" --add-data "assets;assets" app.py
if errorlevel 1 goto err
xcopy /E /I /Y config "dist\PreAssessment_Checker\config" >nul
echo.
echo Done. Run dist\PreAssessment_Checker\PreAssessment_Checker.exe
echo Edit dist\PreAssessment_Checker\config\rules.json to change rules.
pause
exit /b 0
:err
echo Build failed.
pause
exit /b 1
