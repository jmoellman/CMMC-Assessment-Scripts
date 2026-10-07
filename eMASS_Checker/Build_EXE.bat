@echo off
setlocal
cd /d "%~dp0"
REM Optional: builds a standalone eMASS_Checker.exe (no Python needed on the target PC).
REM Run Run_eMASS_Checker.bat once first so the .venv exists.
if not exist ".venv\Scripts\python.exe" (
  echo Run Run_eMASS_Checker.bat once first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q pyinstaller
if errorlevel 1 (echo Could not install PyInstaller. & pause & exit /b 1)
".venv\Scripts\pyinstaller.exe" --noconfirm --clean --windowed --name eMASS_Checker ^
  --add-data "config;config" --add-data "data;data" --add-data "README.md;." ^
  emass_checker.py
if errorlevel 1 (echo Build failed. & pause & exit /b 1)
echo.
echo Done. Your app is in: dist\eMASS_Checker\eMASS_Checker.exe
echo Copy the whole dist\eMASS_Checker folder; the .exe needs the files next to it.
pause
