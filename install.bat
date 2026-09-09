@echo off
cd /d "%~dp0"
py --version >nul 2>&1
if errorlevel 1 (
  echo Install 64-bit Python from python.org first, then run this file again.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" py -m venv .venv
if not exist ".venv\Scripts\python.exe" (
  echo Could not create the Python environment. See README.md.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m pip install --no-cache-dir -r requirements.txt
if errorlevel 1 (
  echo Installation failed. Check your internet connection and see README.md.
  pause
  exit /b 1
)
echo Installation complete. Double-click launch.bat to open the app.
pause
