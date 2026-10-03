@echo off
setlocal
cd /d "%~dp0"
echo Doustan Transport - starting...
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.9+ and enable Add Python to PATH.
  pause
  exit /b 1
)
if not exist requirements.txt (
  echo requirements.txt is missing.
  pause
  exit /b 1
)
python -c "import openpyxl; from PIL import Image" >nul 2>nul
if errorlevel 1 (
  echo Installing required packages...
  python -m pip install -r requirements.txt
  if errorlevel 1 (echo Dependency installation failed.& pause & exit /b 1)
)
python app.py
if errorlevel 1 (echo The application ended with an error (code %errorlevel%).& pause & exit /b %errorlevel%)
endlocal
