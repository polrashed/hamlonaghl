@echo off
setlocal
cd /d "%~dp0"
where python >nul 2>nul || (echo Python not found.& pause & exit /b 1)
python -m pip install -r requirements.txt
if errorlevel 1 (echo Dependency installation failed.& pause & exit /b 1)
python -m PyInstaller --noconfirm --clean --windowed --onefile --name DoustanTransport --icon assets\app_icon.ico --version-file version_info.txt --add-data "assets;assets" app.py
if errorlevel 1 (echo Build failed.& pause & exit /b 1)
echo Build complete: dist\DoustanTransport.exe
pause
