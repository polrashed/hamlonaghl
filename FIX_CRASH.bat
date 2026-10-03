@echo off
setlocal
cd /d "%~dp0"
where python >nul 2>nul || (echo Python not found.& pause & exit /b 1)
echo Repairing compatible core packages...
python -m pip install --upgrade --force-reinstall "numpy<2" "openpyxl>=3.1,<4" "pillow>=9"
if errorlevel 1 (echo Repair failed.& pause & exit /b 1)
echo Repair complete. Restart the application.
pause
