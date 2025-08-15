@echo off
echo Installing SPP Automation Requirements...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from python.org
    pause
    exit /b 1
)

echo Python found. Installing packages...
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install packages
    echo Try running as Administrator or check internet connection
    pause
    exit /b 1
)

echo.
echo ✅ Installation completed successfully!
echo.
echo Next steps:
echo 1. Edit config.ini with your email address
echo 2. Run Quick_Start.bat to begin
echo.
pause
