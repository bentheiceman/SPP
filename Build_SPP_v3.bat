@echo off
REM Build script for SPP Automation Tool v3.0 with PDH Compliance
REM This creates a standalone .exe for team deployment

echo ============================================================
echo SPP Automation Tool v3.0 - Build Script
echo With PDH Compliance Feature
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher
    pause
    exit /b 1
)

echo [Step 1/4] Checking dependencies...
echo.

REM Install/upgrade required packages
echo Installing PyInstaller...
pip install --upgrade pyinstaller >nul 2>&1

echo Installing required packages...
pip install --upgrade pandas openpyxl snowflake-connector-python >nul 2>&1

echo ✓ Dependencies ready
echo.

echo [Step 2/4] Running build script...
echo.

REM Run the Python build script
python build_spp_v3.py

if errorlevel 1 (
    echo.
    echo ❌ Build failed! Check errors above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo ✅ BUILD SUCCESSFUL!
echo ============================================================
echo.
echo Your team deployment package is ready!
echo.
echo Next steps:
echo 1. Locate the SPP_v3.0_PDH_Deployment_* folder
echo 2. Zip it and share with your team
echo 3. Team members extract and run the .exe
echo.
echo New Feature: PDH Compliance data now in Tab4!
echo.
pause
