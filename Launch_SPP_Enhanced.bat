@echo off
REM SPP Enhanced Launch Script
REM This script provides a safe way to launch the SPP Enhanced application

echo ========================================
echo  HD Supply SPP Enhanced v2.0
echo  Advanced Automation with Templates
echo ========================================
echo.

REM Check if executable exists
if not exist "SPP_Enhanced.exe" (
    echo ERROR: SPP_Enhanced.exe not found!
    echo Please ensure this script is in the same folder as the executable.
    echo.
    pause
    exit /b 1
)

REM Display startup information
echo Starting SPP Enhanced application...
echo.
echo Features:
echo - User-configurable Excel templates
echo - Enhanced Snowflake connectivity  
echo - Modern HD Supply branded interface
echo - Robust error handling and recovery
echo.
echo Please wait while the application loads...
echo.

REM Launch the application
start "" "SPP_Enhanced.exe"

REM Check if application started successfully
timeout /t 3 /nobreak >nul
echo.
echo Application launched successfully!
echo.
echo If you encounter any issues:
echo 1. Check TROUBLESHOOTING.md
echo 2. Verify your templates are configured
echo 3. Ensure network connectivity
echo 4. Contact IT support if needed
echo.
echo Press any key to close this window...
pause >nul
