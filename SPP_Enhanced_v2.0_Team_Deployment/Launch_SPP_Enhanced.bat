@echo off
title SPP Automation Tool Enhanced
echo Starting SPP Automation Tool Enhanced...
echo.
echo HD Supply Chain Excellence
echo Developer: Ben F. Benjamaa
echo Manager: Lauren B. Trapani
echo.

REM Check if executable exists
if not exist "SPP_Automation_Tool_Enhanced.exe" (
    echo ERROR: SPP_Automation_Tool_Enhanced.exe not found!
    echo Please ensure this script is in the same directory as the executable.
    pause
    exit /b 1
)

REM Launch the application
start "SPP Automation Tool Enhanced" "SPP_Automation_Tool_Enhanced.exe"

REM Optional: Keep console open for troubleshooting
REM pause
