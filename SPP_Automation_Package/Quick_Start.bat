@echo off
title SPP Automation - Quick Start

echo.
echo  ==========================================
echo    SPP Metric Automation Tool - v1.0
echo  ==========================================
echo.
echo  Quick Start Menu:
echo.
echo  1. Launch GUI Tool (Recommended)
echo  2. Run Test Connection
echo  3. View Output Folder
echo  4. Edit Configuration
echo  5. View Documentation
echo  6. Exit
echo.

set /p choice="  Enter your choice (1-6): "

if "%choice%"=="1" (
    echo.
    echo  Starting SPP Automation GUI...
    python spp_gui.py
) else if "%choice%"=="2" (
    echo.
    echo  Testing Snowflake connection...
    python test_connection.py
    pause
) else if "%choice%"=="3" (
    echo.
    echo  Opening output folder...
    if not exist "Output" mkdir Output
    explorer Output
) else if "%choice%"=="4" (
    echo.
    echo  Opening configuration file...
    notepad config.ini
) else if "%choice%"=="5" (
    echo.
    echo  Opening documentation...
    start README.md
) else if "%choice%"=="6" (
    echo.
    echo  Goodbye!
    exit
) else (
    echo.
    echo  Invalid choice. Please try again.
    pause
    cls
    goto :eof
)

echo.
pause
