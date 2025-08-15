@echo off
REM SPP Metric Automation - Quick Launch
REM This batch file provides easy access to the automation tools

echo SPP Metric Automation Tool
echo ===============================
echo.
echo Please choose an option:
echo 1. Launch GUI Interface (Recommended)
echo 2. Run Quick Automation Script
echo 3. View Sample Usage
echo 4. Edit Configuration File
echo 5. View Output Folder
echo 6. Exit
echo.

set /p choice="Enter your choice (1-6): "

if "%choice%"=="1" (
    echo Launching GUI interface...
    python spp_gui.py
) else if "%choice%"=="2" (
    echo Running quick automation...
    python quick_run.py
) else if "%choice%"=="3" (
    echo Showing sample usage...
    python sample_usage.py
    pause
) else if "%choice%"=="4" (
    echo Opening configuration file...
    if exist config.ini (
        notepad config.ini
    ) else (
        echo Configuration file not found. Creating default...
        python -c "from spp_metric_automation import SPPMetricAutomation; SPPMetricAutomation()"
        notepad config.ini
    )
) else if "%choice%"=="5" (
    echo Opening output folder...
    if not exist "Output" mkdir Output
    explorer Output
) else if "%choice%"=="6" (
    echo Goodbye!
    exit
) else (
    echo Invalid choice. Please try again.
    pause
    goto :eof
)

pause
