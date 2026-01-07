# SPP v3.0 Easy Build Script
# Run this from PowerShell (not Python REPL)

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "SPP v3.0 with PDH Compliance Builder" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if we're in the right directory
if (!(Test-Path "spp_enhanced_gui.py")) {
    Write-Host "ERROR: Not in the SPP directory!" -ForegroundColor Red
    Write-Host "Please navigate to: C:\Users\1015723\Downloads\SPP" -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "✅ Found SPP files" -ForegroundColor Green

# Check Python installation
Write-Host ""
Write-Host "Checking Python..." -ForegroundColor Yellow
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Python not found!" -ForegroundColor Red
    Write-Host "Please install Python from python.org" -ForegroundColor Yellow
    pause
    exit 1
}
Write-Host "✅ Python found: $pythonVersion" -ForegroundColor Green

# Install PyInstaller if needed
Write-Host ""
Write-Host "Checking PyInstaller..." -ForegroundColor Yellow
$pyinstallerCheck = pip show pyinstaller 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
    pip install pyinstaller
}
Write-Host "✅ PyInstaller ready" -ForegroundColor Green

# Run the build script
Write-Host ""
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "Building SPP v3.0 Executable..." -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

python build_spp_v3.py

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "====================================" -ForegroundColor Green
    Write-Host "✅ BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Your deployment package is ready in:" -ForegroundColor Yellow
    Get-ChildItem -Directory -Filter "SPP_v3.0_PDH_Deployment_*" | ForEach-Object {
        Write-Host "  📦 $($_.FullName)" -ForegroundColor Cyan
    }
    Write-Host ""
    Write-Host "Share this folder with your team!" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "====================================" -ForegroundColor Red
    Write-Host "❌ BUILD FAILED" -ForegroundColor Red
    Write-Host "====================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Check the error messages above." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
