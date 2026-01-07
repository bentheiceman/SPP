"""
Build script for SPP Automation Tool v3.0 with PDH Compliance
Creates a standalone executable for team deployment
"""

import PyInstaller.__main__
import os
import shutil
from pathlib import Path
from datetime import datetime

def build_executable():
    """Build the SPP executable with all dependencies."""
    
    print("=" * 60)
    print("Building SPP Automation Tool v3.0 with PDH Compliance")
    print("=" * 60)
    
    # Clean previous builds
    print("\n[1/4] Cleaning previous builds...")
    for folder in ['build', 'dist', '__pycache__']:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"  ✓ Removed {folder}")
    
    # PyInstaller build configuration
    print("\n[2/4] Building executable with PyInstaller...")
    
    PyInstaller.__main__.run([
        'spp_enhanced_gui.py',
        '--name=SPP_Automation_v3.0_PDH',
        '--onefile',
        '--windowed',
        '--icon=NONE',
        '--add-data=*.json;.',
        '--hidden-import=spp_automation_enhanced',
        '--hidden-import=snowflake.connector.auth',
        '--hidden-import=snowflake.connector.auth.webbrowser',
        '--hidden-import=snowflake.connector.network',
        '--hidden-import=snowflake.connector.cursor',
        '--hidden-import=snowflake.connector.connection',
        '--hidden-import=pandas._libs.tslibs.timedeltas',
        '--hidden-import=pandas._libs.tslibs.nattype',
        '--hidden-import=pandas._libs.tslibs.np_datetime',
        '--hidden-import=openpyxl.cell.read_only',
        '--hidden-import=openpyxl.styles.alignment',
        '--hidden-import=openpyxl.styles.borders',
        '--hidden-import=openpyxl.styles.colors',
        '--collect-all=snowflake.connector',
        '--collect-all=pandas',
        '--collect-all=openpyxl',
        '--noconfirm',
    ])
    
    print("  ✓ Executable built successfully")
    
    # Create deployment folder
    print("\n[3/4] Creating deployment package...")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    deploy_folder = f"SPP_v3.0_PDH_Deployment_{timestamp}"
    
    if os.path.exists(deploy_folder):
        shutil.rmtree(deploy_folder)
    
    os.makedirs(deploy_folder)
    os.makedirs(f"{deploy_folder}/Output", exist_ok=True)
    
    # Copy executable
    exe_source = "dist/SPP_Automation_v3.0_PDH.exe"
    exe_dest = f"{deploy_folder}/SPP_Automation_v3.0_PDH.exe"
    
    if os.path.exists(exe_source):
        shutil.copy2(exe_source, exe_dest)
        print(f"  ✓ Copied executable to {deploy_folder}")
    else:
        print(f"  ✗ ERROR: Executable not found at {exe_source}")
        return False
    
    # Create README
    readme_content = """# SPP Automation Tool v3.0 with PDH Compliance

## What's New in v3.0
✨ **PDH Compliance Tab** - New Tab4 provides Product Data Hub Audit compliance data
   - Pass/Fail metrics for supplier compliance
   - Rolling 28-day compliance window
   - Detailed request and response tracking

## Features
This tool automatically generates SPP reports with 4 comprehensive tabs:

1. **Tab1 - Summary Metrics**: High-level KPIs and percentages
2. **Tab2 - Basic Metrics**: Detailed line-level performance data
3. **Tab3 - ASN Data**: Advance Shipping Notice compliance information
4. **Tab4 - PDH Compliance**: Product Data Hub Audit compliance (NEW!)

## Quick Start Guide

### First Time Setup
1. Double-click `SPP_Automation_v3.0_PDH.exe` to launch
2. Enter your HD Supply email address
3. Click "Authenticate" and complete browser login
4. You're ready to generate reports!

### Generating a Report
1. **Enter Vendor Number(s)**: Type one or more vendor numbers (comma-separated)
2. **Select Report Month**: Use format FY2025-JAN
3. **Enter Date Filter**: Use format YYYYMM (e.g., 202501)
4. Click **"🚀 Generate SPP Report"**
5. Report will be saved to the `Output` folder

### Using Templates (Optional)
- Check "Use Excel Template for Output" if you have a custom template
- Browse to select your template file (.xlsm or .xlsx)
- Template will maintain your custom formatting and macros

## System Requirements
- Windows 10 or later
- Internet connection for Snowflake access
- HD Supply VPN connection (if required)
- Web browser for authentication

## Output Location
All generated reports are saved in the `Output` folder in the same directory as the executable.

## Troubleshooting

### Authentication Issues
- Ensure you're connected to VPN (if required)
- Complete the browser authentication within 2 minutes
- Check that your email address is correct

### No Data Found
- Verify the vendor number exists in the system
- Check that the report month format is correct (FY2025-JAN)
- Ensure the date filter matches the report month

### Error During Report Generation
- Check the Activity Log in the application for details
- Ensure you have write permissions to the Output folder
- Contact IT support if issues persist

## Support
Developer: Ben F. Benjamaa
Manager: Lauren B. Trapani
Team: HD Supply Chain Excellence

For questions or issues, please contact your team lead or IT support.

---
Version: 3.0
Build Date: {build_date}
Includes: PDH Compliance Tracking
"""
    
    with open(f"{deploy_folder}/README.txt", 'w') as f:
        f.write(readme_content.format(build_date=datetime.now().strftime('%B %d, %Y')))
    
    print("  ✓ Created README.txt")
    
    # Create Quick Start guide
    quickstart = """QUICK START - SPP Automation Tool v3.0

═══════════════════════════════════════════════════════════════

1. LAUNCH THE APPLICATION
   - Double-click: SPP_Automation_v3.0_PDH.exe

2. AUTHENTICATE (First time only)
   - Enter your @hdsupply.com email
   - Click "Authenticate"
   - Complete browser login

3. GENERATE REPORT
   - Vendor Number: 13479 (or your vendor)
   - Report Month: FY2025-JAN
   - Date Filter: 202501
   - Click: 🚀 Generate SPP Report

4. FIND YOUR REPORT
   - Check the "Output" folder
   - File name format: VENDOR# - NAME - MONTH.xlsx

═══════════════════════════════════════════════════════════════

NEW IN v3.0: PDH Compliance Tab!
Your reports now include Tab4 with Product Data Hub compliance
data - helping you track supplier response times and compliance.

Need help? See README.txt for full documentation.
"""
    
    with open(f"{deploy_folder}/QUICK_START.txt", 'w') as f:
        f.write(quickstart)
    
    print("  ✓ Created QUICK_START.txt")
    
    # Create version info
    version_info = f"""SPP Automation Tool - Version Information

Version: 3.0
Build Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Build Type: Standalone Executable

New Features:
- PDH Compliance Tab (Tab4)
- Rolling 28-day compliance tracking
- Enhanced vendor compliance reporting

Tabs Included:
1. Summary Metrics
2. Basic Metrics
3. ASN Data
4. PDH Compliance (NEW!)

Developer: Ben F. Benjamaa
Manager: Lauren B. Trapani
Organization: HD Supply Chain Excellence
"""
    
    with open(f"{deploy_folder}/VERSION_INFO.txt", 'w') as f:
        f.write(version_info)
    
    print("  ✓ Created VERSION_INFO.txt")
    
    # Final summary
    print("\n[4/4] Deployment package created successfully!")
    print(f"\n{'=' * 60}")
    print(f"Deployment folder: {deploy_folder}")
    print(f"Executable: SPP_Automation_v3.0_PDH.exe")
    print(f"Size: {os.path.getsize(exe_dest) / (1024*1024):.1f} MB")
    print(f"{'=' * 60}")
    print("\n✅ BUILD COMPLETE!")
    print(f"\nTo deploy:")
    print(f"1. Zip the '{deploy_folder}' folder")
    print(f"2. Share with your team")
    print(f"3. Instruct them to extract and run the .exe file")
    print(f"\n⚡ The tool now includes PDH Compliance tracking in Tab4!")
    
    return True

if __name__ == "__main__":
    try:
        success = build_executable()
        if success:
            print("\n✨ Ready for team deployment!")
        else:
            print("\n❌ Build failed. Check errors above.")
    except Exception as e:
        print(f"\n❌ Build error: {e}")
        import traceback
        traceback.print_exc()
