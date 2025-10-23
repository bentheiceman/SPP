"""
SPP Automation Tool - Build Script
Creates a self-contained executable for team deployment
"""

import subprocess
import sys
import os
import shutil
from pathlib import Path

def install_pyinstaller():
    """Install PyInstaller if not already installed."""
    try:
        import PyInstaller
        print("✓ PyInstaller already installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed successfully")

def create_spec_file():
    """Create PyInstaller spec file with all necessary configurations."""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['spp_fixed_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('spp_metric_automation_fixed.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'pandas',
        'snowflake.connector',
        'openpyxl',
        'snowflake.connector.auth',
        'snowflake.connector.auth.browser',
        'snowflake.connector.auth.external_browser_authenticator',
        'snowflake.connector.auth.oauth',
        'snowflake.connector.auth.keypair',
        'cryptography',
        'cffi',
        'pyarrow',
        'numpy',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'idna',
        'requests',
        'six',
        'pytz',
        'dateutil',
        'et_xmlfile',
        'defusedxml',
        'lxml',
        'PIL',
        'keyring',
        'jwt',
        'asn1crypto',
        'OpenSSL',
        'packaging',
        'tomlkit',
        'sortedcontainers',
        'platformdirs',
        'zipp',
        'importlib_metadata',
        'more_itertools',
        'jaraco',
        'win32ctypes',
        'pywin32_bootstrap',
        'pywin32_system32',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SPP_Automation_Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Will create this
    version='version_info.txt'  # Will create this
)
'''
    
    with open('spp_automation.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    print("✓ Created PyInstaller spec file")

def create_version_info():
    """Create version info file for Windows executable."""
    version_info = '''# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1,0,0,0),
    prodvers=(1,0,0,0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'HD Supply'),
        StringStruct(u'FileDescription', u'SPP Automation Tool'),
        StringStruct(u'FileVersion', u'1.0.0.0'),
        StringStruct(u'InternalName', u'SPP_Automation_Tool'),
        StringStruct(u'LegalCopyright', u'© 2025 HD Supply. All rights reserved.'),
        StringStruct(u'OriginalFilename', u'SPP_Automation_Tool.exe'),
        StringStruct(u'ProductName', u'SPP Automation Tool'),
        StringStruct(u'ProductVersion', u'1.0.0.0')])
      ]), 
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
'''
    
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_info)
    print("✓ Created version info file")

def create_icon():
    """Create a simple icon for the application."""
    try:
        from PIL import Image, ImageDraw
        
        # Create a 256x256 icon with HD Supply colors
        size = (256, 256)
        image = Image.new('RGBA', size, (0, 0, 0, 0))  # Transparent background
        draw = ImageDraw.Draw(image)
        
        # Draw a simple SPP icon - blue background with white text
        draw.rectangle([0, 0, 256, 256], fill='#0066CC', outline='#003366', width=4)
        
        # Add "SPP" text (simplified)
        try:
            # Try to use a larger font if available
            from PIL import ImageFont
            try:
                font = ImageFont.truetype("arial.ttf", 60)
            except:
                font = ImageFont.load_default()
        except:
            font = None
        
        # Draw SPP text
        text = "SPP"
        if font:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        else:
            text_width, text_height = 120, 40  # Approximate default font size
        
        x = (256 - text_width) // 2
        y = (256 - text_height) // 2
        
        draw.text((x, y), text, fill='white', font=font)
        
        # Save as ICO
        image.save('icon.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print("✓ Created application icon")
        
    except ImportError:
        # Fallback: copy any existing icon or create a simple one
        print("⚠ PIL not available, skipping icon creation")
        # Create a minimal icon file
        with open('icon.ico', 'wb') as f:
            # Minimal ICO header (will be ignored if file doesn't exist)
            pass

def create_requirements():
    """Create requirements.txt for reference."""
    requirements = '''# SPP Automation Tool Requirements
pandas>=1.5.0
snowflake-connector-python>=3.0.0
openpyxl>=3.1.0
pyinstaller>=5.0.0
pillow>=9.0.0
cryptography>=3.0.0
pyarrow>=10.0.0
requests>=2.25.0
certifi>=2021.0.0
urllib3>=1.26.0
'''
    
    with open('requirements.txt', 'w', encoding='utf-8') as f:
        f.write(requirements)
    print("✓ Created requirements.txt")

def build_executable():
    """Build the executable using PyInstaller."""
    print("Building executable...")
    print("This may take several minutes...")
    
    try:
        # Run PyInstaller with the spec file
        result = subprocess.run([
            sys.executable, '-m', 'PyInstaller',
            '--clean',
            'spp_automation.spec'
        ], capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Build completed successfully!")
            
            # Check if exe was created
            exe_path = Path('dist/SPP_Automation_Tool.exe')
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✓ Executable created: {exe_path}")
                print(f"✓ File size: {size_mb:.1f} MB")
                return True
            else:
                print("❌ Executable file not found after build")
                return False
        else:
            print("❌ Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            return False
            
    except Exception as e:
        print(f"❌ Build error: {e}")
        return False

def create_deployment_package():
    """Create a deployment package with executable and documentation."""
    print("Creating deployment package...")
    
    # Create deployment folder
    deploy_dir = Path('SPP_Deployment')
    if deploy_dir.exists():
        shutil.rmtree(deploy_dir)
    deploy_dir.mkdir()
    
    # Copy executable
    exe_source = Path('dist/SPP_Automation_Tool.exe')
    if exe_source.exists():
        shutil.copy2(exe_source, deploy_dir / 'SPP_Automation_Tool.exe')
    
    # Create README for deployment
    readme_content = '''# SPP Automation Tool - Deployment Package

## Overview
This package contains the SPP (Supplier Performance Program) Automation Tool, 
a self-contained executable that automates metric reporting and ASN data analysis.

## System Requirements
- Windows 10 or later (64-bit)
- Internet connection for Snowflake authentication
- Excel installed (for viewing generated reports)

## Installation
1. Extract all files to a folder (e.g., C:\\SPP_Tools\\)
2. Run SPP_Automation_Tool.exe

## First Time Setup
1. Launch the application
2. Click "Test Connection" and authenticate with your HD Supply credentials
3. The tool will remember your authentication for future sessions

## Usage
1. **Enter Vendor Numbers**: Input one or more vendor numbers (comma-separated)
2. **Select Report Month**: Choose the fiscal year and month for reporting
3. **Test ASN Query**: (Optional) Test if ASN data exists for your vendors
4. **Run Automation**: Generate the complete report

## Output
- Excel files are saved to the "Output" subfolder
- Files include metric data and ASN data in separate tabs
- Use Ctrl+Shift+M in Excel to run macros for additional processing

## Features
- ✅ Automatic Snowflake authentication
- ✅ Multi-vendor support
- ✅ ASN data integration
- ✅ Excel report generation
- ✅ Macro-enabled templates (when available)
- ✅ Error handling and logging

## Troubleshooting
- Check log files (spp_debug_*.log) for detailed error information
- Ensure network connectivity for Snowflake access
- Contact IT support for authentication issues

## Support
Developer: Ben F. Benjamaa
Manager: Lauren B. Trapani
Date: August 2025
Version: 1.0.0

## File Structure After First Run
```
SPP_Automation_Tool.exe    # Main application
Output/                    # Generated Excel reports
spp_debug_*.log           # Debug and error logs
```

For technical support or feature requests, contact the development team.
'''
    
    with open(deploy_dir / 'README.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    # Create a simple batch file for easy launching
    batch_content = '''@echo off
echo Starting SPP Automation Tool...
echo.
SPP_Automation_Tool.exe
pause
'''
    
    with open(deploy_dir / 'Launch_SPP_Tool.bat', 'w', encoding='utf-8') as f:
        f.write(batch_content)
    
    print(f"✓ Deployment package created in: {deploy_dir.absolute()}")
    
    # Create ZIP file for easy distribution
    try:
        shutil.make_archive('SPP_Deployment_Package', 'zip', 'SPP_Deployment')
        print("✓ Created SPP_Deployment_Package.zip for distribution")
    except Exception as e:
        print(f"⚠ Could not create ZIP file: {e}")

def main():
    """Main build process."""
    print("=" * 60)
    print("SPP Automation Tool - Build Script")
    print("=" * 60)
    
    # Step 1: Install PyInstaller
    install_pyinstaller()
    
    # Step 2: Create build files
    create_requirements()
    create_version_info()
    create_icon()
    create_spec_file()
    
    # Step 3: Build executable
    if build_executable():
        # Step 4: Create deployment package
        create_deployment_package()
        
        print("\n" + "=" * 60)
        print("BUILD COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        print("Deployment files:")
        print("- SPP_Deployment_Package.zip (for distribution)")
        print("- SPP_Deployment/ (extracted files)")
        print("- dist/SPP_Automation_Tool.exe (standalone executable)")
        print("\nThe executable is ready for deployment to your team!")
    else:
        print("\n❌ BUILD FAILED!")
        print("Check the error messages above and try again.")

if __name__ == "__main__":
    main()
