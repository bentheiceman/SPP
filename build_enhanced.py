#!/usr/bin/env python3
"""
Enhanced Build Script for SPP Automation Tool v2.2
Creates a self-contained executable with all dependencies and deployment package.
"""

import os
import sys
import subprocess
import shutil
import zipfile
from pathlib import Path
from datetime import datetime
import json
import glob

def install_pyinstaller():
    """Install PyInstaller if not already installed."""
    try:
        import PyInstaller
        print("✓ PyInstaller is already installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed successfully")

def install_requirements():
    """Install project requirements to ensure PyInstaller can bundle modules."""
    req_file = Path("requirements.txt")
    if not req_file.exists():
        print("⚠ requirements.txt not found; skipping dependency install")
        return
    print("Installing project dependencies from requirements.txt ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req_file)])
        print("✓ Dependencies installed")
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to install dependencies: {e}")
        raise

def cleanup_old_artifacts():
    """Remove previous build outputs, deployment folders, and old zip packages."""
    print("Cleaning old build artifacts...")
    dirs_to_remove = [
        Path('build'),
        Path('dist'),
        Path('SPP_Enhanced_Deployment'),
        Path('SPP_Enhanced_v2.2_Team_Deployment'),
    ]
    for d in dirs_to_remove:
        try:
            if d.exists():
                shutil.rmtree(d)
                print(f"✓ Removed {d}")
        except Exception as e:
            print(f"⚠ Could not remove {d}: {e}")
    # Remove old zip packages
    for z in glob.glob('SPP_Automation_Tool_Enhanced_*.zip'):
        try:
            os.remove(z)
            print(f"✓ Removed {z}")
        except Exception as e:
            print(f"⚠ Could not remove {z}: {e}")

def create_enhanced_spec_file():
    """Create enhanced PyInstaller spec file."""
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

# Enhanced spec file for SPP Automation Tool
block_cipher = None

# Define the main GUI entry point
gui_analysis = Analysis(
    ['spp_enhanced_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('*.sql', '.'),
        ('*.json', '.'),
    ],
    hiddenimports=[
        'spp_automation_enhanced',
        'snowflake.connector.auth',
        'snowflake.connector.auth.webbrowser',
        'snowflake.connector.network',
        'snowflake.connector.cursor',
        'snowflake.connector.connection',
        'snowflake.connector.constants',
        'snowflake.connector.errorcode',
        'snowflake.connector.errors',
        'snowflake.connector.util_text',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.parsing',
        'pandas._libs.tslibs.period',
        'pandas._libs.tslibs.strptime',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.tslibs.tzconversion',
        'openpyxl.cell.read_only',
        'openpyxl.chart.marker',
        'openpyxl.chart.text',
        'openpyxl.comments.comment_sheet',
        'openpyxl.compat.strings',
        'openpyxl.descriptors.serialisable',
        'openpyxl.formatting.formatting',
        'openpyxl.pivot.fields',
        'openpyxl.styles.alignment',
        'openpyxl.styles.borders',
        'openpyxl.styles.colors',
        'openpyxl.styles.differential',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.styles.named_styles',
        'openpyxl.styles.numbers',
        'openpyxl.styles.protection',
        'openpyxl.utils.dataframe',
        'openpyxl.workbook.defined_name',
        'openpyxl.workbook.smart_tags',
        'openpyxl.worksheet.datavalidation',
        'openpyxl.worksheet.hyperlink',
        'openpyxl.worksheet.pagebreak',
        'openpyxl.worksheet.related',
        'openpyxl.xml.constants',
        'pkg_resources.py2_warn',
        'pkg_resources.markers',
        'jwt.algorithms',
        'cryptography.hazmat.backends.openssl',
        'cryptography.hazmat.backends.openssl.rsa',
        'cryptography.hazmat.backends.openssl.dsa',
        'cryptography.hazmat.backends.openssl.ec',
        'keyring.backends',
        'defusedxml.expatreader',
        'defusedxml.expatbuilder',
        'defusedxml.sax',
        'defusedxml.minidom',
        'defusedxml.pulldom',
        'defusedxml.xmlrpc',
        'cffi.backend_ctypes',
        'idna.core',
        'idna.idnadata',
        'idna.intranges',
        'charset_normalizer.md',
        'urllib3.util.connection',
        'urllib3.util.url',
        'urllib3.util.retry',
        'urllib3.util.timeout',
        'urllib3.contrib.pyopenssl',
        'urllib3.packages.six',
        'packaging.version',
        'packaging.specifiers',
        'packaging.requirements',
        'packaging.markers',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
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

gui_pyz = PYZ(gui_analysis.pure, gui_analysis.zipped_data, cipher=block_cipher)

gui_exe = EXE(
    gui_pyz,
    gui_analysis.scripts,
    gui_analysis.binaries,
    gui_analysis.zipfiles,
    gui_analysis.datas,
    [],
    name='SPP_Automation_Tool_Enhanced',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Hide console window for GUI
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    icon='hd_supply_icon.ico'
)
"""
    
    with open('spp_enhanced.spec', 'w') as f:
        f.write(spec_content)
    
    print("✓ Enhanced spec file created: spp_enhanced.spec")

# Utility cleanups
def _remove_empty_dirs(root: Path):
    """Recursively remove empty directories under root, excluding venv and .git."""
    exclude = {'.venv', '.git'}
    for current_root, dirs, files in os.walk(root, topdown=False):
        base = os.path.basename(current_root)
        if base in exclude:
            continue
        if not dirs and not files:
            try:
                os.rmdir(current_root)
                print(f"✓ Removed empty folder {current_root}")
            except Exception:
                pass

def cleanup_old_logs(days: int = 14):
    """Remove log files older than the given number of days."""
    import time
    cutoff = time.time() - days * 86400
    patterns = [
        'spp_gui_*.log',
        'spp_automation_*.log',
        'spp_gui*.log',
    ]
    for pattern in patterns:
        for path in glob.glob(pattern):
            try:
                if os.path.getmtime(path) < cutoff:
                    os.remove(path)
                    print(f"✓ Removed old log {path}")
            except Exception:
                pass

def create_version_info():
    """Create version info file for the executable (UTF-8 encoded)."""
    version_info = """# UTF-8
#
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(2,0,0,0),
    prodvers=(2,0,0,0),
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'HD Supply Chain Excellence'),
        StringStruct(u'FileDescription', u'SPP Metric Automation Tool Enhanced'),
        StringStruct(u'FileVersion', u'2.2.0.0'),
        StringStruct(u'InternalName', u'SPP_Automation_Tool_Enhanced'),
        StringStruct(u'LegalCopyright', u'Copyright (c) 2025 HD Supply. All rights reserved.'),
        StringStruct(u'OriginalFilename', u'SPP_Automation_Tool_Enhanced.exe'),
        StringStruct(u'ProductName', u'SPP Metric Automation Tool Enhanced'),
        StringStruct(u'ProductVersion', u'2.2.0.0'),
        StringStruct(u'Developer', u'Ben F. Benjamaa'),
        StringStruct(u'Manager', u'Lauren B. Trapani')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"""
    # Write explicitly in UTF-8 so PyInstaller can decode it
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_info)
    
    print("✓ Version info file created")

def create_icon():
    """Create HD Supply icon for the application."""
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # Create a 256x256 image with black background
        img = Image.new('RGBA', (256, 256), (0, 0, 0, 255))
        draw = ImageDraw.Draw(img)
        
        # Draw HD Supply style design
        # Yellow border
        draw.rectangle([10, 10, 246, 246], outline=(255, 255, 0, 255), width=8)
        
        # HD text
        try:
            font_large = ImageFont.truetype("arial.ttf", 48)
            font_medium = ImageFont.truetype("arial.ttf", 24)
            font_small = ImageFont.truetype("arial.ttf", 18)
        except:
            try:
                font_large = ImageFont.load_default()
                font_medium = ImageFont.load_default()
                font_small = ImageFont.load_default()
            except:
                font_large = font_medium = font_small = None
        
        if font_large:
            # Calculate text positioning
            hd_text = "HD"
            bbox = draw.textbbox((0, 0), hd_text, font=font_large)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            x = (256 - text_width) // 2
            y = 60
            
            draw.text((x, y), hd_text, fill=(255, 255, 0, 255), font=font_large)
            
            # SPP text
            spp_text = "SPP"
            bbox = draw.textbbox((0, 0), spp_text, font=font_medium)
            text_width = bbox[2] - bbox[0]
            x = (256 - text_width) // 2
            y = 120
            
            draw.text((x, y), spp_text, fill=(0, 255, 0, 255), font=font_medium)
            
            # Automation text
            auto_text = "AUTOMATION"
            bbox = draw.textbbox((0, 0), auto_text, font=font_small)
            text_width = bbox[2] - bbox[0]
            x = (256 - text_width) // 2
            y = 160
            
            draw.text((x, y), auto_text, fill=(255, 255, 255, 255), font=font_small)
        else:
            # Fallback without text - just colored rectangles
            draw.rectangle([40, 80, 216, 120], fill=(255, 255, 0, 255))
            draw.rectangle([60, 130, 196, 150], fill=(0, 255, 0, 255))
        
        # Save as ICO file
        img.save('hd_supply_icon.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print("✓ HD Supply icon created")
        
        return True
        
    except ImportError:
        print("⚠ Pillow not installed, using default icon")
        return False
    except Exception as e:
        print(f"⚠ Could not create custom icon: {e}")
        return False

def build_executable():
    """Build the executable using PyInstaller."""
    try:
        print("Building SPP Automation Tool Enhanced executable...")
        
        # Clean previous builds
        if os.path.exists('build'):
            shutil.rmtree('build')
        if os.path.exists('dist'):
            shutil.rmtree('dist')
        
        # Build using spec file
        cmd = [sys.executable, "-m", "PyInstaller", "--clean", "spp_enhanced.spec"]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            exe_path = Path("dist") / "SPP_Automation_Tool_Enhanced.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✓ Executable built successfully: {exe_path} ({size_mb:.1f} MB)")
                return str(exe_path)
            else:
                print("✗ Executable file not found after build")
                return None
        else:
            print(f"✗ Build failed:")
            print(result.stdout)
            print(result.stderr)
            return None
            
    except Exception as e:
        print(f"✗ Build error: {e}")
        return None

def create_enhanced_deployment_package():
    """Create comprehensive deployment package."""
    try:
        print("Creating enhanced deployment package...")
        
        # Create deployment directory
        # Standard team deployment folder name
        deployment_dir = Path("SPP_Enhanced_v2.2_Team_Deployment")
        if deployment_dir.exists():
            shutil.rmtree(deployment_dir)
        deployment_dir.mkdir()
        
        # Copy executable
        exe_source = Path("dist") / "SPP_Automation_Tool_Enhanced.exe"
        exe_dest = deployment_dir / "SPP_Automation_Tool_Enhanced.exe"
        if exe_source.exists():
            shutil.copy2(exe_source, exe_dest)
            print(f"✓ Copied executable to deployment package")
        else:
            print("✗ Executable not found for deployment")
            return None
        
        # Create sample template configuration
        template_config = {
            "template_path": "",
            "use_template": False,
            "template_name": "SPP_Template.xlsm",
            "output_format": "xlsx",
            "search_paths": [
                "C:\\\\Users\\\\%USERNAME%\\\\OneDrive\\\\Documents",
                "C:\\\\Users\\\\%USERNAME%\\\\Documents",
                "."
            ],
            "last_updated": datetime.now().isoformat()
        }
        
        with open(deployment_dir / "template_config.json", 'w') as f:
            json.dump(template_config, f, indent=2)
        
        # Create comprehensive README
        readme_content = """# SPP Metric Automation Tool Enhanced v2.2
## HD Supply Chain Excellence

### Overview
The SPP Metric Automation Tool Enhanced provides advanced reporting capabilities with user-configurable Excel templates, comprehensive error handling, and an intuitive interface.

### New Features in v2.2
- **Multi-Statement SQL Support**: Handles USE DATABASE + SELECT queries seamlessly
- **Vendor Part Number Tracking**: Added manufacturer part numbers to all detailed tabs
- **Enhanced Compliance Tracking**: Compliant/Non-Compliant indicators across all tabs
- **User-Configurable Templates**: Link your own Excel template files for consistent formatting
- **Template Auto-Discovery**: Automatic search in common locations (OneDrive, Documents)
- **Dual Output Modes**: Standard Excel (.xlsx) or Macro-Enabled (.xlsm) formats
- **Enhanced Error Handling**: Graceful fallbacks when templates are unavailable
- **Improved UI**: Modern, branded interface with comprehensive configuration options
- **Background Processing**: Non-blocking operations with progress indicators

### Quick Start
1. **Launch**: Double-click `SPP_Automation_Tool_Enhanced.exe`
2. **Authenticate**: Enter your HD Supply email and click "Authenticate"
3. **Configure Template** (Optional):
   - Check "Use Excel Template for Output"
   - Browse for your template file or use auto-discovery
   - Choose output format (.xlsx or .xlsm)
4. **Enter Parameters**:
   - Vendor Numbers (comma-separated)
   - Report Month (e.g., FY2025-APR)
   - Date Filter (YYYYMM format)
5. **Generate Report**: Click "🚀 Generate SPP Report"

### Template Configuration
The tool supports flexible template configuration:

#### Template Discovery
The tool searches for templates in these locations (in order):
1. User-specified path (if configured)
2. OneDrive Documents folder
3. Local Documents folder
4. Application directory
5. Current working directory

#### Template Requirements
- Must be Excel format (.xlsx or .xlsm)
- Should contain sheets named "Tab1_Basic_Metrics" and "Tab2_ASN_Data"
- Headers will be preserved, data populated starting from row 2

#### Fallback Behavior
If no template is found, the tool automatically creates a standard Excel file with all data in separate tabs.

### System Requirements
- Windows 10/11 (64-bit)
- 4GB RAM minimum
- Internet connection for Snowflake authentication
- HD Supply network access

### Configuration Files
- `template_config.json`: Template settings and paths
- Application logs: Automatic logging with timestamps
- Output folder: Created automatically in same directory

### Troubleshooting

#### Authentication Issues
- Ensure you're on HD Supply network
- Check email format is correct
- Try clearing browser cookies if authentication fails

#### Template Issues
- Verify template file path is accessible
- Check file permissions
- Tool will fallback to standard Excel if template unavailable

#### Data Issues
- Verify vendor numbers are correct
- Check date formats (FY2025-APR, 202507)
- Use "Test Query" button to validate parameters

#### Performance
- Large vendor lists may take several minutes
- Use date filters to limit data volume
- Monitor activity log for progress updates

### Support Information
- **Developer**: Ben F. Benjamaa
- **Manager**: Lauren B. Trapani
- **Department**: HD Supply Chain Excellence
- **Version**: 2.2.0 Enhanced
- **Build Date**: """ + datetime.now().strftime("%Y-%m-%d") + """

### Change Log
#### Version 2.2.0 (Enhanced)
- Multi-statement SQL execution support (USE DATABASE + SELECT)
- Vendor part number tracking in all detailed tabs
- Enhanced compliance indicators
- Fixed statement count mismatch error

#### Version 2.0.0 (Enhanced)
- Added user-configurable template system
- Implemented template auto-discovery
- Enhanced error handling and fallback mechanisms
- Improved UI with template configuration options
- Added comprehensive logging
- Optimized performance and memory usage
- Enhanced deployment packaging

#### Version 1.0.0
- Initial release
- Basic Snowflake connectivity
- Standard Excel output
- Simple GUI interface

For technical support or feature requests, contact the development team.
"""
        
        with open(deployment_dir / "README.md", 'w', encoding='utf-8') as f:
            f.write(readme_content)
        
        # Create launcher script
        launcher_content = """@echo off
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
"""
        
        with open(deployment_dir / "Launch_SPP_Enhanced.bat", 'w', encoding='utf-8') as f:
            f.write(launcher_content)
        
        # Create template guide
        template_guide = """# Excel Template Configuration Guide
## SPP Automation Tool Enhanced

### Template Structure
Your Excel template should contain the following sheets:
- **Tab1_Basic_Metrics**: For basic performance metrics
- **Tab2_ASN_Data**: For ASN (Advance Ship Notice) data

### Required Columns for Tab1_Basic_Metrics:
- VENDOR_NUMBER
- VENDOR_NAME
- VENDOR (Combined format)
- PO_NUMBER
- USN
- ITEM_DESCRIPTION
- DATE_ORIG_ORDERED
- DATE_ORIG_PROMISED
- DATE_FIRST_RECEIVED
- WAREHOUSE_NUM
- WAREHOUSE_NAME
- METRIC
- RPT_MONTH
- And other metric columns...

### Required Columns for Tab2_ASN_Data:
- VENDOR_NUMBER
- PO_NUMBER
- PO_LINE_ITEM
- MATERIAL_NUMBER
- DELIVERY_DATE
- QUANTITY
- UNIT_OF_MEASURE
- NET_PRICE
- NET_VALUE
- And other ASN columns...

### Template Best Practices:
1. Use consistent formatting (fonts, colors, borders)
2. Include company branding in headers
3. Pre-format number and date columns
4. Leave row 1 for headers, data starts at row 2
5. Use conditional formatting for key metrics
6. Include summary sheets if needed

### File Naming:
- Recommended: SPP_Template.xlsm (for macro support)
- Alternative: SPP_Template.xlsx (standard format)

### Macro Support:
- Save as .xlsm for macro-enabled features
- Include VBA code for automatic calculations
- Tool will preserve macros when using templates

### Storage Locations:
Place your template in any of these locations:
1. OneDrive Documents folder (automatically detected)
2. Local Documents folder (automatically detected)  
3. Same folder as the application
4. Custom path (configured in application)

The tool will find your template automatically or you can specify a custom path using the "Browse" button in the template configuration section.
"""
        
        with open(deployment_dir / "Template_Guide.md", 'w', encoding='utf-8') as f:
            f.write(template_guide)
        
        # Create sample config file
        with open(deployment_dir / "config.ini", 'w', encoding='utf-8') as f:
            f.write("""[snowflake]
# Snowflake connection settings are managed automatically
# No manual configuration required

[application]
# Application settings
log_level = INFO
output_directory = Output
auto_open_results = true

[template]
# Template settings are managed through the GUI
# This file is for reference only
""")
        
        # Create output directory
        (deployment_dir / "Output").mkdir()
        with open(deployment_dir / "Output" / "README.txt", 'w', encoding='utf-8') as f:
            f.write("Generated SPP reports will be saved in this folder.\n")
        
        # Create ZIP package
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"SPP_Automation_Tool_Enhanced_v2.2_{timestamp}.zip"

        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
            for root, dirs, files in os.walk(deployment_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, deployment_dir.parent)
                    zipf.write(file_path, arcname)

        zip_size = os.path.getsize(zip_name) / (1024 * 1024)

        print(f"✓ Deployment package created: {deployment_dir}")
        print(f"✓ ZIP package created: {zip_name} ({zip_size:.1f} MB)")

        return str(deployment_dir), zip_name

    except Exception as e:
        print(f"✗ Deployment package creation failed: {e}")
        return None, None

def main():
    """Main build process."""
    print("=" * 60)
    print("SPP Automation Tool Enhanced - Build Script v2.2")
    print("HD Supply Chain Excellence")
    print("=" * 60)
    
    try:
        # Check required files
        required_files = ["spp_enhanced_gui.py", "spp_automation_enhanced.py", "requirements.txt"]
        missing_files = [f for f in required_files if not os.path.exists(f)]
        
        if missing_files:
            print(f"✗ Missing required files: {missing_files}")
            return False

        # Install dependencies
        print("\n1. Installing dependencies...")
        install_pyinstaller()
        install_requirements()

        # Cleanup any previous artifacts
        print("\n1b. Cleaning up old artifacts...")
        cleanup_old_artifacts()
        cleanup_old_logs(14)
        _remove_empty_dirs(Path.cwd())
        
        # Create build files
        print("\n2. Creating build configuration...")
        create_version_info()
        create_icon()
        create_enhanced_spec_file()
        
        # Build executable
        print("\n3. Building executable...")
        exe_path = build_executable()
        
        if not exe_path:
            print("✗ Build failed")
            return False
        
        # Create deployment package
        print("\n4. Creating deployment package...")
        deployment_result = create_enhanced_deployment_package()
        
        if not deployment_result or not deployment_result[0]:
            print("✗ Deployment package creation failed")
            return False
        
        deployment_dir, zip_file = deployment_result
        
        # Success summary
        print("\n" + "=" * 60)
        print("BUILD COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        print(f"✓ Executable: {exe_path}")
        print(f"✓ Deployment: {deployment_dir}")
        print(f"✓ ZIP Package: {zip_file}")
        print("\nReady for distribution to your team!")
        print("=" * 60)
        
        return True
        
    except Exception as e:
        print(f"\n✗ BUILD FAILED: {e}")
        return False

if __name__ == "__main__":
    success = main()
    
    if success:
        input("\nPress Enter to exit...")
    else:
        input("\nBuild failed. Press Enter to exit...")
    
    sys.exit(0 if success else 1)
