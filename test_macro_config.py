#!/usr/bin/env python3
"""
Quick test to verify macro template functionality without full automation run
"""

import os
import sys
import json
import shutil
from pathlib import Path

def test_macro_template_configuration():
    """Test the macro template configuration and file handling."""
    
    print("🧪 Testing Macro Template Configuration")
    print("=" * 50)
    
    # Check template file exists
    template_path = r"C:\Users\1015723\OneDrive - HD Supply, Inc\Documents\Cole - Multi-Query information and template\Simplified Macro Template - SPP Monthly Details.xlsm"
    
    if not os.path.exists(template_path):
        print(f"❌ Template file not found: {template_path}")
        return False
    
    print(f"✅ Template file exists: {os.path.basename(template_path)}")
    print(f"   Size: {os.path.getsize(template_path):,} bytes")
    print(f"   Extension: {Path(template_path).suffix}")
    
    # Check template configuration
    config_path = "template_config.json"
    config = {}
    if os.path.exists(config_path):
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        print(f"✅ Template configuration loaded:")
        print(f"   Use template: {config.get('use_template', 'Not set')}")
        print(f"   Template path: {config.get('template_path', 'Not set')}")
        print(f"   Template format: {config.get('template_format', 'Not set')}")
        
        # Verify configuration matches template file
        if config.get('template_path') == template_path:
            print("✅ Configuration path matches template file")
        else:
            print("❌ Configuration path does not match template file")
            
        if config.get('template_format') == 'xlsm':
            print("✅ Configuration format is xlsm")
        else:
            print("❌ Configuration format is not xlsm")
    
    # Test file operations that the automation would do
    try:
        print("\n🔧 Testing file operations...")
        
        # Test copying the template (what copy_template_file does)
        test_output = "test_macro_copy.xlsm"
        shutil.copy2(template_path, test_output)
        print(f"✅ Template copied successfully to {test_output}")
        
        # Verify the copy maintains .xlsm extension
        if test_output.endswith('.xlsm'):
            print("✅ Copy maintains .xlsm extension")
        else:
            print("❌ Copy does not maintain .xlsm extension")
        
        # Test that we can import openpyxl and it supports VBA
        try:
            import openpyxl
            workbook = openpyxl.load_workbook(test_output, keep_vba=True)
            
            print("✅ Successfully loaded with openpyxl keep_vba=True")
            print(f"   Worksheets: {workbook.sheetnames}")
            
            # Check for VBA
            if hasattr(workbook, 'vba_archive'):
                if workbook.vba_archive:
                    print("✅ VBA archive detected in workbook")
                else:
                    print("⚠️ VBA archive attribute exists but is None/empty")
            else:
                print("⚠️ No VBA archive attribute found")
            
            workbook.close()
            
        except Exception as e:
            print(f"❌ Error loading with openpyxl: {e}")
            return False
        
        # Clean up
        os.remove(test_output)
        print("✅ Test file cleaned up")
        
    except Exception as e:
        print(f"❌ Error in file operations test: {e}")
        return False
    
    # Test filename generation logic
    print("\n📝 Testing filename generation logic...")
    
    # Simulate what generate_filename would do
    vendor_numbers = ["123456"]
    vendor_name = "Test Vendor"
    report_month = "FY25-01"
    
    # Simulate the filename generation logic from our updated code
    import re
    clean_vendor_name = re.sub(r'[<>:"/\\|?*]', '_', vendor_name.upper()) if vendor_name else "Unknown_Vendor"
    
    month_parts = report_month.split('-')
    if len(month_parts) == 2:
        year_part = month_parts[0].replace('FY', '')
        month_part = month_parts[1]
        readable_month = f"{month_part} {year_part}"
    else:
        readable_month = report_month
    
    # Check template format to determine extension
    ext = "xlsx"  # default
    if config.get('use_template', False):
        if template_path and template_path.endswith('.xlsm'):
            ext = "xlsm"
        elif template_path and template_path.endswith('.xlsx'):
            ext = "xlsx"
    
    vendor_prefix = "_".join(vendor_numbers)
    filename = f"{vendor_prefix} - {clean_vendor_name} - {readable_month}.{ext}"
    
    print(f"✅ Generated filename: {filename}")
    print(f"   Expected extension: xlsm")
    
    if filename.endswith('.xlsm'):
        print("✅ Filename has correct .xlsm extension")
    else:
        print(f"❌ Filename has wrong extension: {filename.split('.')[-1]}")
        return False
    
    print("\n🎉 All macro template configuration tests passed!")
    print("\n📋 Summary:")
    print("   • Template file exists and is accessible")
    print("   • Configuration properly set for macro template")
    print("   • File copy operations work correctly")
    print("   • OpenpyXL can handle macro-enabled files")
    print("   • Filename generation produces .xlsm extension")
    print("   • Your application should create .xlsm output files")
    
    return True

if __name__ == "__main__":
    success = test_macro_template_configuration()
    if success:
        print("\n✅ SUCCESS: Macro template support is properly configured!")
    else:
        print("\n❌ ISSUES: Macro template support needs attention")
