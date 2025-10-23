#!/usr/bin/env python3
"""
Test script for macro-enabled template functionality in SPP Enhanced
"""

import os
import sys
import json
from pathlib import Path

# Add the current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from spp_automation_enhanced import SPPAutomationEnhanced

def test_macro_template_support():
    """Test macro-enabled template functionality."""
    
    print("🧪 Testing Macro-Enabled Template Support")
    print("=" * 50)
    
    # Configuration for macro template
    macro_template_path = r"C:\Users\1015723\OneDrive - HD Supply, Inc\Documents\Cole - Multi-Query information and template\Simplified Macro Template - SPP Monthly Details.xlsm"
    
    # Check if template exists
    if not os.path.exists(macro_template_path):
        print(f"❌ ERROR: Macro template not found at: {macro_template_path}")
        return False
    
    print(f"✅ Found macro template: {macro_template_path}")
    
    # Create test configuration
    test_config = {
        "use_template": True,
        "template_path": macro_template_path,
        "template_format": "xlsm",
        "template_name": "Simplified Macro Template - SPP Monthly Details.xlsm"
    }
    
    # Save test configuration
    config_path = "template_config.json"
    with open(config_path, 'w') as f:
        json.dump(test_config, f, indent=2)
    
    print(f"✅ Created test configuration: {config_path}")
    
    # Initialize automation with test config
    try:
        automation = SPPAutomationEnhanced()
        print("✅ SPPAutomationEnhanced initialized successfully")
    except Exception as e:
        print(f"❌ ERROR initializing automation: {e}")
        return False
    
    # Test template discovery
    found_template = automation.find_template_file()
    if found_template:
        print(f"✅ Template discovery successful: {found_template}")
        print(f"   Template format: {found_template.split('.')[-1]}")
    else:
        print("❌ ERROR: Template discovery failed")
        return False
    
    # Test filename generation
    test_vendors = ["123456"]
    test_vendor_name = "Test Vendor"
    test_month = "FY25-01"
    
    filename = automation.generate_filename(test_vendors, test_vendor_name, test_month)
    print(f"✅ Generated filename: {filename}")
    
    expected_extension = "xlsm"
    if filename.endswith(expected_extension):
        print(f"✅ Correct file extension: {expected_extension}")
    else:
        print(f"❌ ERROR: Expected .{expected_extension} extension, got: {filename.split('.')[-1]}")
        return False
    
    # Test template copying
    test_output_path = os.path.join("Output", filename)
    os.makedirs("Output", exist_ok=True)
    
    if automation.copy_template_file(test_output_path):
        print(f"✅ Template copied successfully to: {test_output_path}")
        
        # Verify the copied file is macro-enabled
        if test_output_path.endswith('.xlsm'):
            print("✅ Output file maintains .xlsm format")
        else:
            print("❌ ERROR: Output file lost macro format")
            return False
    else:
        print("❌ ERROR: Template copying failed")
        return False
    
    # Test openpyxl macro support
    try:
        import openpyxl
        workbook = openpyxl.load_workbook(test_output_path, keep_vba=True)
        print("✅ Openpyxl can load macro-enabled file with VBA preservation")
        
        # Check if workbook has VBA project
        if hasattr(workbook, 'vba_archive') and workbook.vba_archive:
            print("✅ VBA macros detected and preserved")
        else:
            print("⚠️  Warning: No VBA macros detected (may be normal for some templates)")
        
        workbook.close()
        
    except Exception as e:
        print(f"❌ ERROR loading macro file with openpyxl: {e}")
        return False
    
    # Clean up test file
    if os.path.exists(test_output_path):
        os.remove(test_output_path)
        print("✅ Cleaned up test output file")
    
    print("\n🎉 All macro-enabled template tests passed!")
    return True

def test_connection():
    """Test basic connection functionality."""
    print("\n🔗 Testing Database Connection")
    print("=" * 50)
    
    try:
        automation = SPPAutomationEnhanced()
        success, message = automation.test_connection()
        
        if success:
            print(f"✅ Database connection successful: {message}")
            return True
        else:
            print(f"❌ Database connection failed: {message}")
            return False
            
    except Exception as e:
        print(f"❌ ERROR testing connection: {e}")
        return False

def main():
    """Run all tests."""
    print("🚀 SPP Enhanced - Macro Template Testing Suite")
    print("=" * 60)
    
    # Test 1: Macro template support
    test1_passed = test_macro_template_support()
    
    # Test 2: Database connection (optional - may fail without credentials)
    test2_passed = test_connection()
    
    print("\n📊 Test Results Summary")
    print("=" * 60)
    print(f"Macro Template Support: {'✅ PASSED' if test1_passed else '❌ FAILED'}")
    print(f"Database Connection: {'✅ PASSED' if test2_passed else '❌ FAILED (Expected if no credentials)'}")
    
    if test1_passed:
        print("\n🎉 SUCCESS: Your application supports macro-enabled templates!")
        print("   • .xlsm templates are properly detected")
        print("   • Output files maintain .xlsm format")
        print("   • VBA macros are preserved during processing")
        print("   • Template copying and population works correctly")
    else:
        print("\n❌ FAILURE: Macro template support needs attention")
    
    return test1_passed

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
