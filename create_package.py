"""
Package Creation Script
Creates a zip file of the SPP Automation package for distribution.
"""

import zipfile
import os
from datetime import datetime

def create_package_zip():
    """Create a zip file of the automation package."""
    
    package_dir = "SPP_Automation_Package"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_filename = f"SPP_Automation_v1.0_{timestamp}.zip"
    
    print(f"Creating package: {zip_filename}")
    print("=" * 50)
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(package_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arc_path = os.path.relpath(file_path, os.path.dirname(package_dir))
                zipf.write(file_path, arc_path)
                print(f"Added: {arc_path}")
    
    file_size = os.path.getsize(zip_filename) / (1024 * 1024)  # MB
    print(f"\n✅ Package created: {zip_filename}")
    print(f"📦 Size: {file_size:.2f} MB")
    print(f"📍 Location: {os.path.abspath(zip_filename)}")
    
    return zip_filename

if __name__ == "__main__":
    create_package_zip()
