#!/usr/bin/env python3
"""
Simple launcher for SPP Automation Enhanced
Ensures proper error handling and logging.
"""

import sys
import os
import traceback

def main():
    """Main launcher function."""
    try:
        # Ensure we can import the GUI module
        try:
            from spp_enhanced_gui import main as gui_main
        except ImportError as e:
            print(f"Error importing GUI module: {e}")
            print("Please ensure spp_enhanced_gui.py is in the same directory.")
            input("Press Enter to exit...")
            return 1
        
        # Launch the GUI application
        print("Starting SPP Automation Tool Enhanced...")
        gui_main()
        return 0
        
    except KeyboardInterrupt:
        print("\nApplication interrupted by user.")
        return 0
    except Exception as e:
        print(f"Fatal error: {e}")
        print(f"Traceback: {traceback.format_exc()}")
        input("Press Enter to exit...")
        return 1

if __name__ == "__main__":
    sys.exit(main())
