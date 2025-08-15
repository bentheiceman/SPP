"""
SPP Automation - Quick Run Script
A simple script to run automation with hardcoded or command-line parameters.
"""

from spp_metric_automation import SPPMetricAutomation
import sys

def quick_run():
    """Quick run with example parameters."""
    
    # Example parameters - modify as needed
    vendor_numbers = ['52889']  # Add your vendor numbers here
    report_month = 'FY2025-APR'  # Format: FY2025-APR
    date_filter = '202507'  # Format: YYYYMM
    
    print("SPP Metric Automation - Quick Run")
    print("=" * 40)
    print(f"Vendors: {', '.join(vendor_numbers)}")
    print(f"Report Month: {report_month}")
    print(f"Date Filter: {date_filter}")
    print()
    
    try:
        # Create automation instance
        automation = SPPMetricAutomation()
        
        # Run automation
        output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
        
        print(f"SUCCESS! Output file created: {output_file}")
        
        # Ask if user wants to open the file
        response = input("\nWould you like to open the output file? (y/n): ")
        if response.lower() in ['y', 'yes']:
            import os
            os.startfile(output_file)
            
    except Exception as e:
        print(f"ERROR: {e}")
        return False
    
    return True

if __name__ == "__main__":
    # Check if command line arguments provided
    if len(sys.argv) > 1:
        if len(sys.argv) != 4:
            print("Usage: python quick_run.py <vendor_numbers> <report_month> <date_filter>")
            print("Example: python quick_run.py '52889,11833' 'FY2025-APR' '202507'")
            sys.exit(1)
        
        vendor_numbers = sys.argv[1].split(',')
        report_month = sys.argv[2]
        date_filter = sys.argv[3]
        
        automation = SPPMetricAutomation()
        try:
            output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
            print(f"SUCCESS! Output file: {output_file}")
        except Exception as e:
            print(f"ERROR: {e}")
            sys.exit(1)
    else:
        # Interactive mode
        quick_run()
