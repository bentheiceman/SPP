"""
Batch Processing Example for SPP Automation
Process multiple vendors or time periods in sequence.
"""

from spp_metric_automation import SPPMetricAutomation
import time

def batch_process_vendors():
    """Process multiple vendors for the same month."""
    
    # Configuration
    vendors_list = [
        ['52889'],           # Vendor 1
        ['11833'],           # Vendor 2  
        ['200000', '210000'] # Multiple vendors together
    ]
    
    report_month = 'FY2025-APR'
    date_filter = '202504'
    
    print("SPP Batch Processing")
    print("=" * 40)
    print(f"Report Month: {report_month}")
    print(f"Date Filter: {date_filter}")
    print(f"Processing {len(vendors_list)} vendor groups...")
    print()
    
    automation = SPPMetricAutomation()
    results = []
    
    for i, vendors in enumerate(vendors_list, 1):
        try:
            print(f"Processing Group {i}: {', '.join(vendors)}")
            
            output_file = automation.run_automation(vendors, report_month, date_filter)
            results.append({
                'vendors': vendors,
                'status': 'SUCCESS',
                'file': output_file
            })
            
            print(f"✅ Success: {output_file}")
            
            # Wait between runs to avoid overwhelming Snowflake
            time.sleep(2)
            
        except Exception as e:
            print(f"❌ Error: {e}")
            results.append({
                'vendors': vendors,
                'status': 'ERROR',
                'error': str(e)
            })
    
    # Summary
    print("\n" + "=" * 40)
    print("BATCH PROCESSING SUMMARY")
    print("=" * 40)
    
    for result in results:
        vendor_str = ', '.join(result['vendors'])
        if result['status'] == 'SUCCESS':
            print(f"✅ {vendor_str}: {result['file']}")
        else:
            print(f"❌ {vendor_str}: {result['error']}")

def batch_process_months():
    """Process the same vendor for multiple months."""
    
    # Configuration
    vendor_numbers = ['52889']
    months_list = [
        ('FY2025-MAR', '202503'),
        ('FY2025-APR', '202504'),
        ('FY2025-MAY', '202505')
    ]
    
    print("SPP Monthly Batch Processing")
    print("=" * 40)
    print(f"Vendor: {', '.join(vendor_numbers)}")
    print(f"Processing {len(months_list)} months...")
    print()
    
    automation = SPPMetricAutomation()
    results = []
    
    for report_month, date_filter in months_list:
        try:
            print(f"Processing {report_month}...")
            
            output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
            results.append({
                'month': report_month,
                'status': 'SUCCESS',
                'file': output_file
            })
            
            print(f"✅ Success: {output_file}")
            time.sleep(2)
            
        except Exception as e:
            print(f"❌ Error: {e}")
            results.append({
                'month': report_month,
                'status': 'ERROR',
                'error': str(e)
            })
    
    # Summary
    print("\n" + "=" * 40)
    print("MONTHLY BATCH SUMMARY")
    print("=" * 40)
    
    for result in results:
        if result['status'] == 'SUCCESS':
            print(f"✅ {result['month']}: {result['file']}")
        else:
            print(f"❌ {result['month']}: {result['error']}")

if __name__ == "__main__":
    print("Choose batch processing type:")
    print("1. Multiple vendors, same month")
    print("2. Same vendor, multiple months")
    
    choice = input("Enter choice (1 or 2): ")
    
    if choice == "1":
        batch_process_vendors()
    elif choice == "2":
        batch_process_months()
    else:
        print("Invalid choice")
    
    input("\nPress Enter to continue...")
