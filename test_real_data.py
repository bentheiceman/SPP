"""
SPP Automation - Real Data Test
Test the automation with actual vendor data using your Snowflake configuration.
"""

from spp_metric_automation import SPPMetricAutomation
import sys

def test_with_real_data():
    """Test automation with real vendor data."""
    print("SPP Metric Automation - Real Data Test")
    print("=" * 50)
    
    # Use the example data from your request
    vendor_numbers = ['52889']  # From your Query 1 example
    report_month = 'FY2025-APR'  # From your Query 1 example
    date_filter = '202504'  # April 2025 for ASN data
    
    print(f"Testing with:")
    print(f"  Vendors: {', '.join(vendor_numbers)}")
    print(f"  Report Month: {report_month}")
    print(f"  Date Filter: {date_filter}")
    print()
    
    try:
        # Create automation instance
        print("Initializing automation...")
        automation = SPPMetricAutomation()
        
        # Test connection first
        print("Testing Snowflake connection...")
        if not automation.connect_to_snowflake():
            print("❌ Connection failed. Please check your configuration.")
            return False
        
        print("✅ Connection successful!")
        
        # Test Query 1 execution
        print("\nTesting Query 1 (Basic Metrics)...")
        query1 = automation.get_query_1_basic_metrics(vendor_numbers, report_month)
        df_metrics = automation.execute_query(query1)
        print(f"✅ Query 1 executed successfully. Returned {len(df_metrics)} rows.")
        
        # Show sample of metrics data
        if not df_metrics.empty:
            print("Sample metrics data:")
            print(df_metrics[['VENDOR', 'PO_NUMBER', 'SKU', 'METRIC']].head())
        
        # Test Query 2 execution  
        print(f"\nTesting Query 2 (ASN Data)...")
        query2 = automation.get_query_2_asn_data(vendor_numbers, date_filter)
        df_asn = automation.execute_query(query2)
        print(f"✅ Query 2 executed successfully. Returned {len(df_asn)} rows.")
        
        # Show sample of ASN data
        if not df_asn.empty:
            print("Sample ASN data:")
            print(df_asn[['Vendor_Name', 'PO_Number', 'Material_Number', 'Inbound_Type']].head())
        
        # Close connection
        if automation.connection:
            automation.connection.close()
        
        print(f"\n🎉 All tests passed!")
        print(f"Metrics data: {len(df_metrics)} rows")
        print(f"ASN data: {len(df_asn)} rows")
        
        # Ask if user wants to run full automation
        if len(df_metrics) > 0 or len(df_asn) > 0:
            print("\n" + "="*50)
            response = input("Data looks good! Run full automation to create Excel file? (y/n): ")
            if response.lower() in ['y', 'yes']:
                print("\nRunning full automation...")
                return run_full_automation(vendor_numbers, report_month, date_filter)
        
        return True
        
    except Exception as e:
        print(f"❌ Error during testing: {e}")
        return False

def run_full_automation(vendor_numbers, report_month, date_filter):
    """Run the complete automation process."""
    try:
        automation = SPPMetricAutomation()
        output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
        
        print(f"\n🎉 Automation completed successfully!")
        print(f"Output file: {output_file}")
        
        # Ask if user wants to open the file
        response = input("\nWould you like to open the output file? (y/n): ")
        if response.lower() in ['y', 'yes']:
            import os
            os.startfile(output_file)
        
        return True
        
    except Exception as e:
        print(f"❌ Full automation failed: {e}")
        return False

if __name__ == "__main__":
    success = test_with_real_data()
    
    if not success:
        print("\n⚠️  Test failed. Please check:")
        print("1. Snowflake connection and permissions")
        print("2. Vendor numbers and date ranges have data")
        print("3. Excel template file exists")
        print("\nRun 'python test_connection.py' to verify Snowflake connection.")
    
    input("\nPress Enter to continue...")
