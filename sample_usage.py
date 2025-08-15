"""
Sample Usage Script for SPP Metric Automation
This script demonstrates how to use the automation tool with different configurations.
"""

from spp_metric_automation import SPPMetricAutomation

def example_single_vendor():
    """Example: Single vendor automation."""
    print("=== Single Vendor Example ===")
    
    automation = SPPMetricAutomation()
    
    # Single vendor configuration
    vendor_numbers = ['52889']
    report_month = 'FY2025-APR'
    date_filter = '202504'  # April 2025
    
    try:
        output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
        print(f"SUCCESS: {output_file}")
        return output_file
    except Exception as e:
        print(f"ERROR: {e}")
        return None

def example_multiple_vendors():
    """Example: Multiple vendors automation."""
    print("=== Multiple Vendors Example ===")
    
    automation = SPPMetricAutomation()
    
    # Multiple vendors configuration
    vendor_numbers = ['52889', '11833']
    report_month = 'FY2025-MAY'
    date_filter = '202505'  # May 2025
    
    try:
        output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
        print(f"SUCCESS: {output_file}")
        return output_file
    except Exception as e:
        print(f"ERROR: {e}")
        return None

def example_custom_config():
    """Example: Using custom configuration file."""
    print("=== Custom Config Example ===")
    
    # Create custom config if needed
    custom_config = "custom_config.ini"
    
    automation = SPPMetricAutomation(custom_config)
    
    vendor_numbers = ['52889']
    report_month = 'FY2025-JUN'
    date_filter = '202506'  # June 2025
    
    try:
        output_file = automation.run_automation(vendor_numbers, report_month, date_filter)
        print(f"SUCCESS: {output_file}")
        return output_file
    except Exception as e:
        print(f"ERROR: {e}")
        return None

def test_queries_only():
    """Test query generation without execution."""
    print("=== Query Generation Test ===")
    
    automation = SPPMetricAutomation()
    
    # Test query generation
    vendor_numbers = ['52889', '11833']
    report_month = 'FY2025-APR'
    date_filter = '202504'
    
    print("Query 1 (Basic Metrics):")
    print("-" * 40)
    query1 = automation.get_query_1_basic_metrics(vendor_numbers, report_month)
    print(query1[:500] + "..." if len(query1) > 500 else query1)
    
    print("\nQuery 2 (ASN Data):")
    print("-" * 40)
    query2 = automation.get_query_2_asn_data(vendor_numbers, date_filter)
    print(query2[:500] + "..." if len(query2) > 500 else query2)

if __name__ == "__main__":
    print("SPP Metric Automation - Sample Usage")
    print("=" * 50)
    
    # Test query generation first (no database connection needed)
    test_queries_only()
    
    print("\n" + "=" * 50)
    print("NOTE: To run the full automation examples, ensure:")
    print("1. config.ini is properly configured with Snowflake credentials")
    print("2. Excel template file exists at specified path")
    print("3. Network access to Snowflake is available")
    print("=" * 50)
    
    # Uncomment the lines below to run full automation examples
    # (requires proper Snowflake configuration)
    
    # example_single_vendor()
    # example_multiple_vendors()
    # example_custom_config()
