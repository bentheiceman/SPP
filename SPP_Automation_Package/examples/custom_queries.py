"""
Custom Query Example for SPP Automation
Shows how to add new queries (Query 3, Query 4, etc.) to the automation.
"""

from spp_metric_automation import SPPMetricAutomation
import pandas as pd
import os

class ExtendedSPPAutomation(SPPMetricAutomation):
    """Extended version with additional queries."""
    
    def get_query_3_vendor_summary(self, vendor_numbers: list, report_month: str) -> str:
        """Example Query 3 - Vendor Summary Statistics."""
        vendor_filter = "', '".join(vendor_numbers)
        
        return f"""
        -- Query 3: Vendor Summary Statistics
        SELECT 
            VENDOR_NUMBER,
            VENDOR_NAME,
            COUNT(DISTINCT PO_NUMBER) as Total_POs,
            COUNT(DISTINCT USN) as Total_Materials,
            SUM(METRIC_NUMERATOR) as Total_Units_Received,
            SUM(METRIC_DENOMINATOR) as Total_Units_Ordered,
            ROUND(AVG(CASE WHEN METRIC_DENOMINATOR > 0 
                          THEN METRIC_NUMERATOR * 100.0 / METRIC_DENOMINATOR 
                          ELSE 0 END), 2) as Avg_Fill_Rate_Pct,
            MIN(DATE_ORIG_ORDERED) as Earliest_Order_Date,
            MAX(DATE_FIRST_RECEIVED) as Latest_Receipt_Date
        FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
        WHERE VENDOR_NUMBER IN ('{vendor_filter}')
            AND RPT_MONTH = '{report_month}'
            AND METRIC IN ('First_Receipt_FR_B1D', 'First_Receipt_FR_B28D')
        GROUP BY VENDOR_NUMBER, VENDOR_NAME
        ORDER BY VENDOR_NUMBER;
        """
    
    def get_query_4_material_analysis(self, vendor_numbers: list, report_month: str) -> str:
        """Example Query 4 - Material Performance Analysis."""
        vendor_filter = "', '".join(vendor_numbers)
        
        return f"""
        -- Query 4: Top Performing Materials
        SELECT 
            VENDOR_NUMBER,
            VENDOR_NAME,
            USN,
            ITEM_DESCRIPTION,
            COUNT(*) as Order_Count,
            SUM(METRIC_NUMERATOR) as Total_Received,
            SUM(METRIC_DENOMINATOR) as Total_Ordered,
            ROUND(SUM(METRIC_NUMERATOR) * 100.0 / SUM(METRIC_DENOMINATOR), 2) as Fill_Rate_Pct,
            CASE 
                WHEN SUM(METRIC_NUMERATOR) = SUM(METRIC_DENOMINATOR) THEN 'Perfect'
                WHEN SUM(METRIC_NUMERATOR) * 100.0 / SUM(METRIC_DENOMINATOR) >= 95 THEN 'Excellent'
                WHEN SUM(METRIC_NUMERATOR) * 100.0 / SUM(METRIC_DENOMINATOR) >= 90 THEN 'Good'
                ELSE 'Needs Improvement'
            END as Performance_Rating
        FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
        WHERE VENDOR_NUMBER IN ('{vendor_filter}')
            AND RPT_MONTH = '{report_month}'
            AND METRIC = 'First_Receipt_FR_B28D'
        GROUP BY VENDOR_NUMBER, VENDOR_NAME, USN, ITEM_DESCRIPTION
        HAVING SUM(METRIC_DENOMINATOR) > 0
        ORDER BY Fill_Rate_Pct DESC, Total_Ordered DESC
        LIMIT 50;
        """
    
    def run_extended_automation(self, vendor_numbers: list, report_month: str, date_filter: str) -> str:
        """Run automation with additional queries."""
        
        print("Extended SPP Automation with Additional Queries")
        print("=" * 50)
        
        # Connect to Snowflake
        if not self.connect_to_snowflake():
            raise Exception("Failed to connect to Snowflake")
        
        # Execute all queries
        print("Executing Query 1 - Basic Metrics...")
        query1 = self.get_query_1_basic_metrics(vendor_numbers, report_month)
        df_metrics = self.execute_query(query1)
        
        print("Executing Query 2 - ASN Data...")
        query2 = self.get_query_2_asn_data(vendor_numbers, date_filter)
        df_asn = self.execute_query(query2)
        
        print("Executing Query 3 - Vendor Summary...")
        query3 = self.get_query_3_vendor_summary(vendor_numbers, report_month)
        df_summary = self.execute_query(query3)
        
        print("Executing Query 4 - Material Analysis...")
        query4 = self.get_query_4_material_analysis(vendor_numbers, report_month)
        df_materials = self.execute_query(query4)
        
        # Get vendor name for filename
        vendor_name = self.get_vendor_name_from_data(df_metrics) or self.get_vendor_name_from_data(df_asn)
        
        # Create output directory and filename
        output_dir = self.create_output_directory()
        filename = self.generate_filename(vendor_numbers, vendor_name, report_month)
        filename = filename.replace('.xlsm', '_Extended.xlsm')  # Mark as extended version
        output_path = os.path.join(output_dir, filename)
        
        # Copy template
        self.copy_template(output_path)
        
        # Prepare data dictionary with additional tabs
        data_dict = {
            'METRIC DATA': df_metrics,
            'ASN Data': df_asn,
            'Vendor Summary': df_summary,
            'Material Analysis': df_materials
        }
        
        # Populate Excel file
        self.populate_excel_data(output_path, data_dict)
        
        # Close connection
        if self.connection:
            self.connection.close()
        
        print(f"Extended automation completed! File: {output_path}")
        return output_path

def demo_extended_automation():
    """Demonstrate the extended automation with additional queries."""
    
    # Configuration
    vendor_numbers = ['52889']
    report_month = 'FY2025-APR'
    date_filter = '202504'
    
    try:
        # Create extended automation instance
        automation = ExtendedSPPAutomation()
        
        # Run extended automation
        output_file = automation.run_extended_automation(vendor_numbers, report_month, date_filter)
        
        print(f"\n✅ Extended automation completed!")
        print(f"File created with 4 tabs: {output_file}")
        
        return output_file
        
    except Exception as e:
        print(f"❌ Extended automation failed: {e}")
        return None

if __name__ == "__main__":
    print("Extended SPP Automation Demo")
    print("This example shows how to add Query 3 and Query 4 to the automation.")
    print()
    
    response = input("Run extended automation demo? (y/n): ")
    if response.lower() in ['y', 'yes']:
        demo_extended_automation()
    
    input("\nPress Enter to continue...")
