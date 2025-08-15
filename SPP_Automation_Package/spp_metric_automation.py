"""
SPP Metric Automation Tool
Automates the process of running multiple queries and populating Excel templates
with vendor performance metrics and ASN data.
"""

import pandas as pd
import snowflake.connector
import openpyxl
from openpyxl import load_workbook
import os
import shutil
from datetime import datetime
import xlwings as xw
from typing import Any, Dict, List, Optional, Tuple
import configparser
import logging

class SPPMetricAutomation:
    def __init__(self, config_file: str = "config.ini"):
        """Initialize the SPP Metric Automation tool."""
        # Setup logging first so it is available during config loading
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('spp_automation.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # Load configuration
        self.config = self._load_config(config_file)
        
        # Initialize connection and template path
        self.connection: Optional[Any] = None
        self.template_path = r"C:\Users\1015723\OneDrive - HD Supply, Inc\Documents\Cole - Multi-Query information and template\Simplified Macro Template - SPP Monthly Details.xlsm"
        
    def _load_config(self, config_file: str) -> configparser.ConfigParser:
        """Load configuration from file."""
        config = configparser.ConfigParser()
        if os.path.exists(config_file):
            config.read(config_file)
        else:
            # Create default config
            config['SNOWFLAKE'] = {
                'account': 'HDSUPPLY-DATA',
                'user': 'your_email@hdsupply.com',
                'authenticator': 'externalbrowser',
                'insecure_mode': 'True',
                'warehouse': '',
                'database': '',
                'schema': ''
            }
            config['PATHS'] = {
                'template_path': r"C:\Users\1015723\OneDrive - HD Supply, Inc\Documents\Cole - Multi-Query information and template\Simplified Macro Template - SPP Monthly Details.xlsm",
                'output_directory': r"c:\Users\1015723\Downloads\SPP\Output"
            }
            with open(config_file, 'w') as f:
                config.write(f)
            self.logger.warning(f"Created default config file: {config_file}. Please update with your credentials.")
        return config
    
    def connect_to_snowflake(self) -> bool:
        """Establish connection to Snowflake."""
        try:
            # Check if using external browser authentication
            if self.config.get('SNOWFLAKE', 'authenticator', fallback='') == 'externalbrowser':
                self.connection = snowflake.connector.connect(
                    account=self.config['SNOWFLAKE']['account'],
                    user=self.config['SNOWFLAKE']['user'],
                    authenticator=self.config['SNOWFLAKE']['authenticator'],
                    insecure_mode=self.config.getboolean('SNOWFLAKE', 'insecure_mode', fallback=True)
                )
            else:
                # Traditional username/password authentication
                self.connection = snowflake.connector.connect(
                    account=self.config['SNOWFLAKE']['account'],
                    user=self.config['SNOWFLAKE']['user'],
                    password=self.config['SNOWFLAKE']['password'],
                    warehouse=self.config.get('SNOWFLAKE', 'warehouse', fallback=''),
                    database=self.config.get('SNOWFLAKE', 'database', fallback=''),
                    schema=self.config.get('SNOWFLAKE', 'schema', fallback='')
                )
            
            self.logger.info("Successfully connected to Snowflake")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect to Snowflake: {e}")
            return False
    
    def get_query_1_basic_metrics(self, vendor_numbers: List[str], report_month: str) -> str:
        """Generate Query 1 - Basic Metrics with dynamic filters."""
        vendor_filter = "', '".join(vendor_numbers)
        
        return f"""
WITH primary_metric AS (
    SELECT
        VENDOR_NUMBER,
        VENDOR_NAME,
        CONCAT(VENDOR_NUMBER,' - ',vendor_name) AS VENDOR,
        PO_NUMBER,
        USN,
        ITEM_DESCRIPTION,
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        CONCAT (USN, ' - ', ITEM_DESCRIPTION) as SKU,
        TO_DATE(DATE_ORIG_ORDERED) AS DATE_ORIG_ORDERED,
        TO_DATE(DATE_ORIG_PROMISED) AS DATE_ORIG_PROMISED,
        TO_DATE(DATE_FIRST_RECEIVED) AS DATE_FIRST_RECEIVED,
        WAREHOUSE_NUM,
        WAREHOUSE_NAME,
        METRIC,
        RPT_MONTH,
        FSCL_YR_PRD,
        METRIC_NUMERATOR,
        METRIC_DENOMINATOR,
        NETWORK
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE 
        VENDOR_NUMBER IN ('{vendor_filter}')
        AND RPT_MONTH = '{report_month}'
        AND METRIC IN ('First_Receipt_FR_B1D', 'First_Receipt_FR_B28D','Units_On_Time_Complete')
),

hds_receipts AS (
    SELECT 
        CONCAT(ebeln, ':', LTRIM(MATNR, '0')) AS Metric_Concatenate,
        MAX(TRY_TO_DATE(TO_CHAR(budat), 'yyyymmdd')) AS receipt_date
    FROM edp.std_ecc.ekbe
    WHERE bwart IN ('101', '102')
    GROUP BY ebeln, MATNR
),

hdp_receipts AS (
    SELECT
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        MAX(TO_DATE(DATE_RECEIVED)) AS receipt_date
    FROM DM_SUPPLYCHAIN.PRO_INVENTORY_ANALYTICS.REPORT_PURCHASE_ORDER_VISIBILITY_SHIPMENTS
    GROUP BY PO_NUMBER, USN
),

combined_receipts AS (
    SELECT
        Metric_Concatenate,
        MAX(receipt_date) AS receipt_date
    FROM (
        SELECT * FROM hds_receipts
        UNION ALL
        SELECT * FROM hdp_receipts
    ) all_receipts
    GROUP BY Metric_Concatenate
)

SELECT
    pm.RPT_MONTH as Report_Month,
    pm.NETWORK,
    pm.VENDOR,
    pm.WAREHOUSE_NUM,
    pm.WAREHOUSE_NAME,
    pm.PO_NUMBER,
    pm.SKU,
    pm.DATE_ORIG_ORDERED as Date_Ordered,
    pm.DATE_FIRST_RECEIVED,
    cr.receipt_date as Date_Last_Received,
    pm.METRIC,
    zeroifnull(pm.METRIC_NUMERATOR) as Metric_Units_Received,
    pm.METRIC_DENOMINATOR as Metric_Units_Ordered,
    case 
        when Metric_Units_Received < Metric_Units_Ordered
        then 'Non-Compliant'
        else 'Compliant'
    end as "Result"
    
FROM primary_metric pm
LEFT JOIN combined_receipts cr
    ON pm.Metric_Concatenate = cr.Metric_Concatenate
ORDER BY pm.VENDOR, pm.PO_NUMBER, pm.USN;
"""
    
    def get_query_2_asn_data(self, vendor_numbers: List[str], date_filter: str) -> str:
        """Generate Query 2 - ASN Data with dynamic filters."""
        vendor_filter = "', '".join(vendor_numbers)
        
        return f"""
SELECT
    CASE 
        WHEN IH.VBELN LIKE '06%' AND IH.ERNAM IN ('BPAREMOTE', 'SCEBATCH', 'P2P_IDOC', 'P2PBATCH') THEN 'ASN'
        WHEN IH.VBELN LIKE '10%' THEN 'EGR'
        ELSE 'Manually Created'
    END AS Inbound_Type,

    V.NAME1 AS Vendor_Name,
    LTRIM(IH.LIFNR, 0) AS Vendor_Number,
    TO_DATE(IH.ERDAT, 'YYYYMMDD') AS Create_Date,
    IH.VSTEL AS DC,
    ZUKRL AS PO_Number,
    LTRIM(IL.MATNR, 0) AS Material_Number,
    IL.ARKTX AS Material_Description,
    IL.LFIMG AS Quantity,
    IL.MEINS AS Unit_of_Measure,
    LTRIM(IH.VBELN, 0) AS Delivery_Number,
    LIFEX AS Supplier_Provided_ID,
    ZCARRIER_T AS Carrier_Name,
    BOLNR AS Provided_BOL

FROM EDP.STD_ECC.LIKP IH 

INNER JOIN EDP.STD_ECC.LIPS IL
    ON IH.MANDT = IL.MANDT
    AND IH.VBELN = IL.VBELN

INNER JOIN EDP.STD_ECC.LFA1 V
    ON IH.MANDT = V.MANDT
    AND IH.LIFNR = V.LIFNR

WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND LTRIM(IH.LIFNR, 0) IN ('{vendor_filter}')
    AND IH.ERDAT LIKE '{date_filter}%'
"""
    def execute_query(self, query: str) -> pd.DataFrame:
        """Execute a query and return results as DataFrame."""
        if self.connection is None:
            self.logger.error("No active Snowflake connection. Call connect_to_snowflake() before executing queries.")
            raise RuntimeError("No active Snowflake connection. Call connect_to_snowflake() before executing queries.")
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(query)
                
                # Get column names
                columns = [desc[0] for desc in cursor.description]
                
                # Fetch all results
                results = cursor.fetchall()
            
            # Create DataFrame
            df = pd.DataFrame(results, columns=columns)
            
            self.logger.info(f"Query executed successfully. Returned {len(df)} rows.")
            return df
            
        except Exception as e:
            self.logger.error(f"Error executing query: {e}")
            raise
    
    def get_vendor_name_from_data(self, df: pd.DataFrame) -> str:
        """Extract vendor name from query results."""
        if 'VENDOR' in df.columns and not df.empty:
            return df['VENDOR'].iloc[0].split(' - ', 1)[1] if ' - ' in df['VENDOR'].iloc[0] else df['VENDOR'].iloc[0]
        elif 'Vendor_Name' in df.columns and not df.empty:
            return df['Vendor_Name'].iloc[0]
        else:
            return "Unknown_Vendor"
    
    def generate_filename(self, vendor_numbers: List[str], vendor_name: str, report_month: str) -> str:
        """Generate filename based on vendor and month information."""
        # Convert report month to readable format
        if 'FY' in report_month:
            # Format: FY2025-APR -> Apr 2025
            parts = report_month.split('-')
            if len(parts) == 2:
                year = parts[0].replace('FY', '')  # FY2025 -> 2025
                month = parts[1].title()  # APR -> Apr
                month_year = f"{month} {year}"
            else:
                month_year = report_month
        else:
            month_year = report_month
        
        # Create vendor string
        vendor_string = "_".join(vendor_numbers)
        
        # Clean vendor name for filename
        clean_vendor_name = "".join(c for c in vendor_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        clean_vendor_name = clean_vendor_name.replace(' ', '_')
        
        filename = f"{vendor_string} - {clean_vendor_name} - {month_year}.xlsm"
        return filename
    
    def create_output_directory(self) -> str:
        """Create and return output directory path."""
        output_dir = self.config.get('PATHS', 'output_directory', fallback=r"c:\Users\1015723\Downloads\SPP\Output")
        os.makedirs(output_dir, exist_ok=True)
        return output_dir
    
    def copy_template(self, output_path: str) -> str:
        """Copy template file to output location."""
        try:
            shutil.copy2(self.template_path, output_path)
            self.logger.info(f"Template copied to: {output_path}")
            return output_path
        except Exception as e:
            self.logger.error(f"Error copying template: {e}")
            raise
    
    def populate_excel_data(self, file_path: str, data_dict: Dict[str, pd.DataFrame]) -> None:
        """Populate Excel file with data from multiple queries."""
        try:
            # Use xlwings for better Excel integration
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            
            # Populate each sheet with corresponding data
            for sheet_name, df in data_dict.items():
                if sheet_name in [sheet.name for sheet in wb.sheets]:
                    ws = wb.sheets[sheet_name]
                    
                    # Clear existing data (starting from A1)
                    if not df.empty:
                        # Clear the range first
                        max_row = max(1000, len(df) + 10)  # Clear enough rows
                        max_col = max(50, len(df.columns) + 10)  # Clear enough columns
                        ws.range(f'A1:{xw.utils.col_name(max_col)}{max_row}').clear()
                        
                        # Write data starting at A1
                        ws.range('A1').value = df.values
                        # Write headers
                        ws.range('A1').value = [df.columns.tolist()] + df.values.tolist()
                        
                        self.logger.info(f"Populated {sheet_name} with {len(df)} rows")
                    else:
                        self.logger.warning(f"No data to populate in {sheet_name}")
                else:
                    self.logger.warning(f"Sheet {sheet_name} not found in template")
            
            # Run the macro if it exists
            try:
                wb.macro('RefreshAndCopy')()
                self.logger.info("Macro executed successfully")
            except Exception as macro_error:
                self.logger.warning(f"Could not run macro: {macro_error}")
            
            # Save and close
            wb.save()
            wb.close()
            app.quit()
            
            self.logger.info(f"Excel file populated and saved: {file_path}")
            
        except Exception as e:
            self.logger.error(f"Error populating Excel file: {e}")
            raise
    
    def run_automation(self, vendor_numbers: List[str], report_month: str, date_filter: str) -> str:
        """Run the complete automation process."""
        try:
            self.logger.info(f"Starting automation for vendors: {vendor_numbers}, month: {report_month}")
            
            # Connect to Snowflake
            if not self.connect_to_snowflake():
                raise Exception("Failed to connect to Snowflake")
            
            # Execute Query 1 - Basic Metrics
            self.logger.info("Executing Query 1 - Basic Metrics")
            query1 = self.get_query_1_basic_metrics(vendor_numbers, report_month)
            df_metrics = self.execute_query(query1)
            
            # Execute Query 2 - ASN Data
            self.logger.info("Executing Query 2 - ASN Data")
            query2 = self.get_query_2_asn_data(vendor_numbers, date_filter)
            df_asn = self.execute_query(query2)
            
            # Get vendor name for filename
            vendor_name = self.get_vendor_name_from_data(df_metrics) or self.get_vendor_name_from_data(df_asn)
            
            # Create output directory and filename
            output_dir = self.create_output_directory()
            filename = self.generate_filename(vendor_numbers, vendor_name, report_month)
            output_path = os.path.join(output_dir, filename)
            
            # Copy template to output location
            self.copy_template(output_path)
            
            # Prepare data dictionary for Excel population
            data_dict = {
                'METRIC DATA': df_metrics,
                'ASN Data': df_asn
            }
            
            # Populate Excel file
            self.populate_excel_data(output_path, data_dict)
            
            # Close Snowflake connection
            if self.connection:
                self.connection.close()
            
            self.logger.info(f"Automation completed successfully. File saved: {output_path}")
            return output_path
            
        except Exception as e:
            self.logger.error(f"Automation failed: {e}")
            if self.connection:
                self.connection.close()
            raise

# Command-line interface
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='SPP Metric Automation Tool')
    parser.add_argument('--vendors', nargs='+', required=True, help='Vendor numbers (space-separated)')
    parser.add_argument('--month', required=True, help='Report month (e.g., FY2025-APR)')
    parser.add_argument('--date-filter', required=True, help='Date filter for ASN query (e.g., 202507)')
    parser.add_argument('--config', default='config.ini', help='Configuration file path')
    
    args = parser.parse_args()
    
    # Create automation instance
    automation = SPPMetricAutomation(args.config)
    
    try:
        output_file = automation.run_automation(args.vendors, args.month, args.date_filter)
        print(f"Automation completed successfully!")
        print(f"Output file: {output_file}")
    except Exception as e:
        print(f"Automation failed: {e}")
        exit(1)
