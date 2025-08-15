# SPP Metric Automation Tool

## Overview
This tool automates the process of running multiple Snowflake queries and populating Excel templates with vendor performance metrics and ASN data. It eliminates manual data entry and formatting, saving significant time in monthly supplier reporting.

## Features
- **Multiple Query Support**: Handles Query 1 (Basic Metrics) and Query 2 (ASN Data) with extensibility for future queries
- **Dynamic Filtering**: Configurable vendor numbers, report months, and date ranges
- **Automatic File Naming**: Generates descriptive filenames based on vendor and month information
- **Excel Integration**: Populates Excel templates and runs macros automatically
- **Flexible Configuration**: Easy-to-edit configuration file for database connections
- **GUI Interface**: User-friendly graphical interface for non-technical users
- **Command Line Support**: Scriptable for automation workflows

## Setup Instructions

### 1. Install Required Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Database Connection
The tool uses external browser authentication for Snowflake. The configuration is already set up for HD Supply:

```ini
[SNOWFLAKE]
account = HDSUPPLY-DATA
user = your_email@hdsupply.com
authenticator = externalbrowser
insecure_mode = True
```

Simply update the `user` field with your HD Supply email address in `config.ini`.

### 3. Verify Template Path
Ensure the Excel template path in `config.ini` is correct:
```ini
[PATHS]
template_path = C:\Users\1015723\OneDrive - HD Supply, Inc\Documents\Cole - Multi-Query information and template\Simplified Macro Template - SPP Monthly Details.xlsm
```

## Usage Options

### Option 1: GUI Interface (Recommended)
```bash
python spp_gui.py
```
- Fill in vendor numbers (comma-separated): `52889, 11833`
- Enter report month: `FY2025-APR`
- Enter date filter: `202507`
- Click "Run Automation"

### Option 2: Quick Run Script
```bash
python quick_run.py
```
Edit the script to change default parameters or use command line:
```bash
python quick_run.py "52889,11833" "FY2025-APR" "202507"
```

### Option 3: Command Line
```bash
python spp_metric_automation.py --vendors 52889 11833 --month FY2025-APR --date-filter 202507
```

## Query Details

### Query 1: Basic Metrics
- Pulls vendor performance metrics from `DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE`
- Includes fill rates, compliance status, and receipt dates
- Populates the "METRIC DATA" tab in Excel

### Query 2: ASN Data
- Pulls ASN information from SAP ECC tables (`LIKP`, `LIPS`, `LFA1`)
- Includes delivery numbers, carrier information, and BOL data
- Populates the "ASN Data" tab in Excel

## Output
- Files are saved to `c:\Users\1015723\Downloads\SPP\Output\`
- Filename format: `{Vendor_Number(s)} - {Vendor_Name} - {Month} {Year}.xlsm`
- Example: `52889 - ACME_Corporation - Apr 2025.xlsm`

## Adding New Queries
To add additional queries (Query 3, Query 4, etc.):

1. Add a new method in `SPPMetricAutomation` class:
```python
def get_query_3_new_data(self, vendor_numbers: List[str], additional_params: str) -> str:
    # Your new query here
    return query_string
```

2. Execute the query in `run_automation()`:
```python
df_new_data = self.execute_query(self.get_query_3_new_data(vendor_numbers, params))
```

3. Add to data dictionary:
```python
data_dict = {
    'METRIC DATA': df_metrics,
    'ASN Data': df_asn,
    'New Data Tab': df_new_data  # New tab
}
```

## Troubleshooting

### Common Issues
1. **Snowflake Connection Failed**: Check credentials in `config.ini`
2. **Template Not Found**: Verify template path exists
3. **Excel Macro Errors**: Ensure Excel allows macros and template has required macro
4. **Empty Results**: Verify vendor numbers and date filters are correct

### Logs
- Check `spp_automation.log` for detailed execution logs
- GUI displays real-time status messages

## Macro Integration
The tool attempts to run the existing Excel macro (`ctrl+shift+m` equivalent) automatically. The macro should:
1. Refresh pivots
2. Copy/paste pivot tables to appropriate locations
3. Format data as needed

## Security Notes
- Store sensitive credentials securely
- Consider using environment variables for production deployments
- Regularly rotate database passwords

## Support
For issues or enhancements, contact the SPP team or modify the code as needed. The tool is designed to be extensible and maintainable.
