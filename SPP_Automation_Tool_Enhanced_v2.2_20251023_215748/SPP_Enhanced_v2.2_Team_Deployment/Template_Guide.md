# Excel Template Configuration Guide
## SPP Automation Tool Enhanced

### Template Structure
Your Excel template should contain the following sheets:
- **Tab1_Basic_Metrics**: For basic performance metrics
- **Tab2_ASN_Data**: For ASN (Advance Ship Notice) data

### Required Columns for Tab1_Basic_Metrics:
- VENDOR_NUMBER
- VENDOR_NAME
- VENDOR (Combined format)
- PO_NUMBER
- USN
- ITEM_DESCRIPTION
- DATE_ORIG_ORDERED
- DATE_ORIG_PROMISED
- DATE_FIRST_RECEIVED
- WAREHOUSE_NUM
- WAREHOUSE_NAME
- METRIC
- RPT_MONTH
- And other metric columns...

### Required Columns for Tab2_ASN_Data:
- VENDOR_NUMBER
- PO_NUMBER
- PO_LINE_ITEM
- MATERIAL_NUMBER
- DELIVERY_DATE
- QUANTITY
- UNIT_OF_MEASURE
- NET_PRICE
- NET_VALUE
- And other ASN columns...

### Template Best Practices:
1. Use consistent formatting (fonts, colors, borders)
2. Include company branding in headers
3. Pre-format number and date columns
4. Leave row 1 for headers, data starts at row 2
5. Use conditional formatting for key metrics
6. Include summary sheets if needed

### File Naming:
- Recommended: SPP_Template.xlsm (for macro support)
- Alternative: SPP_Template.xlsx (standard format)

### Macro Support:
- Save as .xlsm for macro-enabled features
- Include VBA code for automatic calculations
- Tool will preserve macros when using templates

### Storage Locations:
Place your template in any of these locations:
1. OneDrive Documents folder (automatically detected)
2. Local Documents folder (automatically detected)  
3. Same folder as the application
4. Custom path (configured in application)

The tool will find your template automatically or you can specify a custom path using the "Browse" button in the template configuration section.
