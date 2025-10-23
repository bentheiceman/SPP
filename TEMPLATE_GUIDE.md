# SPP Template Setup Guide

## 📋 Overview

The enhanced SPP automation tool now supports user-configurable Excel templates, allowing each team member to customize their report output according to their specific needs while maintaining data consistency.

## 🎯 Template Requirements

### Minimum Requirements
Your Excel template must include these essential elements:

#### Required Columns (any order):
- **Product Number** or **Item Number**
- **Description** or **Product Description**
- **Quantity** or **Qty**
- **Unit Price** or **Price**
- **Extended Price** or **Total**
- **Vendor** or **Supplier**
- **Category** or **Product Category**

#### Optional Enhancements:
- Custom formatting and styling
- Additional calculated columns
- Charts and graphs
- Company branding
- Custom headers and footers

## 📁 Template Locations

The application automatically searches for templates in these locations (in order):

### 1. OneDrive Folder
- **Path**: `%UserProfile%\OneDrive\Documents`
- **Benefits**: Automatic sync across devices
- **Recommended**: Place templates here for best experience

### 2. Documents Folder
- **Path**: `%UserProfile%\Documents`
- **Benefits**: Local access, no sync required
- **Use Case**: Offline or local-only templates

### 3. Application Directory
- **Path**: Same folder as `SPP_Enhanced.exe`
- **Benefits**: Travels with the application
- **Use Case**: Portable installations

## 🛠 Creating Your Template

### Step 1: Start with Sample
1. Use the provided `SPP_Template.xlsx` as your starting point
2. Copy it to your preferred location (OneDrive recommended)
3. Rename it to reflect your needs (e.g., `My_SPP_Template.xlsx`)

### Step 2: Customize Layout
```
Column A: Product Number
Column B: Description
Column C: Quantity
Column D: Unit Price
Column E: Extended Price
Column F: Vendor
Column G: Category
[Add your custom columns here]
```

### Step 3: Add Your Branding
- Insert company logo
- Customize header/footer
- Apply your color scheme
- Add any required disclaimers

### Step 4: Include Calculations (Optional)
- Subtotals and totals
- Tax calculations
- Discount applications
- Custom formulas

### Step 5: Add Visual Elements (Optional)
- Charts and graphs
- Conditional formatting
- Data validation
- Pivot tables

## 📊 Template Formats

### Excel Workbook (.xlsx)
- **Use for**: Standard reports with formatting
- **Features**: All Excel formatting, formulas, charts
- **Limitations**: No macros or VBA code
- **Recommended**: Most common use case

### Excel Macro-Enabled Workbook (.xlsm)
- **Use for**: Advanced reports with automation
- **Features**: Full Excel functionality + macros
- **Requirements**: Macro security settings
- **Advanced**: Custom automation and processing

## ⚙ Configuration Process

### In the Application:
1. **Launch** `SPP_Enhanced.exe`
2. **Navigate** to "Template Configuration" section
3. **Browse** for your template file
4. **Select** the appropriate format (.xlsx or .xlsm)
5. **Save** your configuration

### Template Validation:
The application will automatically:
- Verify the template file exists
- Check file format compatibility
- Validate essential column presence
- Provide feedback on configuration status

## 🔧 Advanced Template Features

### Dynamic Data Ranges
Your template can include:
```excel
=OFFSET(A2,0,0,COUNTA(A:A)-1,8)  # Auto-expanding data range
=SUM(E:E)                        # Total calculations
=COUNTIF(G:G,"Electronics")      # Category counts
```

### Conditional Formatting
```excel
Rule: Cell Value > 1000
Format: Bold, Green background
Range: Extended Price column
```

### Custom Formulas
```excel
Profit Margin: =(E2-D2)/E2*100
Status: =IF(C2>0,"In Stock","Out of Stock")
Priority: =IF(E2>5000,"High","Normal")
```

## 🎨 Styling Guidelines

### Professional Appearance
- **Headers**: Bold, colored background
- **Borders**: Clean, consistent lines
- **Fonts**: Business standard (Arial, Calibri)
- **Colors**: Company brand colors
- **Alignment**: Consistent throughout

### Data Formatting
- **Numbers**: Proper decimal places
- **Currency**: Standard currency format
- **Dates**: Consistent date format
- **Text**: Left-aligned, proper case

## 🔄 Template Management

### Version Control
- Include version number in filename
- Date templates for easy tracking
- Keep backup copies
- Document changes made

### Sharing Templates
- Store in shared OneDrive folder
- Create team template library
- Document template purposes
- Maintain access permissions

## ⚠️ Troubleshooting

### Common Issues

#### "Template Not Found"
- **Solution**: Check file path and permissions
- **Verify**: Template is in expected location
- **Check**: File name and extension

#### "Format Not Supported"
- **Solution**: Ensure file is .xlsx or .xlsm
- **Avoid**: .xls, .csv, or other formats
- **Convert**: Use Excel "Save As" to proper format

#### "Missing Columns"
- **Solution**: Verify required columns exist
- **Check**: Column headers match expected names
- **Update**: Template with missing columns

### Performance Tips
- Keep templates under 10MB for faster loading
- Avoid complex macros that slow processing
- Use efficient formulas and references
- Regular template cleanup and optimization

## 📞 Support

### Template Issues:
1. Verify template meets minimum requirements
2. Check file format and location
3. Test template manually in Excel
4. Review application error messages

### Getting Help:
- Consult application logs
- Check template validation messages
- Test with sample template first
- Contact IT support if needed

---

**Template Setup Guide - SPP Enhanced v2.0**  
*Creating powerful, personalized SPP reports*

🎯 Remember: A well-designed template saves time and improves report consistency across your team!
