# SPP Metric Automation Tool - Quick Start Guide

## 🚀 Welcome to SPP Automation!
This tool automates your monthly vendor performance reporting by:
- Pulling data from Snowflake (both vendor metrics and ASN data)
- Creating formatted Excel files with multiple tabs
- Auto-naming files based on vendor and month
- Running Excel macros automatically

---

## 📦 Package Contents
- `spp_automation.exe` - Main automation tool (GUI)
- `config.ini` - Configuration file (edit with your email)
- `Quick_Start.bat` - Easy launcher
- `template/` - Excel template folder
- `examples/` - Sample usage files

---

## ⚡ Quick Start (2 minutes)

### Step 1: Setup
1. **Extract** all files to a folder (e.g., `C:\SPP_Automation\`)
2. **Edit** `config.ini` - Replace `your_email@hdsupply.com` with your actual HD Supply email
3. **Verify** template path in `config.ini` points to your Excel template

### Step 2: Run
1. **Double-click** `Quick_Start.bat`
2. **Choose Option 1** (GUI Interface)
3. **Fill in the form**:
   - Vendors: `52889` (or multiple: `52889, 11833`)
   - Report Month: `FY2025-APR`
   - Date Filter: `202504`
4. **Click "Run Automation"**

### Step 3: Authenticate
- Browser will open for Snowflake login
- Login with your HD Supply credentials
- Return to the tool

### Step 4: Results
- Excel file created in `Output/` folder
- File named: `52889 - BOXER_HOME_LLC - Apr 2025.xlsm`
- Ready to email to suppliers!

---

## 🎯 Input Examples

| Field | Example | Description |
|-------|---------|-------------|
| Vendors | `52889` | Single vendor |
| Vendors | `52889, 11833, 200000` | Multiple vendors |
| Report Month | `FY2025-APR` | Fiscal year format |
| Date Filter | `202504` | YYYYMM for ASN data |

---

## 📋 Monthly Workflow

1. **Determine vendors** to report on
2. **Run automation** for each vendor/group
3. **Review** generated Excel files
4. **Email** files to suppliers
5. **Archive** files for records

---

## 🆘 Troubleshooting

| Problem | Solution |
|---------|----------|
| Connection fails | Check email in `config.ini`, verify Snowflake access |
| No data returned | Verify vendor numbers and date ranges |
| Excel template error | Check template path in `config.ini` |
| Macro doesn't run | Open Excel file manually and run macro |

---

## 📞 Support
- Check `spp_automation.log` for detailed error messages
- Contact SPP team for assistance
- Review `examples/` folder for usage patterns

---

## 🔧 Advanced Usage
- **Command Line**: Run `spp_automation.exe --help` for CLI options
- **Batch Processing**: Use `examples/batch_process.py` for multiple vendors
- **Custom Queries**: Modify queries in the source code

---

**Ready to automate your supplier reporting? Start with Quick_Start.bat!** 🎉
