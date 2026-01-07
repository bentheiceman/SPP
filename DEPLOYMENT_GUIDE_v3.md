# SPP Automation Tool v3.0 - Deployment Guide with PDH Compliance

## 🎉 What's New in v3.0

Your SPP tool now includes **PDH Compliance tracking** in Tab4!
- Product Data Hub Audit compliance data
- Pass/Fail metrics for supplier performance
- Rolling 28-day compliance window
- Request and response tracking

---

## 📦 Quick Deployment for Your Team

### Option 1: Build Executable (Recommended)

**Step 1: Open PowerShell (not Python REPL)**
```powershell
cd C:\Users\1015723\Downloads\SPP
```

**Step 2: Exit Python if you're in it**
```
exit()
```

**Step 3: Run the build script**
```powershell
python build_spp_v3.py
```

**OR use the batch file:**
```powershell
.\Build_SPP_v3.bat
```

This will create a `SPP_v3.0_PDH_Deployment_*` folder with:
- `SPP_Automation_v3.0_PDH.exe` (standalone executable)
- `README.txt` (full documentation)
- `QUICK_START.txt` (quick reference)
- `VERSION_INFO.txt` (version details)
- `Output/` (folder for generated reports)

---

### Option 2: Direct Python Distribution (If Build Fails)

If the executable build doesn't work, you can share the Python scripts directly:

**Files to share with your team:**
1. `spp_enhanced_gui.py` (GUI application)
2. `spp_automation_enhanced.py` (backend engine)
3. `config.ini` (optional configuration)
4. `template_config.json` (optional template settings)
5. `requirements.txt` (dependencies)

**Team members need to:**
```powershell
# Install Python 3.8+ from python.org
# Then install dependencies:
pip install pandas openpyxl snowflake-connector-python

# Run the application:
python spp_enhanced_gui.py
```

---

## 🚀 User Guide for Your Team

### First Time Setup
1. **Launch** the application
   - Double-click `SPP_Automation_v3.0_PDH.exe`
   - OR run `python spp_enhanced_gui.py`

2. **Authenticate**
   - Enter your @hdsupply.com email
   - Click "Authenticate"
   - Complete browser authentication
   - You only need to do this once per session

3. **Ready to Go!**

### Generating Reports

**Required Inputs:**
- **Vendor Number(s)**: One or more vendor numbers (comma-separated)
  - Example: `13479` or `13479, 52889, 12345`
- **Report Month**: Fiscal year format
  - Example: `FY2025-JAN`, `FY2026-APR`
- **Date Filter**: YYYYMM format
  - Example: `202501` for January 2025

**Steps:**
1. Fill in the vendor number(s)
2. Select report month
3. Enter date filter
4. Click **"🚀 Generate SPP Report"**
5. Wait for completion (30 seconds - 2 minutes)
6. Find your report in the `Output` folder

### Report Contents (All 4 Tabs)

**Tab1 - Summary Metrics**
- High-level KPIs
- Percentage calculations
- ASN success rate

**Tab2 - Basic Metrics**
- Line-level performance data
- Detailed PO information
- Receipt tracking

**Tab3 - ASN Data**
- Advance Shipping Notice compliance
- Delivery tracking
- Carrier information

**Tab4 - PDH Compliance** ⭐ NEW!
- Product Data Hub audit compliance
- Request/response tracking
- Compliance labels (Compliant/Non-Compliant)
- Days since request
- Supplier action status

---

## 🔧 Advanced Features

### Using Custom Templates
1. Check **"Use Excel Template for Output"**
2. Click **"Browse"** to select your template file
3. Template must have sheets named:
   - `Tab1_Summary_Metrics`
   - `Tab2_Basic_Metrics`
   - `Tab3_ASN_Data`
   - `Tab4_PDH_Compliance`

### Output Formats
- **Standard Excel (.xlsx)**: Default, compatible with all Excel versions
- **Macro-Enabled (.xlsm)**: For templates with macros/VBA

---

## 🐛 Troubleshooting

### "Failed to connect to Snowflake"
- ✅ Check VPN connection
- ✅ Complete browser authentication within 2 minutes
- ✅ Verify email address is correct
- ✅ Contact IT if issues persist

### "No data found"
- ✅ Verify vendor number exists in system
- ✅ Check report month format (FY2025-JAN)
- ✅ Ensure date filter matches report month

### "Module not found" or Import Errors
```powershell
pip install --upgrade pandas openpyxl snowflake-connector-python
```

### Build Fails
- Make sure you're in PowerShell (not Python REPL)
- Type `exit()` if you see `>>>`
- Ensure PyInstaller is installed: `pip install pyinstaller`

---

## 📋 System Requirements

- **OS**: Windows 10 or later
- **Python**: 3.8 - 3.11 (if running from source)
- **Internet**: Required for Snowflake connection
- **VPN**: HD Supply VPN (if required by your organization)
- **Browser**: For Snowflake authentication

---

## 📞 Support

**Developer**: Ben F. Benjamaa  
**Manager**: Lauren B. Trapani  
**Team**: HD Supply Chain Excellence

For questions or issues:
1. Check the Activity Log in the application
2. Review this documentation
3. Contact your team lead
4. IT support for technical issues

---

## 📝 Version History

**v3.0** (Current)
- ✨ NEW: PDH Compliance Tab (Tab4)
- ✨ NEW: Rolling 28-day compliance tracking
- ✅ Enhanced vendor performance reporting
- ✅ Improved error handling

**v2.2**
- Template configuration support
- Multi-vendor support
- Enhanced ASN tracking

**v2.0**
- Initial multi-tab implementation
- Snowflake integration
- GUI interface

---

## 🎯 Quick Reference

| Feature | Example |
|---------|---------|
| Vendor Number | `13479` or `13479, 52889` |
| Report Month | `FY2025-JAN` |
| Date Filter | `202501` |
| Output Location | `Output/` folder |
| File Name Format | `VENDOR# - NAME - MONTH.xlsx` |

**Keyboard Shortcuts:**
- `Tab`: Navigate between fields
- `Enter`: Submit in input fields
- `Ctrl+Q`: Quick quit (when in window)

---

**🎉 Enjoy the new PDH Compliance feature!**

This enhancement provides your suppliers with valuable visibility into their Product Data Hub compliance performance.
