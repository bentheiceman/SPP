===============================================================================
  SPP AUTOMATION TOOL v3.0 - WITH PDH COMPLIANCE
  HD Supply Chain Excellence Team
===============================================================================

WHAT'S NEW IN v3.0
------------------
✨ PDH Compliance Tab (Tab4) - NEW!
   - Product Data Hub audit compliance tracking
   - Pass/Fail metrics for supplier performance
   - Rolling 28-day compliance window
   - Request and response tracking

✅ Enhanced vendor performance reporting across 4 comprehensive tabs
✅ Improved multi-vendor support
✅ Streamlined authentication process


HOW TO USE
----------
1. LAUNCH THE APPLICATION
   - Double-click: SPP_Automation_v3.0_PDH.exe
   - Application window will open with HD Supply branding

2. AUTHENTICATE (First Time Only)
   - Enter your @hdsupply.com email address
   - Click "Authenticate" button
   - Complete browser authentication within 2 minutes
   - You only need to authenticate once per session

3. GENERATE REPORTS
   Required Inputs:
   - Vendor Number(s): Enter one or more (comma-separated)
     Example: 13479  OR  13479, 52889, 200000
   
   - Report Month: Fiscal year format
     Example: FY2025-JAN, FY2026-APR
   
   - Date Filter: YYYYMM format
     Example: 202601 for January 2026
   
   - Click "🚀 Generate SPP Report"
   - Wait for completion (typically 30 seconds - 2 minutes)
   - Find your report in the Output folder


REPORT STRUCTURE
----------------
Your Excel report will contain 4 tabs:

📊 Tab1 - Summary Metrics
   - High-level KPIs and performance summary
   - Key percentage calculations
   - ASN success rates

📊 Tab2 - Basic Metrics  
   - Detailed line-level performance data
   - PO information and tracking
   - Receipt and delivery details

📊 Tab3 - ASN Data
   - Advance Shipping Notice compliance
   - Delivery tracking and carrier information
   - ASN submission patterns

📊 Tab4 - PDH Compliance ⭐ NEW!
   - Product Data Hub audit compliance
   - Request/response tracking with timestamps
   - Compliance labels (Compliant/Non-Compliant)
   - Days since request metrics
   - Supplier action status


OUTPUT FILES
------------
- Reports are saved in the "Output" folder
- File naming format: VENDOR# - VENDOR_NAME - MONTH.xlsx
- Example: 13479 - ACME_CORP - Jan 2026.xlsx


SYSTEM REQUIREMENTS
-------------------
✅ Windows 10 or later
✅ Internet connection (for Snowflake database)
✅ HD Supply VPN (if required by your organization)
✅ Web browser (for authentication)
✅ No additional software needed - this is a standalone executable!


TROUBLESHOOTING
---------------
Problem: "Failed to connect to Snowflake"
Solution: 
  - Check VPN connection
  - Complete browser authentication within 2 minutes
  - Verify email address format
  - Contact IT if issues persist

Problem: "No data found for vendor"
Solution:
  - Verify vendor number is correct
  - Check report month format (FY2025-JAN)
  - Ensure date filter matches report month (202501 for JAN)

Problem: Application won't start
Solution:
  - Check Windows Defender/antivirus settings
  - Right-click exe → Properties → Unblock (if present)
  - Run as Administrator if needed


ADVANCED FEATURES
-----------------
🔧 Custom Templates
   - Check "Use Excel Template for Output"
   - Browse to select your template file
   - Template must have sheets named:
     Tab1_Summary_Metrics
     Tab2_Basic_Metrics
     Tab3_ASN_Data
     Tab4_PDH_Compliance

🔧 Multi-Vendor Reports
   - Enter multiple vendor numbers separated by commas
   - Example: 13479, 52889, 200000
   - One report will be generated per vendor


SUPPORT
-------
Developer: Ben F. Benjamaa
Manager: Lauren B. Trapani
Team: HD Supply Chain Excellence

For questions or issues:
1. Check the Activity Log in the application
2. Review this documentation
3. Contact your team lead
4. IT support for technical issues


VERSION INFORMATION
-------------------
Version: 3.0
Release Date: January 7, 2026
Build: 20260107_065608
Features: 4-Tab Reporting with PDH Compliance


LICENSE & USAGE
---------------
This tool is for HD Supply internal use only.
Developed for the Chain Excellence team to automate supplier performance reporting.

===============================================================================
  Enjoy the new PDH Compliance feature! 🎉
===============================================================================
