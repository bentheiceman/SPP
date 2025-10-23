# SPP Enhanced Automation Tool v2.0 - Deployment Guide

## 🚀 Quick Start

1. **Extract** the deployment package to your desired location
2. **Run** `SPP_Enhanced.exe` to launch the application
3. **Configure** your Excel templates through the new Template Configuration section
4. **Enjoy** enhanced SPP automation with user-configurable templates!

## 📋 What's New in v2.0

### ✨ Enhanced Features
- **User-Configurable Templates**: Each user can now link custom Excel templates
- **Template Auto-Discovery**: Automatically finds templates in OneDrive, Documents, and app directory
- **Modern Interface**: Refreshed GUI with HD Supply branding and improved usability
- **Robust Error Handling**: Graceful fallbacks when templates are unavailable
- **Background Processing**: Non-blocking operations with progress feedback
- **Enhanced Logging**: Comprehensive error tracking and debugging information

### 🛠 Template Configuration System
- **Multi-Format Support**: Works with both .xlsx and .xlsm template files
- **Smart Discovery**: Searches multiple locations for your templates automatically
- **Graceful Fallbacks**: Continues with standard Excel format if templates unavailable
- **Real-Time Validation**: Instant feedback on template selection and format

## 📁 Files Included

```
SPP_Enhanced_Deployment/
├── SPP_Enhanced.exe          # Main application executable
├── DEPLOYMENT_GUIDE.md       # This guide
├── TEMPLATE_GUIDE.md         # Template setup instructions
├── TROUBLESHOOTING.md        # Common issues and solutions
└── sample_templates/         # Example template files
    ├── SPP_Template.xlsx
    └── SPP_Advanced_Template.xlsm
```

## 🔧 System Requirements

- **Operating System**: Windows 10/11
- **Memory**: 4GB RAM minimum (8GB recommended)
- **Storage**: 500MB free space
- **Network**: Internet connection for Snowflake database access
- **Permissions**: Standard user permissions (no admin required)

## 📊 Template Setup

1. **Place Templates**: Save your Excel templates in one of these locations:
   - OneDrive folder (automatically discovered)
   - Documents folder
   - Same directory as the application

2. **Template Format**: Templates should include:
   - Standard SPP data columns
   - Proper headers and formatting
   - Any custom calculations or charts you need

3. **Configure in App**: Use the Template Configuration section to:
   - Browse and select your template file
   - Choose between .xlsx and .xlsm formats
   - Verify template path and format

## 🚀 Usage Instructions

### First Time Setup
1. Launch `SPP_Enhanced.exe`
2. Navigate to the "Template Configuration" section
3. Click "Browse Template" to select your Excel template
4. Choose your preferred format (.xlsx or .xlsm)
5. Save configuration

### Daily Operations
1. Launch the application
2. Enter your search criteria (same as before)
3. Click "Generate SPP Report"
4. Your report will be generated using your configured template
5. Reports are saved with timestamp for easy organization

## 🔧 Configuration Options

### Template Settings
- **Template Path**: Full path to your Excel template file
- **Format**: Choose between Excel (.xlsx) or Macro-Enabled Excel (.xlsm)
- **Auto-Discovery**: Enable automatic template discovery
- **Fallback Mode**: Use standard Excel if template unavailable

### Advanced Options
- **Background Processing**: Keep UI responsive during operations
- **Enhanced Logging**: Detailed operation logs for troubleshooting
- **Connection Retry**: Automatic retry for database connections
- **Progress Feedback**: Real-time status updates

## 📈 Performance Tips

1. **Template Location**: Keep templates in OneDrive for automatic discovery
2. **File Size**: Smaller template files load faster
3. **Network**: Ensure stable internet connection for Snowflake access
4. **Resources**: Close unnecessary applications for optimal performance

## 🔒 Security Notes

- Application runs with standard user permissions
- No admin rights required
- Database credentials handled securely
- Templates remain on local system
- No data transmitted except to authorized Snowflake database

## 📞 Support

For technical support or questions:
1. Check `TROUBLESHOOTING.md` for common solutions
2. Review application logs in the same directory
3. Contact your IT team or application administrator
4. Refer to `TEMPLATE_GUIDE.md` for template-specific questions

## 🔄 Updates

This is a self-contained executable. Updates will be provided as new executable files. Simply replace the old .exe with the new version while preserving your template configurations.

---

**SPP Enhanced Automation Tool v2.0**  
*Empowering teams with flexible, user-configurable reporting*

Built with ❤️ for enhanced productivity and user experience.
