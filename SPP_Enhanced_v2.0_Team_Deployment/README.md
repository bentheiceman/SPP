# SPP Metric Automation Tool Enhanced v2.0
## HD Supply Chain Excellence

### Overview
The SPP Metric Automation Tool Enhanced provides advanced reporting capabilities with user-configurable Excel templates, comprehensive error handling, and an intuitive interface.

### New Features in v2.0
- **User-Configurable Templates**: Link your own Excel template files for consistent formatting
- **Template Auto-Discovery**: Automatic search in common locations (OneDrive, Documents)
- **Dual Output Modes**: Standard Excel (.xlsx) or Macro-Enabled (.xlsm) formats
- **Enhanced Error Handling**: Graceful fallbacks when templates are unavailable
- **Improved UI**: Modern, branded interface with comprehensive configuration options
- **Background Processing**: Non-blocking operations with progress indicators

### Quick Start
1. **Launch**: Double-click `SPP_Automation_Tool_Enhanced.exe`
2. **Authenticate**: Enter your HD Supply email and click "Authenticate"
3. **Configure Template** (Optional):
   - Check "Use Excel Template for Output"
   - Browse for your template file or use auto-discovery
   - Choose output format (.xlsx or .xlsm)
4. **Enter Parameters**:
   - Vendor Numbers (comma-separated)
   - Report Month (e.g., FY2025-APR)
   - Date Filter (YYYYMM format)
5. **Generate Report**: Click "🚀 Generate SPP Report"

### Template Configuration
The tool supports flexible template configuration:

#### Template Discovery
The tool searches for templates in these locations (in order):
1. User-specified path (if configured)
2. OneDrive Documents folder
3. Local Documents folder
4. Application directory
5. Current working directory

#### Template Requirements
- Must be Excel format (.xlsx or .xlsm)
- Should contain sheets named "Tab1_Basic_Metrics" and "Tab2_ASN_Data"
- Headers will be preserved, data populated starting from row 2

#### Fallback Behavior
If no template is found, the tool automatically creates a standard Excel file with all data in separate tabs.

### System Requirements
- Windows 10/11 (64-bit)
- 4GB RAM minimum
- Internet connection for Snowflake authentication
- HD Supply network access

### Configuration Files
- `template_config.json`: Template settings and paths
- Application logs: Automatic logging with timestamps
- Output folder: Created automatically in same directory

### Troubleshooting

#### Authentication Issues
- Ensure you're on HD Supply network
- Check email format is correct
- Try clearing browser cookies if authentication fails

#### Template Issues
- Verify template file path is accessible
- Check file permissions
- Tool will fallback to standard Excel if template unavailable

#### Data Issues
- Verify vendor numbers are correct
- Check date formats (FY2025-APR, 202507)
- Use "Test Query" button to validate parameters

#### Performance
- Large vendor lists may take several minutes
- Use date filters to limit data volume
- Monitor activity log for progress updates

### Support Information
- **Developer**: Ben F. Benjamaa
- **Manager**: Lauren B. Trapani
- **Department**: HD Supply Chain Excellence
- **Version**: 2.0.0 Enhanced
- **Build Date**: 2025-10-23

### Change Log
#### Version 2.0.0 (Enhanced)
- Added user-configurable template system
- Implemented template auto-discovery
- Enhanced error handling and fallback mechanisms
- Improved UI with template configuration options
- Added comprehensive logging
- Optimized performance and memory usage
- Enhanced deployment packaging

#### Version 1.0.0
- Initial release
- Basic Snowflake connectivity
- Standard Excel output
- Simple GUI interface

For technical support or feature requests, contact the development team.
