# SPP Automation Package - Installation Guide

## 📋 System Requirements
- Windows 10/11
- Python 3.8 or higher
- Internet connection for Snowflake
- HD Supply network access
- Excel 2016 or higher

## 🚀 Installation Steps

### Option A: Quick Install (Recommended)
1. Extract the package to `C:\SPP_Automation\`
2. Run `install_requirements.bat` as Administrator
3. Edit `config.ini` with your email
4. Run `Quick_Start.bat`

### Option B: Manual Install
1. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Configure your settings in `config.ini`
3. Test connection: `python test_connection.py`
4. Run GUI: `python spp_gui.py`

## ⚙️ Configuration

### config.ini Settings
```ini
[SNOWFLAKE]
user = your_email@hdsupply.com    # <- UPDATE THIS
account = HDSUPPLY-DATA           # (pre-configured)
authenticator = externalbrowser   # (pre-configured)

[PATHS]
template_path = path\to\your\template.xlsm    # <- UPDATE THIS
output_directory = Output                     # (or custom path)
```

## 🔧 Troubleshooting

### Connection Issues
- Ensure you have Snowflake access through HD Supply
- Check your email address in config.ini
- Try running `test_connection.py`

### Template Issues
- Verify template path exists
- Ensure template has correct sheet names:
  - "METRIC DATA"
  - "ASN Data"
  - "Metric Pivots" (for macro)

### Permission Issues
- Run as Administrator if needed
- Check Windows Defender/antivirus settings
- Ensure Python is in system PATH

## 📞 Support
1. Check `spp_automation.log` for errors
2. Review examples in `examples/` folder
3. Contact SPP team for assistance

## 🔄 Updates
To update the tool:
1. Backup your `config.ini`
2. Extract new package files
3. Restore your `config.ini`
4. Run `test_connection.py` to verify
