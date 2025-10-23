"""
SPP Automation Tool - Fixed Unicode Version
Now with proper Windows console encoding handling.

Developer: Ben F. Benjamaa
Manager: Lauren B. Trapani
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import re
import sys
import traceback
import logging
from datetime import datetime

# Set console encoding for Windows compatibility - simplified approach
if os.name == 'nt':
    import codecs
    # Just ensure we handle encoding issues gracefully
    pass

# Set up logging with UTF-8 encoding
log_file = f"spp_debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def safe_log_error(title, message):
    """Log error safely without Unicode issues."""
    safe_title = str(title).encode('ascii', 'replace').decode('ascii')
    safe_message = str(message).encode('ascii', 'replace').decode('ascii')
    
    logger.error(f"{safe_title}: {safe_message}")
    logger.error(f"Traceback: {traceback.format_exc()}")
    
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(safe_title, f"{safe_message}\n\nCheck {log_file} for details.")
        root.destroy()
    except Exception as e:
        print(f"ERROR: {safe_title} - {safe_message}")

class SPPAutomationGUI:
    def __init__(self, root):
        logger.info("Initializing SPP Automation GUI")
        
        self.root = root
        self.root.title("SPP Metric Automation Tool - HD Supply")
        self.root.geometry("900x800")
        
        # HD Supply colors
        self.colors = {
            'primary_black': '#000000',
            'secondary_black': '#1A1A1A',
            'accent_yellow': '#FFFF00',
            'dim_yellow': '#CCCC00',
            'success_green': '#00FF00',
            'dark_green': '#006600',
            'white': '#FFFFFF',
            'error_red': '#FF0000'
        }
        
        # Configure root styling
        self.root.configure(bg=self.colors['primary_black'])
        
        # Initialize variables
        self.automation = None
        self.user_email = ""
        self.authenticated = False
        
        try:
            self.setup_styles()
            self.create_widgets()
            self.center_window()
            logger.info("GUI initialization completed successfully")
        except Exception as e:
            logger.error(f"GUI initialization failed: {str(e)}")
            raise
    
    def setup_styles(self):
        """Configure custom styles for the application."""
        logger.info("Setting up GUI styles")
        style = ttk.Style()
        
        # Configure styles with safe colors
        style.configure('Header.TFrame', background=self.colors['primary_black'])
        style.configure('Main.TFrame', background=self.colors['primary_black'])
        style.configure('Footer.TFrame', background=self.colors['secondary_black'])
        style.configure('Auth.TFrame', background=self.colors['secondary_black'])
        
        style.configure('Header.TLabel', 
                       background=self.colors['primary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 20, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=self.colors['primary_black'],
                       foreground=self.colors['dim_yellow'],
                       font=('Arial', 12))
        
        style.configure('Credit.TLabel',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 10, 'bold'))
        
        style.configure('Field.TLabel',
                       background=self.colors['primary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 11, 'bold'))
        
        style.configure('Auth.TLabel',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 12, 'bold'))
        
        style.configure('Primary.TButton',
                       background=self.colors['success_green'],
                       foreground=self.colors['primary_black'],
                       font=('Arial', 12, 'bold'),
                       padding=(20, 10))
        
        style.map('Primary.TButton',
                 background=[('active', self.colors['dark_green'])])
        
        style.configure('Secondary.TButton',
                       background=self.colors['success_green'],
                       foreground=self.colors['primary_black'],
                       font=('Arial', 10, 'bold'),
                       padding=(10, 5))
        
        style.map('Secondary.TButton',
                 background=[('active', self.colors['dark_green'])])
        
        style.configure('Custom.TEntry',
                       fieldbackground=self.colors['white'],
                       foreground=self.colors['primary_black'],
                       borderwidth=2,
                       font=('Arial', 11))
        
        style.configure('Custom.Horizontal.TProgressbar',
                       background=self.colors['success_green'],
                       troughcolor=self.colors['secondary_black'])
    
    def center_window(self):
        """Center the window on the screen."""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.root.winfo_screenheight() // 2) - (800 // 2)
        self.root.geometry(f"900x800+{x}+{y}")
    
    def validate_email(self, email):
        """Validate email format."""
        pattern = r'^[a-zA-Z0-9._%+-]+@hdsupply\.com$'
        return re.match(pattern, email) is not None
    
    def authenticate_user(self):
        """Authenticate user with Snowflake SSO."""
        logger.info("Starting user authentication")
        email = self.email_entry.get().strip()
        
        if not email:
            messagebox.showerror("Error", "Please enter your HD Supply email address")
            return
        
        if not self.validate_email(email):
            messagebox.showerror("Error", "Please enter a valid HD Supply email address (must end with @hdsupply.com)")
            return
        
        self.user_email = email
        self.auth_button.config(text="Authenticating...", state='disabled')
        
        # Run authentication in thread
        thread = threading.Thread(target=self._authenticate_thread)
        thread.daemon = True
        thread.start()
    
    def _authenticate_thread(self):
        """Run authentication in separate thread."""
        try:
            logger.info(f"Attempting authentication for {self.user_email}")
            
            # Import the fixed automation class
            from spp_metric_automation_fixed import SPPMetricAutomationFixed
            
            # Create temporary automation instance for authentication
            temp_automation = SPPMetricAutomationFixed("config.ini", user_email=self.user_email)
            
            if temp_automation.connect_to_snowflake():
                self.authenticated = True
                logger.info("Authentication successful")
                self.root.after(0, lambda: self.auth_status_label.config(
                    text=f"[OK] Authenticated as: {self.user_email}",
                    foreground=self.colors['success_green']))
                self.root.after(0, lambda: self.auth_button.config(
                    text="Re-authenticate", state='normal'))
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success", f"Successfully authenticated as {self.user_email}"))
                
                # Close the connection
                if temp_automation.connection:
                    temp_automation.connection.close()
            else:
                self.authenticated = False
                logger.warning("Authentication failed")
                self.root.after(0, lambda: self.auth_status_label.config(
                    text="[ERROR] Authentication Failed",
                    foreground=self.colors['error_red']))
                self.root.after(0, lambda: self.auth_button.config(
                    text="Authenticate", state='normal'))
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", "Authentication failed. Please check your credentials and try again."))
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            self.authenticated = False
            self.root.after(0, lambda: self.auth_status_label.config(
                text="[ERROR] Authentication Error",
                foreground=self.colors['error_red']))
            self.root.after(0, lambda: self.auth_button.config(
                text="Authenticate", state='normal'))
            self.root.after(0, lambda: messagebox.showerror(
                "Error", f"Authentication error: {str(e)}"))
    
    def create_widgets(self):
        logger.info("Creating GUI widgets")
        
        # Header frame
        header_frame = ttk.Frame(self.root, style='Header.TFrame', padding="20")
        header_frame.pack(fill=tk.X)
        
        # Title and branding  
        title_label = ttk.Label(header_frame, text="SPP METRIC AUTOMATION", 
                               style='Header.TLabel')
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(header_frame, text="HD Supply Vendor Performance Analytics", 
                                  style='Subtitle.TLabel')
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Version info
        version_label = ttk.Label(header_frame, text="Version 2.0 | Enterprise Edition", 
                                 style='Subtitle.TLabel')
        version_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Authentication section
        auth_frame = ttk.Frame(self.root, style='Auth.TFrame', padding="20")
        auth_frame.pack(fill=tk.X)
        
        auth_title = ttk.Label(auth_frame, text="SNOWFLAKE SSO AUTHENTICATION", 
                              style='Auth.TLabel')
        auth_title.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 15))
        
        # Email input
        ttk.Label(auth_frame, text="HD Supply Email:", style='Auth.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.email_entry = ttk.Entry(auth_frame, width=40, style='Custom.TEntry')
        self.email_entry.grid(row=1, column=1, pady=8, padx=(0, 10), sticky=tk.W)
        self.email_entry.insert(0, "your.name@hdsupply.com")
        
        # Authenticate button
        self.auth_button = ttk.Button(auth_frame, text="Authenticate",
                                     command=self.authenticate_user,
                                     style='Secondary.TButton')
        self.auth_button.grid(row=1, column=2, pady=8)
        
        # Authentication status - using safe ASCII characters
        self.auth_status_label = ttk.Label(auth_frame, text="[X] Not Authenticated", 
                                          style='Auth.TLabel')
        self.auth_status_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        # Main content frame
        main_frame = ttk.Frame(self.root, style='Main.TFrame', padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configuration section
        config_label = ttk.Label(main_frame, text="AUTOMATION CONFIGURATION", 
                                style='Field.TLabel')
        config_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))
        
        # Vendor Numbers
        ttk.Label(main_frame, text="Vendor Numbers:", style='Field.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.vendor_entry = ttk.Entry(main_frame, width=50, style='Custom.TEntry')
        self.vendor_entry.grid(row=1, column=1, pady=8, sticky="we")
        self.vendor_entry.insert(0, "52889")
        
        # Helper text
        helper_label = tk.Label(main_frame, text="(comma-separated, e.g., 52889, 11833)", 
                               font=('Arial', 9), bg=self.colors['primary_black'], 
                               fg=self.colors['dim_yellow'])
        helper_label.grid(row=2, column=1, sticky=tk.W, pady=(2, 8))
        
        # Report Month
        ttk.Label(main_frame, text="Report Month:", style='Field.TLabel').grid(
            row=3, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.month_entry = ttk.Entry(main_frame, width=50, style='Custom.TEntry')
        self.month_entry.grid(row=3, column=1, pady=8, sticky="we")
        self.month_entry.insert(0, "FY2025-APR")
        
        month_helper = tk.Label(main_frame, text="(fiscal year format, e.g., FY2025-APR)", 
                               font=('Arial', 9), bg=self.colors['primary_black'], 
                               fg=self.colors['dim_yellow'])
        month_helper.grid(row=4, column=1, sticky=tk.W, pady=(2, 8))
        
        # Date Filter
        ttk.Label(main_frame, text="Date Filter:", style='Field.TLabel').grid(
            row=5, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.date_entry = ttk.Entry(main_frame, width=50, style='Custom.TEntry')
        self.date_entry.grid(row=5, column=1, pady=8, sticky="we")
        self.date_entry.insert(0, "202504")
        
        date_helper = tk.Label(main_frame, text="(YYYYMM format for ASN data)", 
                              font=('Arial', 9), bg=self.colors['primary_black'], 
                              fg=self.colors['dim_yellow'])
        date_helper.grid(row=6, column=1, sticky=tk.W, pady=(2, 8))
        
        # Action buttons frame
        button_frame = ttk.Frame(main_frame, style='Main.TFrame')
        button_frame.grid(row=8, column=0, columnspan=2, pady=30)
        
        # Primary action button
        self.run_button = ttk.Button(button_frame, text=">> RUN AUTOMATION <<", 
                                    style='Primary.TButton',
                                    command=self.run_automation)
        self.run_button.grid(row=0, column=1, padx=10)
        
        # Secondary buttons
        ttk.Button(button_frame, text="Test Connection", style='Secondary.TButton',
                  command=self.test_connection).grid(row=0, column=0, padx=10)
        
        ttk.Button(button_frame, text="Test ASN Query", style='Secondary.TButton',
                  command=self.test_asn_query).grid(row=0, column=2, padx=10)
        
        ttk.Button(button_frame, text="View Output", style='Secondary.TButton',
                  command=self.view_output_folder).grid(row=0, column=3, padx=10)
        
        # Status text area
        status_frame = ttk.Frame(main_frame, style='Main.TFrame')
        status_frame.grid(row=11, column=0, columnspan=2, sticky="nsew", pady=10)
        
        self.status_text = tk.Text(status_frame, height=8, width=70, 
                                  bg=self.colors['primary_black'], fg=self.colors['accent_yellow'],
                                  font=('Consolas', 10), wrap=tk.WORD)
        self.status_text.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Footer frame with credits
        footer_frame = ttk.Frame(self.root, style='Footer.TFrame', padding="15")
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Developer credits
        credits_frame = ttk.Frame(footer_frame, style='Footer.TFrame')
        credits_frame.pack(fill=tk.X)
        
        dev_label = ttk.Label(credits_frame, text="Developer: Ben F. Benjamaa", 
                             style='Credit.TLabel')
        dev_label.pack(side=tk.LEFT)
        
        manager_label = ttk.Label(credits_frame, text="Manager: Lauren B. Trapani", 
                                 style='Credit.TLabel')
        manager_label.pack(side=tk.RIGHT)
        
        # Copyright
        copyright_label = ttk.Label(footer_frame, 
                                   text="(C) 2025 HD Supply, Inc. | Supply Performance & Procurement",
                                   background=self.colors['secondary_black'],
                                   foreground=self.colors['white'],
                                   font=('Arial', 9))
        copyright_label.pack(pady=(5, 0))
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        # Add debug info to status
        self.log_message("SPP Automation Tool v2.0 Ready")
        self.log_message(f"Debug log: {log_file}")
        self.log_message("Please authenticate to begin")
    
    def view_output_folder(self):
        """Open output folder."""
        try:
            output_dir = os.path.join(os.getcwd(), "Output")
            os.makedirs(output_dir, exist_ok=True)
            
            if os.name == 'nt':  # Windows
                os.startfile(output_dir)
            else:
                import subprocess
                subprocess.run(['xdg-open', output_dir])
                
        except Exception as e:
            logger.error(f"Failed to open output folder: {str(e)}")
            messagebox.showerror("Error", f"Cannot open output folder: {str(e)}")
    
    def test_connection(self):
        """Test Snowflake connection."""
        if not self.authenticated:
            messagebox.showerror("Error", "Please authenticate first before testing connection")
            return
        
        self.log_message("Testing Snowflake connection...")
        
        # Run test in thread
        thread = threading.Thread(target=self._test_connection_thread)
        thread.daemon = True
        thread.start()
    
    def _test_connection_thread(self):
        """Test connection in thread."""
        try:
            from spp_metric_automation_fixed import SPPMetricAutomationFixed
            
            self.root.after(0, lambda: self.log_message("Creating test connection..."))
            test_automation = SPPMetricAutomationFixed("config.ini", user_email=self.user_email)
            
            if test_automation.connect_to_snowflake():
                self.root.after(0, lambda: self.log_message("Connection test successful!"))
                self.root.after(0, lambda: messagebox.showinfo("Test Result", "Connection test successful!"))
                
                if test_automation.connection:
                    test_automation.connection.close()
            else:
                self.root.after(0, lambda: self.log_message("Connection test failed"))
                self.root.after(0, lambda: messagebox.showerror("Test Result", "Connection test failed"))
                
        except Exception as e:
            error_msg = f"Connection test error: {str(e)}"
            self.root.after(0, lambda: self.log_message(error_msg))
            self.root.after(0, lambda: messagebox.showerror("Test Error", error_msg))
    
    def test_asn_query(self):
        """Test ASN query specifically to debug data pulling issues."""
        if not self.authenticated:
            messagebox.showerror("Error", "Please authenticate first before testing ASN query")
            return
        
        self.log_message("Testing ASN query...")
        
        # Run ASN test in thread
        thread = threading.Thread(target=self._test_asn_query_thread)
        thread.daemon = True
        thread.start()
    
    def _test_asn_query_thread(self):
        """Test ASN query in separate thread."""
        try:
            from spp_metric_automation_fixed import SPPMetricAutomationFixed
            
            # Get current parameters from GUI
            vendor_numbers = [v.strip() for v in self.vendor_entry.get().split(',')]
            date_filter = self.date_entry.get().strip()
            
            self.root.after(0, lambda: self.log_message(f"Testing ASN for vendors: {vendor_numbers}"))
            self.root.after(0, lambda: self.log_message(f"Date filter: {date_filter}"))
            
            # Create test automation instance
            test_automation = SPPMetricAutomationFixed("config.ini", user_email=self.user_email)
            
            # Run ASN test
            df_asn = test_automation.test_asn_query_standalone(vendor_numbers, date_filter)
            
            if df_asn is not None and not df_asn.empty:
                self.root.after(0, lambda: self.log_message(f"ASN test successful! Found {len(df_asn)} records"))
                self.root.after(0, lambda: messagebox.showinfo("ASN Test Result", 
                    f"ASN query successful!\nFound {len(df_asn)} records\nCheck log for details"))
            else:
                self.root.after(0, lambda: self.log_message("ASN test: No records found"))
                self.root.after(0, lambda: messagebox.showwarning("ASN Test Result", 
                    f"ASN query executed but no records found.\nCheck:\n- Vendor number: {vendor_numbers}\n- Date filter: {date_filter}%"))
                
        except Exception as e:
            error_msg = f"ASN test error: {str(e)}"
            self.root.after(0, lambda: self.log_message(error_msg))
            self.root.after(0, lambda: messagebox.showerror("ASN Test Error", error_msg))
    
    def run_automation(self):
        """Run the full automation."""
        if not self.authenticated:
            messagebox.showerror("Authentication Required", 
                               "Please authenticate with your HD Supply email before running automation")
            return
        
        # Get parameters
        vendor_numbers = self.vendor_entry.get().strip()
        report_month = self.month_entry.get().strip()
        date_filter = self.date_entry.get().strip()
        
        if not all([vendor_numbers, report_month, date_filter]):
            messagebox.showerror("Missing Parameters", "Please fill in all required fields")
            return
        
        # Disable run button
        self.run_button.config(text="Running...", state='disabled')
        
        # Run automation in thread
        thread = threading.Thread(target=self._run_automation_thread, 
                                 args=(vendor_numbers, report_month, date_filter))
        thread.daemon = True
        thread.start()
    
    def _run_automation_thread(self, vendor_numbers, report_month, date_filter):
        """Run automation in separate thread."""
        try:
            self.root.after(0, lambda: self.log_message("Starting automation..."))
            
            from spp_metric_automation_fixed import SPPMetricAutomationFixed
            
            # Create automation instance
            automation = SPPMetricAutomationFixed("config.ini", user_email=self.user_email)
            
            # Parse vendor numbers
            vendors = [v.strip() for v in vendor_numbers.split(',')]
            
            self.root.after(0, lambda: self.log_message(f"Processing {len(vendors)} vendor(s)"))
            
            # Run automation
            result = automation.run_full_automation(vendors, report_month, date_filter)
            
            if result and "successfully" in result.lower():
                self.root.after(0, lambda: self.log_message("Automation completed successfully!"))
                self.root.after(0, lambda: messagebox.showinfo("Success", 
                    "Automation completed successfully! Check the Output folder for results."))
            else:
                self.root.after(0, lambda: self.log_message("Automation completed with issues"))
                self.root.after(0, lambda: messagebox.showwarning("Warning", 
                    f"Automation completed but may have issues: {result}"))
                
        except Exception as e:
            error_msg = f"Automation error: {str(e)}"
            self.root.after(0, lambda: self.log_message(error_msg))
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            
        finally:
            self.root.after(0, lambda: self.run_button.config(text=">> RUN AUTOMATION <<", state='normal'))
    
    def log_message(self, message):
        """Add message to status text area."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}"
        self.status_text.insert(tk.END, full_message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
        logger.info(message)

def main():
    """Main entry point with comprehensive error handling."""
    logger.info("Starting SPP Automation Tool - Fixed Version")
    
    try:
        # Test imports first
        logger.info("Testing imports...")
        import pandas as pd
        logger.info("pandas: OK")
        
        import snowflake.connector
        logger.info("snowflake: OK")
        
        import openpyxl
        logger.info("openpyxl: OK")
        
        # Create GUI
        logger.info("Creating main window...")
        root = tk.Tk()
        
        logger.info("Initializing application...")
        app = SPPAutomationGUI(root)
        
        logger.info("Starting GUI...")
        root.mainloop()
        
        logger.info("Application closed normally")
        
    except ImportError as e:
        error_msg = f"Missing module: {str(e)}"
        safe_log_error("Import Error", error_msg)
        
    except Exception as e:
        error_msg = f"Application error: {str(e)}"
        safe_log_error("Application Error", error_msg)
        
    finally:
        logger.info("Application shutdown complete")

if __name__ == "__main__":
    main()
