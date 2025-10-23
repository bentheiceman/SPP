"""
SPP Metric Automation - Enhanced GUI with Template Configuration
Modern, branded interface with user-configurable template paths.

Developer: Ben F. Benjamaa
Manager: Lauren B. Trapani
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import json
import traceback
import logging
from datetime import datetime
from pathlib import Path
import re
from typing import Optional

# Import our enhanced automation module
try:
    from spp_automation_enhanced import SPPAutomationEnhanced
except ImportError as e:
    print(f"Error importing automation module: {e}")
    SPPAutomationEnhanced = None

# Set up logging
log_file = f"spp_gui_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class SPPEnhancedGUI:
    """Enhanced SPP Automation GUI with template configuration capabilities."""
    
    def __init__(self, root):
        logger.info("Initializing Enhanced SPP Automation GUI")
        
        self.root = root
        self.root.title("SPP Metric Automation Tool - HD Supply Enhanced")
        self.root.geometry("1000x900")
        
        # HD Supply Brand Colors
        self.colors = {
            'primary_black': '#000000',
            'secondary_black': '#1A1A1A',
            'accent_yellow': '#FFFF00',
            'dim_yellow': '#CCCC00',
            'success_green': '#00FF00',
            'dark_green': '#006600',
            'white': '#FFFFFF',
            'error_red': '#FF0000',
            'info_blue': '#0066CC'
        }
        
        # Configure root
        self.root.configure(bg=self.colors['primary_black'])
        
        # Initialize variables
        self.automation = None  # Will be SPPAutomationEnhanced instance
        self.user_email = ""
        self.authenticated = False
        self.template_config = {}
        
        # GUI State variables
        self.vendor_var = tk.StringVar(value="52889")
        self.month_var = tk.StringVar(value="FY2025-APR")
        self.date_var = tk.StringVar(value="202507")
        self.email_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.use_template_var = tk.BooleanVar(value=False)
        self.output_format_var = tk.StringVar(value="xlsx")
        
        try:
            self.setup_styles()
            self.create_widgets()
            self.center_window()
            self.load_template_config()
            logger.info("GUI initialization completed successfully")
        except Exception as e:
            logger.error(f"GUI initialization failed: {str(e)}")
            self.show_error("Initialization Error", f"Failed to initialize GUI: {str(e)}")
    
    def setup_styles(self):
        """Configure custom styles for the application."""
        logger.info("Setting up GUI styles")
        style = ttk.Style()
        
        # Configure frame styles
        style.configure('Header.TFrame', background=self.colors['primary_black'])
        style.configure('Main.TFrame', background=self.colors['primary_black'])
        style.configure('Section.TFrame', background=self.colors['secondary_black'], relief='ridge', borderwidth=2)
        style.configure('Footer.TFrame', background=self.colors['secondary_black'])
        
        # Configure label styles
        style.configure('Header.TLabel', 
                       background=self.colors['primary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 24, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=self.colors['primary_black'],
                       foreground=self.colors['dim_yellow'],
                       font=('Arial', 14))
        
        style.configure('Section.TLabel',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 16, 'bold'))
        
        style.configure('Field.TLabel',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 12, 'bold'))
        
        style.configure('Info.TLabel',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['dim_yellow'],
                       font=('Arial', 10))
        
        # Configure button styles
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
        
        # Configure entry styles
        style.configure('Custom.TEntry',
                       fieldbackground=self.colors['white'],
                       foreground=self.colors['primary_black'],
                       font=('Arial', 11))
        
        # Configure checkbutton styles
        style.configure('Custom.TCheckbutton',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 11, 'bold'))
        
        # Configure radiobutton styles
        style.configure('Custom.TRadiobutton',
                       background=self.colors['secondary_black'],
                       foreground=self.colors['accent_yellow'],
                       font=('Arial', 11))
        
        # Configure progress bar
        style.configure('Custom.Horizontal.TProgressbar',
                       background=self.colors['success_green'],
                       troughcolor=self.colors['secondary_black'])
    
    def center_window(self):
        """Center the window on the screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def _sanitize_email(self, email: str) -> str:
        """Normalize/sanitize email input to avoid hidden characters causing validation to fail."""
        if not isinstance(email, str):
            return ""
        # Remove common zero-width and whitespace characters
        zero_width = "\u200b\u200c\u200d\ufeff\u2060"
        trans_table = dict.fromkeys(map(ord, zero_width), None)
        cleaned = email.translate(trans_table)
        # Collapse whitespace and strip
        cleaned = " ".join(cleaned.split()).strip()
        return cleaned

    def validate_email(self, email: str) -> bool:
        """Validate email format robustly.

        Accepts common email formats and ignores case. Falls back to a simple
        check for the presence of a single '@' and a dot in the domain to avoid
        blocking valid users due to over-strict patterns or hidden characters.
        """
        cleaned = self._sanitize_email(email)
        # Primary regex (case-insensitive)
        pattern = re.compile(r'^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$', re.IGNORECASE)
        if pattern.match(cleaned):
            return True
        # Permissive fallback: basic structure check
        try:
            local, domain = cleaned.split('@', 1)
            if local and domain and '.' in domain and not domain.startswith('.') and not domain.endswith('.'):
                return True
        except ValueError:
            pass
        logger.debug(f"Email validation failed for input: {repr(email)} | sanitized: {repr(cleaned)}")
        return False
    
    def create_widgets(self):
        """Create all GUI widgets."""
        # Main container with scrollbar
        main_canvas = tk.Canvas(self.root, bg=self.colors['primary_black'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas, style='Main.TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # Header Section
        header_frame = ttk.Frame(scrollable_frame, style='Header.TFrame', padding=20)
        header_frame.pack(fill='x', padx=10, pady=(10, 0))
        
        ttk.Label(header_frame, text="SPP Metric Automation Tool", 
                 style='Header.TLabel').pack()
        ttk.Label(header_frame, text="Enhanced Version v2.2 with Template Configuration", 
                 style='Subtitle.TLabel').pack()
        
        # Authentication Section
        auth_frame = ttk.Frame(scrollable_frame, style='Section.TFrame', padding=15)
        auth_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(auth_frame, text="Authentication", style='Section.TLabel').pack()
        
        email_frame = ttk.Frame(auth_frame, style='Section.TFrame')
        email_frame.pack(fill='x', pady=5)
        
        ttk.Label(email_frame, text="HD Supply Email:", style='Field.TLabel').pack(side='left')
        email_entry = ttk.Entry(email_frame, textvariable=self.email_var, style='Custom.TEntry', width=30)
        email_entry.pack(side='left', padx=(10, 5))
        
        auth_btn = ttk.Button(email_frame, text="Authenticate", style='Secondary.TButton',
                             command=self.authenticate_user)
        auth_btn.pack(side='left', padx=5)
        
        self.auth_status = ttk.Label(email_frame, text="Not authenticated", style='Info.TLabel')
        self.auth_status.pack(side='left', padx=10)
        
        # Template Configuration Section
        template_frame = ttk.Frame(scrollable_frame, style='Section.TFrame', padding=15)
        template_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(template_frame, text="Template Configuration", style='Section.TLabel').pack()
        
        # Template enable/disable
        template_check_frame = ttk.Frame(template_frame, style='Section.TFrame')
        template_check_frame.pack(fill='x', pady=5)
        
        template_check = ttk.Checkbutton(template_check_frame, 
                                       text="Use Excel Template for Output",
                                       variable=self.use_template_var,
                                       style='Custom.TCheckbutton',
                                       command=self.on_template_toggle)
        template_check.pack(side='left')
        
        # Template path selection
        template_path_frame = ttk.Frame(template_frame, style='Section.TFrame')
        template_path_frame.pack(fill='x', pady=5)
        
        ttk.Label(template_path_frame, text="Template File:", style='Field.TLabel').pack(side='left')
        self.template_entry = ttk.Entry(template_path_frame, textvariable=self.template_path_var, 
                                      style='Custom.TEntry', width=50)
        self.template_entry.pack(side='left', padx=(10, 5))
        
        browse_btn = ttk.Button(template_path_frame, text="Browse", style='Secondary.TButton',
                               command=self.browse_template)
        browse_btn.pack(side='left', padx=5)
        
        # Output format selection
        format_frame = ttk.Frame(template_frame, style='Section.TFrame')
        format_frame.pack(fill='x', pady=5)
        
        ttk.Label(format_frame, text="Output Format:", style='Field.TLabel').pack(side='left')
        
        format_xlsx = ttk.Radiobutton(format_frame, text="Standard Excel (.xlsx)", 
                                    variable=self.output_format_var, value="xlsx",
                                    style='Custom.TRadiobutton')
        format_xlsx.pack(side='left', padx=(10, 5))
        
        format_xlsm = ttk.Radiobutton(format_frame, text="Macro-Enabled (.xlsm)", 
                                    variable=self.output_format_var, value="xlsm",
                                    style='Custom.TRadiobutton')
        format_xlsm.pack(side='left', padx=5)
        
        # Template info
        self.template_info = ttk.Label(template_frame, 
                                     text="Template allows consistent formatting across all reports",
                                     style='Info.TLabel')
        self.template_info.pack(pady=5)
        
        # Data Input Section
        input_frame = ttk.Frame(scrollable_frame, style='Section.TFrame', padding=15)
        input_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(input_frame, text="Report Parameters", style='Section.TLabel').pack()
        
        # Vendor numbers
        vendor_frame = ttk.Frame(input_frame, style='Section.TFrame')
        vendor_frame.pack(fill='x', pady=5)
        
        ttk.Label(vendor_frame, text="Vendor Numbers:", style='Field.TLabel').pack(side='left')
        vendor_entry = ttk.Entry(vendor_frame, textvariable=self.vendor_var, style='Custom.TEntry', width=30)
        vendor_entry.pack(side='left', padx=(10, 5))
        ttk.Label(vendor_frame, text="(comma-separated)", style='Info.TLabel').pack(side='left', padx=5)
        
        # Report month
        month_frame = ttk.Frame(input_frame, style='Section.TFrame')
        month_frame.pack(fill='x', pady=5)
        
        ttk.Label(month_frame, text="Report Month:", style='Field.TLabel').pack(side='left')
        month_entry = ttk.Entry(month_frame, textvariable=self.month_var, style='Custom.TEntry', width=20)
        month_entry.pack(side='left', padx=(10, 5))
        ttk.Label(month_frame, text="(e.g., FY2025-APR)", style='Info.TLabel').pack(side='left', padx=5)
        
        # Date filter
        date_frame = ttk.Frame(input_frame, style='Section.TFrame')
        date_frame.pack(fill='x', pady=5)
        
        ttk.Label(date_frame, text="Date Filter:", style='Field.TLabel').pack(side='left')
        date_entry = ttk.Entry(date_frame, textvariable=self.date_var, style='Custom.TEntry', width=20)
        date_entry.pack(side='left', padx=(10, 5))
        ttk.Label(date_frame, text="(YYYYMM format)", style='Info.TLabel').pack(side='left', padx=5)
        
        # Control Buttons Section
        control_frame = ttk.Frame(scrollable_frame, style='Section.TFrame', padding=15)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(control_frame, text="Actions", style='Section.TLabel').pack()
        
        button_frame = ttk.Frame(control_frame, style='Section.TFrame')
        button_frame.pack(pady=10)
        
        test_conn_btn = ttk.Button(button_frame, text="Test Connection", 
                                  style='Secondary.TButton', command=self.test_connection)
        test_conn_btn.pack(side='left', padx=5)
        
        test_query_btn = ttk.Button(button_frame, text="Test Query", 
                                   style='Secondary.TButton', command=self.test_query)
        test_query_btn.pack(side='left', padx=5)
        
        save_template_btn = ttk.Button(button_frame, text="Save Template Config", 
                                      style='Secondary.TButton', command=self.save_template_config)
        save_template_btn.pack(side='left', padx=5)
        
        output_btn = ttk.Button(button_frame, text="Open Output Folder", 
                               style='Secondary.TButton', command=self.open_output_folder)
        output_btn.pack(side='left', padx=5)
        
        # Main run button
        run_frame = ttk.Frame(control_frame, style='Section.TFrame')
        run_frame.pack(pady=10)
        
        self.run_btn = ttk.Button(run_frame, text="🚀 Generate SPP Report", 
                                 style='Primary.TButton', command=self.run_automation)
        self.run_btn.pack()
        
        # Progress bar
        self.progress = ttk.Progressbar(control_frame, style='Custom.Horizontal.TProgressbar', 
                                       mode='indeterminate', length=400)
        self.progress.pack(pady=10)
        
        # Log output
        log_frame = ttk.Frame(scrollable_frame, style='Section.TFrame', padding=15)
        log_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(log_frame, text="Activity Log", style='Section.TLabel').pack()
        
        # Text widget with scrollbar for log
        log_text_frame = ttk.Frame(log_frame, style='Section.TFrame')
        log_text_frame.pack(fill='both', expand=True, pady=5)
        
        self.log_text = tk.Text(log_text_frame, height=12, bg='#2A2A2A', fg=self.colors['accent_yellow'],
                               font=('Consolas', 10), wrap='word')
        log_scrollbar = ttk.Scrollbar(log_text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scrollbar.pack(side='right', fill='y')
        
        # Footer
        footer_frame = ttk.Frame(scrollable_frame, style='Footer.TFrame', padding=10)
        footer_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(footer_frame, text="Developer: Ben F. Benjamaa | Manager: Lauren B. Trapani | HD Supply Chain Excellence",
                 style='Info.TLabel').pack()
        
        # Pack canvas and scrollbar
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def load_template_config(self):
        """Load template configuration from file."""
        try:
            if SPPAutomationEnhanced:
                # Create a temporary instance to load config
                temp_automation = SPPAutomationEnhanced()
                self.template_config = temp_automation.template_config
                
                # Update GUI with loaded config
                self.template_path_var.set(self.template_config.get('template_path', ''))
                self.use_template_var.set(self.template_config.get('use_template', False))
                self.output_format_var.set(self.template_config.get('output_format', 'xlsx'))
                
                self.log_message(f"Template configuration loaded: {len(self.template_config)} settings")
                self.on_template_toggle()  # Update UI state
        except Exception as e:
            self.log_message(f"Error loading template config: {e}")
    
    def save_template_config(self):
        """Save current template configuration."""
        try:
            if not self.automation and SPPAutomationEnhanced:
                self.automation = SPPAutomationEnhanced(user_email=self.user_email)
            
            if self.automation:
                self.automation.update_template_config(
                    template_path=self.template_path_var.get(),
                    use_template=self.use_template_var.get(),
                    output_format=self.output_format_var.get()
                )
                
                self.log_message("✓ Template configuration saved successfully")
                messagebox.showinfo("Success", "Template configuration saved!")
            else:
                raise Exception("Automation module not available")
            
        except Exception as e:
            error_msg = f"Error saving template config: {e}"
            self.log_message(f"✗ {error_msg}")
            self.show_error("Configuration Error", error_msg)
    
    def on_template_toggle(self):
        """Handle template checkbox toggle."""
        use_template = self.use_template_var.get()
        
        # Enable/disable template path entry
        state = 'normal' if use_template else 'disabled'
        self.template_entry.config(state=state)
        
        # Update output format based on template usage
        if use_template:
            self.output_format_var.set('xlsm')
            info_text = "Template enabled: Output will use macro-enabled format with consistent formatting"
        else:
            self.output_format_var.set('xlsx')
            info_text = "Template disabled: Output will be standard Excel format"
        
        self.template_info.config(text=info_text)
    
    def browse_template(self):
        """Browse for template file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel Template File",
            filetypes=[
                ("Excel Macro Files", "*.xlsm"),
                ("Excel Files", "*.xlsx"),
                ("All Excel Files", "*.xls*")
            ]
        )
        
        if file_path:
            self.template_path_var.set(file_path)
            self.use_template_var.set(True)
            self.on_template_toggle()
            self.log_message(f"Template file selected: {file_path}")
    
    def authenticate_user(self):
        """Authenticate user with Snowflake."""
        email = self._sanitize_email(self.email_var.get())
        if not email:
            self.show_error("Authentication Error", "Please enter your HD Supply email address")
            return
        
        if not self.validate_email(email):
            self.show_error("Authentication Error", "Please enter a valid email address")
            return
        
        # Keep the sanitized email for downstream use
        self.user_email = email
        self.log_message(f"Authenticating user: {email}")
        self.auth_status.config(text="⏳ Authenticating...", foreground=self.colors['info_blue'])
        self.root.update()
        
        # Run authentication in background
        threading.Thread(target=self._authenticate_thread, daemon=True).start()
    
    def _authenticate_thread(self):
        """Background authentication thread."""
        try:
            if not SPPAutomationEnhanced:
                raise Exception("SPP Automation module not available")
                
            # Use the sanitized email captured during authenticate_user
            if not self.user_email:
                self.user_email = self._sanitize_email(self.email_var.get())
            
            self.log_message(f"Creating automation instance for {self.user_email}")
            self.automation = SPPAutomationEnhanced(user_email=self.user_email)
            
            self.log_message("Attempting Snowflake connection...")
            # Use direct connection instead of test_connection for better error handling
            success = self.automation.connect_to_snowflake()
            
            if success:
                message = f"Successfully authenticated as {self.user_email}"
                self.root.after(0, self._handle_auth_result, True, message)
            else:
                message = ("Failed to authenticate with Snowflake.\n\n"
                          "Please ensure:\n"
                          "1. You completed the browser authentication\n"
                          "2. Your VPN is connected (if required)\n"
                          "3. You have access to the DM_SUPPLYCHAIN database")
                self.root.after(0, self._handle_auth_result, False, message)
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"Authentication error: {error_msg}")
            
            # Provide more specific error messages
            if "timeout" in error_msg.lower():
                message = ("Authentication timed out.\n\n"
                          "Please try again and complete the browser authentication promptly.")
            elif "browser" in error_msg.lower():
                message = ("Failed to complete browser authentication.\n\n"
                          "Please ensure your default browser is accessible.")
            else:
                message = f"Authentication failed:\n{error_msg}"
            
            self.root.after(0, self._handle_auth_result, False, message)
    
    def _handle_auth_result(self, success, message):
        """Handle authentication result."""
        if success:
            self.authenticated = True
            self.auth_status.config(text="✓ Authenticated", foreground=self.colors['success_green'])
            self.log_message(f"✓ Authentication successful: {message}")
            messagebox.showinfo("Authentication Success", message)
        else:
            self.authenticated = False
            self.auth_status.config(text="✗ Failed", foreground=self.colors['error_red'])
            self.log_message(f"✗ Authentication failed: {message}")
            messagebox.showerror("Authentication Failed", message)
    
    def test_connection(self):
        """Test Snowflake connection."""
        if not self.authenticated:
            self.show_error("Authentication Required", "Please authenticate first")
            return
        
        self.log_message("Testing Snowflake connection...")
        threading.Thread(target=self._test_connection_thread, daemon=True).start()
    
    def _test_connection_thread(self):
        """Background connection test thread."""
        try:
            if self.automation:
                success, message = self.automation.test_connection()
                self.root.after(0, self._handle_connection_test_result, success, message)
            else:
                self.root.after(0, self._handle_connection_test_result, False, "No authentication instance available")
        except Exception as e:
            self.root.after(0, self._handle_connection_test_result, False, str(e))
    
    def _handle_connection_test_result(self, success, message):
        """Handle connection test result."""
        if success:
            self.log_message(f"✓ Connection test successful: {message}")
        else:
            self.log_message(f"✗ Connection test failed: {message}")
    
    def test_query(self):
        """Test query execution."""
        if not self.authenticated:
            self.show_error("Authentication Required", "Please authenticate first")
            return
        
        vendor_numbers = [v.strip() for v in self.vendor_var.get().split(',') if v.strip()]
        report_month = self.month_var.get().strip()
        
        if not vendor_numbers or not report_month:
            self.show_error("Input Error", "Please provide vendor numbers and report month")
            return
        
        self.log_message(f"Testing query with vendors: {vendor_numbers}, month: {report_month}")
        threading.Thread(target=self._test_query_thread, args=(vendor_numbers, report_month), daemon=True).start()
    
    def _test_query_thread(self, vendor_numbers, report_month):
        """Background query test thread."""
        try:
            if self.automation:
                success, message, count = self.automation.test_query(vendor_numbers, report_month)
                self.root.after(0, self._handle_query_test_result, success, message, count)
            else:
                self.root.after(0, self._handle_query_test_result, False, "No authentication instance available", 0)
        except Exception as e:
            self.root.after(0, self._handle_query_test_result, False, str(e), 0)
    
    def _handle_query_test_result(self, success, message, count):
        """Handle query test result."""
        if success:
            self.log_message(f"✓ Query test successful: {message}")
        else:
            self.log_message(f"✗ Query test failed: {message}")
    
    def run_automation(self):
        """Run the full automation process."""
        if not self.authenticated:
            self.show_error("Authentication Required", "Please authenticate first")
            return
        
        # Validate inputs
        vendor_numbers = [v.strip() for v in self.vendor_var.get().split(',') if v.strip()]
        report_month = self.month_var.get().strip()
        date_filter = self.date_var.get().strip()
        
        if not vendor_numbers:
            self.show_error("Input Error", "Please provide at least one vendor number")
            return
        
        if not report_month:
            self.show_error("Input Error", "Please provide report month")
            return
        
        if not date_filter:
            self.show_error("Input Error", "Please provide date filter")
            return
        
        # Save template configuration before running
        if self.use_template_var.get():
            try:
                self.save_template_config()
            except Exception as e:
                self.log_message(f"Warning: Could not save template config: {e}")
        
        # Disable run button and start progress
        self.run_btn.config(state='disabled')
        self.progress.start()
        
        self.log_message("=== Starting SPP Automation Process ===")
        self.log_message(f"Vendors: {vendor_numbers}")
        self.log_message(f"Report Month: {report_month}")
        self.log_message(f"Date Filter: {date_filter}")
        self.log_message(f"Use Template: {self.use_template_var.get()}")
        
        # Run automation in background
        threading.Thread(target=self._run_automation_thread, 
                        args=(vendor_numbers, report_month, date_filter), daemon=True).start()
    
    def _run_automation_thread(self, vendor_numbers, report_month, date_filter):
        """Background automation thread."""
        try:
            if not self.automation:
                raise Exception("No authentication instance available")
            
            # Update automation with current template configuration
            self.automation.update_template_config(
                template_path=self.template_path_var.get(),
                use_template=self.use_template_var.get(),
                output_format=self.output_format_var.get()
            )
            
            # Run automation
            output_file, status_message = self.automation.run_full_automation(
                vendor_numbers, report_month, date_filter
            )
            
            self.root.after(0, self._handle_automation_result, True, output_file, status_message)
            
        except Exception as e:
            error_msg = f"Automation failed: {str(e)}"
            self.root.after(0, self._handle_automation_result, False, "", error_msg)
    
    def _handle_automation_result(self, success, output_file, message):
        """Handle automation result."""
        # Stop progress and re-enable button
        self.progress.stop()
        self.run_btn.config(state='normal')
        
        if success and output_file:
            self.log_message("=== Automation Completed Successfully ===")
            self.log_message(f"✓ {message}")
            self.log_message(f"Output file: {output_file}")
            
            # Ask if user wants to open the file
            result = messagebox.askyesno(
                "Success!", 
                f"Report generated successfully!\\n\\n{message}\\n\\nWould you like to open the output file?",
                icon='question'
            )
            
            if result:
                try:
                    os.startfile(output_file)
                except Exception as e:
                    self.log_message(f"Could not open file automatically: {e}")
        else:
            self.log_message("=== Automation Failed ===")
            self.log_message(f"✗ {message}")
            self.show_error("Automation Failed", message)
    
    def open_output_folder(self):
        """Open the output folder."""
        try:
            output_dir = os.path.abspath("Output")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            os.startfile(output_dir)
            self.log_message(f"Opened output folder: {output_dir}")
        except Exception as e:
            error_msg = f"Could not open output folder: {e}"
            self.log_message(f"✗ {error_msg}")
            self.show_error("Folder Error", error_msg)
    
    def log_message(self, message):
        """Add message to log display."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\\n"
        
        self.log_text.insert('end', log_entry)
        self.log_text.see('end')
        self.root.update_idletasks()
        
        # Also log to file
        logger.info(message)
    
    def show_error(self, title, message):
        """Show error dialog and log error."""
        self.log_message(f"ERROR - {title}: {message}")
        messagebox.showerror(title, message)

def main():
    """Main entry point."""
    try:
        root = tk.Tk()
        
        # Check if automation module is available
        if SPPAutomationEnhanced is None:
            messagebox.showerror(
                "Module Error",
                "Could not import SPP automation module. Please ensure spp_automation_enhanced.py is in the same directory."
            )
            return
        
        app = SPPEnhancedGUI(root)
        
        # Handle window closing
        def on_closing():
            if messagebox.askokcancel("Quit", "Do you want to quit?"):
                logger.info("Application closing")
                root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        logger.info("Starting GUI application")
        root.mainloop()
        
    except Exception as e:
        error_msg = f"Fatal error starting application: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        
        try:
            messagebox.showerror("Fatal Error", error_msg)
        except:
            print(f"FATAL ERROR: {error_msg}")

if __name__ == "__main__":
    main()
