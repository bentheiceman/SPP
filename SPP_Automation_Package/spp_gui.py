"""
SPP Metric Automation - GUI Interface
Provides an easy-to-use graphical interface for running SPP metric automation.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from spp_metric_automation import SPPMetricAutomation
import os

class SPPAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SPP Metric Automation Tool")
        self.root.geometry("600x500")
        
        # Initialize automation class
        self.automation = None
        
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="SPP Metric Automation Tool", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Vendor Numbers
        ttk.Label(main_frame, text="Vendor Numbers:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.vendor_entry = ttk.Entry(main_frame, width=50)
        self.vendor_entry.grid(row=1, column=1, pady=5, padx=(10, 0))
        self.vendor_entry.insert(0, "52889")  # Default example
        
        ttk.Label(main_frame, text="(comma-separated)", 
                 font=("Arial", 8)).grid(row=2, column=1, sticky=tk.W, padx=(10, 0))
        
        # Report Month
        ttk.Label(main_frame, text="Report Month:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.month_entry = ttk.Entry(main_frame, width=50)
        self.month_entry.grid(row=3, column=1, pady=5, padx=(10, 0))
        self.month_entry.insert(0, "FY2025-APR")  # Default example
        
        ttk.Label(main_frame, text="(e.g., FY2025-APR)", 
                 font=("Arial", 8)).grid(row=4, column=1, sticky=tk.W, padx=(10, 0))
        
        # Date Filter
        ttk.Label(main_frame, text="Date Filter:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.date_entry = ttk.Entry(main_frame, width=50)
        self.date_entry.grid(row=5, column=1, pady=5, padx=(10, 0))
        self.date_entry.insert(0, "202507")  # Default example
        
        ttk.Label(main_frame, text="(YYYYMM format for ASN query)", 
                 font=("Arial", 8)).grid(row=6, column=1, sticky=tk.W, padx=(10, 0))
        
        # Configuration File
        ttk.Label(main_frame, text="Config File:").grid(row=7, column=0, sticky=tk.W, pady=5)
        config_frame = ttk.Frame(main_frame)
        config_frame.grid(row=7, column=1, pady=5, padx=(10, 0), sticky=(tk.W, tk.E))
        
        self.config_entry = ttk.Entry(config_frame, width=40)
        self.config_entry.grid(row=0, column=0, padx=(0, 5))
        self.config_entry.insert(0, "config.ini")
        
        ttk.Button(config_frame, text="Browse", 
                  command=self.browse_config).grid(row=0, column=1)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=20)
        
        self.run_button = ttk.Button(button_frame, text="Run Automation", 
                                    command=self.run_automation)
        self.run_button.grid(row=0, column=0, padx=5)
        
        ttk.Button(button_frame, text="Edit Config", 
                  command=self.edit_config).grid(row=0, column=1, padx=5)
        
        ttk.Button(button_frame, text="View Output Folder", 
                  command=self.view_output_folder).grid(row=0, column=2, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Status text
        self.status_text = tk.Text(main_frame, height=10, width=70)
        self.status_text.grid(row=10, column=0, columnspan=2, pady=10)
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=10, column=2, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def browse_config(self):
        filename = filedialog.askopenfilename(
            title="Select Configuration File",
            filetypes=[("INI files", "*.ini"), ("All files", "*.*")]
        )
        if filename:
            self.config_entry.delete(0, tk.END)
            self.config_entry.insert(0, filename)
    
    def edit_config(self):
        config_file = self.config_entry.get()
        if os.path.exists(config_file):
            os.startfile(config_file)
        else:
            messagebox.showerror("Error", f"Configuration file not found: {config_file}")
    
    def view_output_folder(self):
        output_dir = r"c:\Users\1015723\Downloads\SPP\Output"
        if os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showinfo("Info", f"Output directory will be created: {output_dir}")
    
    def log_message(self, message):
        """Add message to status text area."""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def run_automation(self):
        """Run the automation in a separate thread."""
        # Validate inputs
        vendors = [v.strip() for v in self.vendor_entry.get().split(',') if v.strip()]
        month = self.month_entry.get().strip()
        date_filter = self.date_entry.get().strip()
        config_file = self.config_entry.get().strip()
        
        if not vendors:
            messagebox.showerror("Error", "Please enter at least one vendor number")
            return
        
        if not month:
            messagebox.showerror("Error", "Please enter a report month")
            return
            
        if not date_filter:
            messagebox.showerror("Error", "Please enter a date filter")
            return
        
        # Disable run button and start progress
        self.run_button.config(state='disabled')
        self.progress.start()
        self.status_text.delete(1.0, tk.END)
        
        # Run automation in thread
        thread = threading.Thread(
            target=self._run_automation_thread,
            args=(vendors, month, date_filter, config_file)
        )
        thread.daemon = True
        thread.start()
    
    def _run_automation_thread(self, vendors, month, date_filter, config_file):
        """Run automation in separate thread."""
        try:
            self.log_message("Starting SPP Metric Automation...")
            self.log_message(f"Vendors: {', '.join(vendors)}")
            self.log_message(f"Report Month: {month}")
            self.log_message(f"Date Filter: {date_filter}")
            
            # Create automation instance
            self.automation = SPPMetricAutomation(config_file)
            
            # Run automation
            output_file = self.automation.run_automation(vendors, month, date_filter)
            
            self.log_message("Automation completed successfully!")
            self.log_message(f"Output file: {output_file}")
            
            # Ask if user wants to open the file
            self.root.after(0, lambda: self._ask_open_file(output_file))
            
        except Exception as e:
            error_msg = f"Automation failed: {str(e)}"
            self.log_message(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
        
        finally:
            # Re-enable button and stop progress
            self.root.after(0, self._automation_finished)
    
    def _ask_open_file(self, output_file):
        """Ask user if they want to open the output file."""
        if messagebox.askyesno("Success", 
                              f"Automation completed successfully!\n\nOutput file: {os.path.basename(output_file)}\n\nWould you like to open the file?"):
            try:
                os.startfile(output_file)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")
    
    def _automation_finished(self):
        """Clean up after automation finishes."""
        self.progress.stop()
        self.run_button.config(state='normal')

def main():
    root = tk.Tk()
    app = SPPAutomationGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
