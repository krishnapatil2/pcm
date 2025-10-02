"""
PCM Application - Refactored Version
Clean, modular, and maintainable code structure
"""
# C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files

import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
from db_manager import setup_database
from ui_components import (
    HomePage, CompactHomePage, MinimalistHomePage, NavigationBar, MonthlyFloatReportPage, 
    NMASSAllocationPage, ObligationSettlementPage, SegregationReportPage
)
from email_config_page import EmailConfigPage
from client_position_page import ClientPositionPage
from report_processors import (
    MonthlyFloatProcessor, NMASSAllocationProcessor, 
    ObligationSettlementProcessor, SegregationReportProcessor, ClientPositionProcessor
)
from utils import ErrorLogger, MessageHandler, WindowManager, Constants


class PCMApplication:
    """Main PCM Application class - Clean and focused"""
    
    def __init__(self, root, db_path=None, home_page_style='compact'):
        self.root = root
        self.db_path = db_path
        self.home_page_style = home_page_style  # 'original', 'compact', 'minimalist'
        
        # Setup window
        self._setup_window()
        
        # Initialize components
        self._init_components()
        
        # Create UI
        self._create_ui()
        
        # Show home page by default
        self.show_page('home')
    
    def _setup_window(self):
        """Setup main window properties"""
        icon_path = self._get_icon_path()
        WindowManager.setup_main_window(self.root, icon_path)
    
    def _get_icon_path(self):
        """Get icon path for the application"""
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, "logo.ico")
        else:
            return os.path.abspath("logo.ico")
    
    def _init_components(self):
        """Initialize application components"""
        # Error logger
        self.error_logger = ErrorLogger()
        
        # Message handler
        self.message_handler = MessageHandler()
        
        # Processors
        self.processors = {
            'monthly_float': MonthlyFloatProcessor(self.db_path, self.error_logger.log_error),
            'nmass_allocation': NMASSAllocationProcessor(self.db_path, self.error_logger.log_error),
            'obligation_settlement': ObligationSettlementProcessor(self.db_path, self.error_logger.log_error),
            'segregation_report': SegregationReportProcessor(self.db_path, self.error_logger.log_error),
            'client_position': ClientPositionProcessor(self.db_path, self.error_logger.log_error)
        }
        
        # Pages dictionary
        self.pages = {}
    
    def _create_ui(self):
        """Create the main UI"""
        # Create main container
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create navigation
        self._create_navigation()
        
        # Create content area with company green theme
        self.content_frame = tk.Frame(self.main_container, bg='#f0f8f0')  # Light green background
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Create pages
        self._create_pages()
    
    def _create_navigation(self):
        """Create navigation bar"""
        self.nav_bar = NavigationBar(
            self.main_container,
            on_home_click=self._on_home_click,
            on_processing_select=self._on_processing_select,
            on_email_config_click=self._on_email_config_click
        )
    
    def _create_pages(self):
        """Create all application pages"""
        # Home page - choose style based on configuration
        if self.home_page_style == 'original':
            self.pages['home'] = HomePage(
                self.content_frame,
                on_feature_click=self._on_feature_click,
                on_info_click=self._on_info_click
            )
        elif self.home_page_style == 'minimalist':
            self.pages['home'] = MinimalistHomePage(
                self.content_frame,
                on_feature_click=self._on_feature_click,
                on_info_click=self._on_info_click
            )
        else:  # default to compact
            self.pages['home'] = CompactHomePage(
                self.content_frame,
                on_feature_click=self._on_feature_click,
                on_info_click=self._on_info_click
            )
        
        # Processing pages
        self._create_processing_pages()
        
        # Email configuration page
        self._create_email_config_page()
        
        # Settings page
    
    def _create_processing_pages(self):
        """Create processing pages with notebook"""
        # Create processing frame
        processing_frame = tk.Frame(self.content_frame, bg=Constants.PROCESSING_BG)
        
        # Header
        header_label = tk.Label(processing_frame, text="Processing", 
                               font=Constants.HEADER_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT)
        header_label.pack(pady=8)
        
        # Style for notebook
        style = ttk.Style()
        style.theme_use('default')
        style.configure('Custom.TNotebook', background=Constants.PROCESSING_BG, borderwidth=0)
        style.configure('Custom.TNotebook.Tab', background=Constants.PROCESSING_BG, 
                       foreground=Constants.PRIMARY_TEXT, padding=[10, 5])
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(processing_frame, style='Custom.TNotebook')
        self.notebook.pack(pady=4, padx=10, fill=tk.BOTH, expand=True)
        
        # Create individual pages
        self._create_monthly_float_page()
        self._create_nmass_allocation_page()
        self._create_obligation_settlement_page()
        self._create_segregation_report_page()
        self._create_client_position_page()
        
        self.pages['processing'] = processing_frame
    
    def _create_email_config_page(self):
        """Create email configuration page"""
        # Create email config page directly
        email_config_page = EmailConfigPage(self.content_frame)
        self.pages['email_config'] = email_config_page
    
    def _create_monthly_float_page(self):
        """Create monthly float report page"""
        page = MonthlyFloatReportPage(
            self.notebook,
            on_process_click=self._process_monthly_float
        )
        self.notebook.add(page.frame, text="Monthly Float Report")
        self.pages['monthly_float'] = page
    
    def _create_nmass_allocation_page(self):
        """Create NMASS allocation page"""
        page = NMASSAllocationPage(
            self.notebook,
            on_generate_click=self._process_nmass_allocation
        )
        self.notebook.add(page.frame, text="NMASS Allocation Report")
        self.pages['nmass_allocation'] = page
    
    def _create_obligation_settlement_page(self):
        """Create obligation settlement page"""
        page = ObligationSettlementPage(
            self.notebook,
            on_generate_click=self._process_obligation_settlement
        )
        self.notebook.add(page.frame, text="Obligation Settlement")
        self.pages['obligation_settlement'] = page
    
    def _create_segregation_report_page(self):
        """Create segregation report page"""
        page = SegregationReportPage(
            self.notebook,
            on_generate_click=self._process_segregation_report
        )
        self.notebook.add(page.frame, text="Segregation Report")
        self.pages['segregation_report'] = page
    
    def _create_client_position_page(self):
        """Create client position page"""
        page = ClientPositionPage(
            self.notebook,
            on_process_click=self._process_client_position
        )
        self.notebook.add(page.frame, text="Client Position Report")
        self.pages['client_position'] = page
    
    # Navigation handlers
    def _on_home_click(self):
        """Handle home button click"""
        self.show_page('home')
    
    def _on_processing_select(self, selection):
        """Handle processing dropdown selection"""
        if selection == "Reports Dashboard":
            self.show_page('processing')
        self.nav_bar.fno_mcx_var.set("Processing")
    
    def _on_email_config_click(self):
        """Handle email configuration button click"""
        self.show_page('email_config')
        # Don't change the processing dropdown text - email config is separate
    
    def _on_feature_click(self, feature_name):
        """Handle feature button click"""
        self.show_page('processing')
        self.root.update_idletasks()
        
        # Switch to appropriate tab
        tab_mapping = {
            "Monthly Float Report": "monthly_float",
            "NMASS Allocation Report": "nmass_allocation", 
            "Obligation Settlement": "obligation_settlement",
            "Segregation Report": "segregation_report",
            "Client Position Report": "client_position"
        }
        
        if feature_name in tab_mapping:
            tab_name = tab_mapping[feature_name]
            for i in range(self.notebook.index("end")):
                if self.notebook.tab(i, "text") == feature_name:
                    self.notebook.select(i)
                    break
    
    def _on_info_click(self, feature_type):
        """Handle info button click"""
        icon_path = self._get_icon_path()
        self.message_handler.show_feature_popup(self.root, feature_type, icon_path)
    
    # Page management
    def show_page(self, page_name):
        """Show the selected page and hide others"""
        for page in self.pages.values():
            if hasattr(page, 'pack_forget'):
                page.pack_forget()
            elif hasattr(page, 'frame'):
                page.frame.pack_forget()
        
        if page_name in self.pages:
            self.pages[page_name].pack(fill=tk.BOTH, expand=True)
    
    # Processing handlers
    def _handle_process(self, processor_key, page_key, success_title, build_success_msg):
        """Generic processor execution and messaging handler to avoid duplication."""
        try:
            values = self.pages[page_key].get_values()
            result = self.processors[processor_key].process(**values)
            if result is None or "error" in str(result).lower() or result == "PERMISSION_ERROR_HANDLED":
                if result == "PERMISSION_ERROR_HANDLED":
                    return
                elif result is None:
                    self.message_handler.show_error("Error", "❌ Processing failed. Please check the logs for details.")
                else:
                    self.message_handler.show_error("Error", f"❌ Failed: {result}")
            else:
                msg = build_success_msg(values, result)
                self.message_handler.show_success(success_title, msg)
        except Exception as e:
            if "file permission error" not in str(e).lower():
                self.message_handler.show_error("Error", f"❌ Failed: {str(e)}")

    def _process_monthly_float(self):
        """Process monthly float report"""
        def _msg(values, result):
            return (
                f"✅ Excel created successfully!\n\n"
                f"📊 FNO Files Processed: {result['fno_count']}\n"
                f"📊 MCX Files Processed: {result['mcx_count']}\n"
                f"ℹ️ Missing Dates Filled: {result['missing_filled']} rows\n"
                f"ℹ️ Monthly Status: Missing dates have been filled automatically. Please check the monthly_status.txt file.\n"
                f"📂 Reconciliation Note: Kindly verify and reconcile the final merged data with:\n"
                f"   - merged_fno_mcx_data.xlsx\n"
                f"   - cp_code_separate_sheets.xlsx.\n\n"
                f"   - And process for the Next Step\n"
                f"📁 Output File: {result['output_file']}"
            )
        self._handle_process('monthly_float', 'monthly_float', "Process Complete", _msg)

    def _process_nmass_allocation(self):
        """Process NMASS allocation report"""
        def _msg(values, result):
            return (
                f"✅ NMASS Allocation Report completed successfully!\n\n"
                f"📅 Selected Date: {values['date']}\n"
                f"📄 Selected Sheet: {values['sheet']}\n"
                f"📎 Attachment 1: {os.path.basename(values['input1_path'])}\n"
                f"📎 Attachment 2: {os.path.basename(values['input2_path'])}\n"
                f"📁 Output Folder: {values['output_path']}\n\n"
                f"📊 Processing Results:\n{result}"
            )
        self._handle_process('nmass_allocation', 'nmass_allocation', "Generate NMASS Allocation Report", _msg)

    def _process_obligation_settlement(self):
        """Process obligation settlement"""
        def _msg(values, result):
            return (
                f"✅ Physical Settlement Processing completed successfully!\n\n"
                f"📁 Output Folder: {values['output_path']}\n"
                f"💾 Backup stored in database.\n\n"
                f"📊 Processing Results:\n{result}"
            )
        self._handle_process('obligation_settlement', 'obligation_settlement', "Success", _msg)

    def _process_segregation_report(self):
        """Process segregation report"""
        def _msg(values, result):
            return (
                f"✅ Segregation Report completed successfully!\n\n"
                f"📅 Selected Date: {values['date']}\n"
                f"🆔 CP PAN: {values['cp_pan']}\n"
                f"📁 Output Folder: {values['output_path']}\n\n"
                f"📊 Processing Results:\n{result}"
            )
        self._handle_process('segregation_report', 'segregation_report', "Generate Segregation Report", _msg)

    def _process_client_position(self):
        """Process client position report"""
        def _msg(values, result):
            selected_cp_info = ""
            if values.get('selected_cp_codes'):
                cp_count = len(values['selected_cp_codes'])
                cp_list = ', '.join(values['selected_cp_codes'][:5])  # Show first 5
                if cp_count > 5:
                    cp_list += f" ... (+{cp_count - 5} more)"
                selected_cp_info = f"✓ Selected CP Codes ({cp_count}): {cp_list}\n"
            else:
                selected_cp_info = "✓ Processed ALL CP codes from the file\n"
            
            return (
                f"✅ Client Position Report completed successfully!\n\n"
                f"📄 Input File: {os.path.basename(values['client_position_path'])}\n"
                f"{selected_cp_info}"
                f"📁 Output Folder: {values['output_path']}\n\n"
                f"📊 Processing Results:\n{result}\n\n"
                f"💡 Tip: Individual encrypted files created for each CP code (no totals by default)"
            )
        self._handle_process('client_position', 'client_position', "Process Client Position", _msg)


def main():
    """Main entry point"""
    # Create splash screen FIRST - before any heavy operations
    splash = tk.Tk()
    splash.title("PCM")
    splash.geometry("400x300")
    splash.resizable(False, False)
    
    # Remove window decorations for cleaner look
    splash.overrideredirect(True)
    
    # Center the splash screen
    splash.eval('tk::PlaceWindow . center')
    
    # Splash screen content
    splash.configure(bg='#2e7d32')
    
    # Logo/Title
    title_label = tk.Label(splash, text="PCM", font=('Segoe UI', 32, 'bold'), 
                          bg='#2e7d32', fg='white')
    title_label.pack(pady=50)
    
    subtitle_label = tk.Label(splash, text="Professional Clearing Member", 
                             font=('Segoe UI', 12), bg='#2e7d32', fg='#c8e6c9')
    subtitle_label.pack(pady=10)
    
    # Loading text
    loading_label = tk.Label(splash, text="Loading...", font=('Segoe UI', 10), 
                            bg='#2e7d32', fg='white')
    loading_label.pack(pady=20)
    
    # Force splash screen to show immediately
    splash.update_idletasks()
    splash.update()
    
    # Initialize main application after splash shows
    def open_main_app():
        try:
            # Safely destroy splash screen
            try:
                splash.destroy()
            except:
                pass  # Ignore errors if splash is already destroyed
            
            root = tk.Tk()
            db_path = setup_database()
            app = PCMApplication(root, db_path=db_path, home_page_style='compact')
            root.mainloop()
        except Exception as e:
            # Safely destroy splash screen on error
            try:
                splash.destroy()
            except:
                pass  # Ignore errors if splash is already destroyed
            
            import tkinter.messagebox as mb
            mb.showerror("Error", f"Failed to start application: {e}")
    
    # Start heavy operations after splash is visible
    splash.after(100, open_main_app)
    splash.mainloop()


if __name__ == "__main__":
    main()

# "Cash placed with NCL"