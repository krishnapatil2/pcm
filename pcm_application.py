"""
PCM Application - Refactored Version
Clean, modular, and maintainable code structure
"""

import sys
import os
import tkinter as tk
from tkinter import ttk
from db_manager import setup_database
from ui_components import (
    HomePage, CompactHomePage, MinimalistHomePage, NavigationBar, MonthlyFloatReportPage, 
    NMASSAllocationPage, ObligationSettlementPage, SegregationReportPage
)
from report_processors import (
    MonthlyFloatProcessor, NMASSAllocationProcessor, 
    ObligationSettlementProcessor, SegregationReportProcessor
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
            'segregation_report': SegregationReportProcessor(self.db_path, self.error_logger.log_error)
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
            on_processing_select=self._on_processing_select
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
        
        self.pages['processing'] = processing_frame
    
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
    
    # Navigation handlers
    def _on_home_click(self):
        """Handle home button click"""
        self.show_page('home')
    
    def _on_processing_select(self, selection):
        """Handle processing dropdown selection"""
        if selection == "Reports Dashboard":
            self.show_page('processing')
        self.nav_bar.fno_mcx_var.set("Processing")
    
    def _on_feature_click(self, feature_name):
        """Handle feature button click"""
        self.show_page('processing')
        self.root.update_idletasks()
        
        # Switch to appropriate tab
        tab_mapping = {
            "Monthly Float Report": "monthly_float",
            "NMASS Allocation Report": "nmass_allocation", 
            "Obligation Settlement": "obligation_settlement",
            "Segregation Report": "segregation_report"
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
    def _process_monthly_float(self):
        """Process monthly float report"""
        try:
            values = self.pages['monthly_float'].get_values()
            result = self.processors['monthly_float'].process(**values)
            
            # Check if result indicates success or failure
            if result is None or "error" in str(result).lower() or result == "PERMISSION_ERROR_HANDLED":
                # Result indicates failure
                if result == "PERMISSION_ERROR_HANDLED":
                    # Permission error popup was already shown, do nothing
                    pass
                elif result is None:
                    # Generic failure - show error message
                    self.message_handler.show_error("Error", "❌ Processing failed. Please check the logs for details.")
                else:
                    # Other error - show the error message
                    self.message_handler.show_error("Error", f"❌ Failed: {result}")
            else:
                # Result indicates success
                msg = (
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
                
                self.message_handler.show_success("Process Complete", msg)
            
        except Exception as e:
            # Don't show error if it was a permission error (already shown as popup)
            if "file permission error" not in str(e).lower():
                self.message_handler.show_error("Error", f"❌ Failed: {str(e)}")
    
    def _process_nmass_allocation(self):
        """Process NMASS allocation report"""
        try:
            values = self.pages['nmass_allocation'].get_values()
            result = self.processors['nmass_allocation'].process(**values)
            
            # Check if result indicates success or failure
            if result is None or "error" in str(result).lower() or result == "PERMISSION_ERROR_HANDLED":
                # Result indicates failure
                if result == "PERMISSION_ERROR_HANDLED":
                    # Permission error popup was already shown, do nothing
                    pass
                elif result is None:
                    # Generic failure - show error message
                    self.message_handler.show_error("Error", "❌ Processing failed. Please check the logs for details.")
                else:
                    # Other error - show the error message
                    self.message_handler.show_error("Error", f"❌ Failed: {result}")
            else:
                # Result indicates success
                msg = f"✅ NMASS Allocation Report completed successfully!\n\n" \
                      f"📅 Selected Date: {values['date']}\n" \
                      f"📄 Selected Sheet: {values['sheet']}\n" \
                      f"📎 Attachment 1: {os.path.basename(values['input1'])}\n" \
                      f"📎 Attachment 2: {os.path.basename(values['input2'])}\n" \
                      f"📁 Output Folder: {values['output_path']}\n\n" \
                      f"📊 Processing Results:\n{result}"
                
                self.message_handler.show_success("Generate NMASS Allocation Report", msg)
            
        except Exception as e:
            # Don't show error if it was a permission error (already shown as popup)
            if "file permission error" not in str(e).lower():
                self.message_handler.show_error("Error", f"❌ Failed: {str(e)}")
    
    def _process_obligation_settlement(self):
        """Process obligation settlement"""
        try:
            values = self.pages['obligation_settlement'].get_values()
            result = self.processors['obligation_settlement'].process(**values)
            
            # Check if result indicates success or failure
            if result is None or "error" in str(result).lower() or result == "PERMISSION_ERROR_HANDLED":
                # Result indicates failure
                if result == "PERMISSION_ERROR_HANDLED":
                    # Permission error popup was already shown, do nothing
                    pass
                elif result is None:
                    # Generic failure - show error message
                    self.message_handler.show_error("Error", "❌ Processing failed. Please check the logs for details.")
                else:
                    # Other error - show the error message
                    self.message_handler.show_error("Error", f"❌ Failed: {result}")
            else:
                # Result indicates success
                msg = f"✅ Physical Settlement Processing completed successfully!\n\n" \
                      f"📁 Output Folder: {values['output_path']}\n" \
                      f"💾 Backup stored in database.\n\n" \
                      f"📊 Processing Results:\n{result}"
                
                self.message_handler.show_success("Success", msg)
            
        except Exception as e:
            # Don't show error if it was a permission error (already shown as popup)
            if "file permission error" not in str(e).lower():
                self.message_handler.show_error("Error", f"❌ Failed: {str(e)}")
    
    def _process_segregation_report(self):
        """Process segregation report"""
        try:
            values = self.pages['segregation_report'].get_values()
            result = self.processors['segregation_report'].process(**values)
            
            # Check if result indicates success or failure
            if result is None or "error" in str(result).lower() or result == "PERMISSION_ERROR_HANDLED":
                # Result indicates failure
                if result == "PERMISSION_ERROR_HANDLED":
                    # Permission error popup was already shown, do nothing
                    pass
                elif result is None:
                    # Generic failure - show error message
                    self.message_handler.show_error("Error", "❌ Processing failed. Please check the logs for details.")
                else:
                    # Other error - show the error message
                    self.message_handler.show_error("Error", f"❌ Failed: {result}")
            else:
                # Result indicates success
                msg = f"✅ Segregation Report completed successfully!\n\n" \
                      f"📅 Selected Date: {values['date']}\n" \
                      f"🆔 CP PAN: {values['cp_pan']}\n" \
                      f"📁 Output Folder: {values['output_path']}\n\n" \
                      f"📊 Processing Results:\n{result}"
                
                self.message_handler.show_success("Generate Segregation Report", msg)
            
        except Exception as e:
            # Don't show error if it was a permission error (already shown as popup)
            if "file permission error" not in str(e).lower():
                self.message_handler.show_error("Error", f"❌ Failed: {str(e)}")


def main():
    """Main entry point"""
    root = tk.Tk()
    db_path = setup_database()
    
    # You can change the home page style here:
    # 'original' - Grid layout with cards (takes more screen space)
    # 'compact' - Horizontal buttons (balanced)
    # 'minimalist' - Simple list (takes minimal screen space)
    app = PCMApplication(root, db_path=db_path, home_page_style='compact')
    root.mainloop()


if __name__ == "__main__":
    main()

# "Cash placed with NCL"