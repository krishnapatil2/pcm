"""
Utility functions for PCM Application
Common utilities and helper functions
"""

import os
import traceback
import tkinter as tk
import math
import threading
import time
from tkinter import messagebox, ttk


class LoadingSpinner(tk.Toplevel):
    def __init__(self, parent, text="Loading..."):
        super().__init__(parent)
        self.title("Please wait")
        self.geometry("250x200")
        self.resizable(False, False)
        self.configure(bg="white")
        self.transient(parent)
        self.grab_set()

        # Center window
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 125
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 100
        self.geometry(f"+{x}+{y}")

        # Main frame
        main_frame = tk.Frame(self, bg="white")
        main_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # Canvas for spinner
        self.canvas = tk.Canvas(main_frame, width=200, height=100, bg="white", highlightthickness=0)
        self.canvas.pack(pady=20)

        # Animation variables
        self.current_dot = 0
        self.dots = []
        self.is_running = True

        # Create 8 dots in a circle
        self.create_dots()
        
        # Start animation using after() method - this is the most reliable way
        self.animate()

    def create_dots(self):
        """Create 8 dots positioned in a circle"""
        center_x, center_y = 100, 50
        radius = 35
        
        for i in range(8):
            angle = (i * 45) * math.pi / 180  # 45 degrees apart
            x = center_x + radius * math.cos(angle)
            y = center_y + radius * math.sin(angle)
            
            # Create dot with light blue color initially
            dot = self.canvas.create_oval(x-4, y-4, x+4, y+4, fill="#ccccff", outline="")
            self.dots.append(dot)

    def animate(self):
        """Animate using after() method - most reliable for Tkinter"""
        if not self.is_running:
            return
            
        try:
            # Reset all dots to light blue
            for dot in self.dots:
                self.canvas.itemconfig(dot, fill="#ccccff")
            
            # Highlight current dot in bright blue
            if self.dots and self.current_dot < len(self.dots):
                self.canvas.itemconfig(self.dots[self.current_dot], fill="#0066ff")
            
            # Move to next dot
            self.current_dot = (self.current_dot + 1) % len(self.dots)
            
            # Force update to ensure animation is visible
            self.canvas.update_idletasks()
            
            # Schedule next animation frame - this is the key!
            self.after(80, self.animate)  # Faster for better visibility
            
        except Exception as e:
            # If any error, stop animation
            self.is_running = False

    def close(self):
        """Close the spinner"""
        self.is_running = False
        try:
            self.destroy()
        except:
            pass

class ErrorLogger:
    """Error logging utility"""
    
    @staticmethod
    def log_error(output_dir, file_path, error):
        """Log errors to a file inside the output folder"""
        # Ensure the folder exists
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        error_log_path = os.path.join(output_dir, "error_log.txt")

        try:
            with open(error_log_path, "a", encoding="utf-8") as f:
                f.write(f"[ERROR] File: {file_path}\n")
                f.write(f"Exception: {str(error)}\n")
                f.write("Traceback:\n")
                f.write(traceback.format_exc())
                f.write("\n" + "="*80 + "\n\n")
        except Exception as e:
            pass


class MessageHandler:
    """Message handling utilities"""
    
    @staticmethod
    def show_feature_popup(parent, feature_type, icon_path=None):
        """Show feature information popup"""
        popup = tk.Toplevel(parent)
        try:
            if icon_path:
                popup.iconbitmap(icon_path)
        except:
            pass
        popup.title(f"About {feature_type}")
        popup.geometry("700x400")
        popup.grab_set()
        popup.transient(parent)

        text = tk.Text(popup, wrap=tk.WORD, font=("Times New Roman", 12), padx=10, pady=10)
        text.pack(expand=True, fill=tk.BOTH)

        text.tag_configure("title", font=("Times New Roman", 14, "bold"), spacing3=10)
        text.tag_configure("bold", font=("Times New Roman", 12, "bold"))
        text.tag_configure("bullet", lmargin1=25, lmargin2=45, spacing3=5)

        feature_descriptions = {
            "Monthly Float Report": {
                "title": "Average Monthly Float:\n\n",
                "description": [
                    "This application processes segregation wise report:",
                    "- Upload NSE and MCX files to generate segregation reports based on the CP Code Excel file.",
                    "- The output file shows total & average balances for each CP Code, along with a summary.",
                    "- Missing dates are auto-filled with previous date data for each CP Code.",
                    "- For reconciliation, CSV files from both folders are merged automatically."
                ]
            },
            "NMASS Allocation Report": {
                "title": "Client Allocation and Deallocation in Exchange – Daily Report:\n",
                "description": [
                    "- Attach both files: 'Client-Level Collaterals' and 'LEDGER'.",
                    "- Compare the Client Code (CP CODE) values in LEDGER ('F&O Margin') with those in Client-Level Collaterals ('Cash Allocated (b)').",
                    "- Calculation:",
                    "   • FO-SEGMENT : TotalCollateral – Cash Allocated (b)",
                    "   • CD-SEGMENT : TotalCollateral – Cash Allocated (b)",
                    "- If the balance is positive, mark it as Upward (U).",
                    "- If the balance is negative, mark it as Downward (D).",
                    "- If the balance is zero, skip it in the report.",
                    "- If the client code exists in LEDGER but not in Client-Level Collaterals, mark it as Upward (U) by default.",
                    "- Final report format:",
                    "    FO-SEGMENT → date, FO, Member_Code, , cp_code, , C, TotalCollateral,,,,,,, U/D",
                    "    CD-SEGMENT → date, CD, Member_Code, , cp_code, , C, TotalCollateral,,,,,,, U/D",
                    "    All the down (D) records to be reflected first folowed by all the up (U) reords",
                    "    As cp_code 90072 corresponds to a TM, this TM code has been excluded."
                ]
            },
            "Obligation Settlement": {
                "title": "Physical Settlement Report Generation:\n",
                "description": [
                    "- Upload the Obligation, STT, and Stamp Duty files.",
                    "- The application will process these files to generate physical settlement reports.",
                    "- Ensure that the files are in the correct format (CSV or Excel) for successful processing.",
                    "- The output will be saved in the specified output folder.",
                    "- Any errors during processing will be logged in an error log file within the output folder."
                ]
            },
            "Segregation Report": {
                "title": "Segregation Report Generation:\n",
                "description": [
                    "- Upload all required files: Cash Collateral (CDS & FNO), Daily Margin Reports (NSECR & NSEFNO), CP Master Data (X & F), Collateral Valuation Report, and Security Pledge file.",
                    "- Enter the Date and CP PAN for the report.",
                    "- The application will process all files and generate a comprehensive segregation report.",
                    "- Supports both regular CSV and compressed (.gz) files for Security Pledge data.",
                    "- The output includes all required columns with proper calculations and data mapping.",
                    "- All input files and the generated report are packaged into a ZIP file for easy storage and backup.",
                    "- The report follows the standard segregation format with FO and CD segments."
                ]
            },
            "File Comparison": {
                "title": "File Comparison & Reconciliation:\n",
                "description": [
                    "- Attach the two files that need to be reconciled.",
                    "- Select whether to compare System against Manual, Manual against System, or both directions.",
                    "- The comparison exports directional sheets with records that exist in one attachment but not the other.",
                    "- A summary tab highlights the number of unmatched records per direction for quick review.",
                    "- Use the output workbook as an audit trail before finalising downstream reports."
                ]
            }
        }

        if feature_type in feature_descriptions:
            desc = feature_descriptions[feature_type]
            text.insert(tk.END, desc["title"], "bold")
            for line in desc["description"]:
                text.insert(tk.END, f"{line}\n", "bullet")

        text.config(state=tk.DISABLED)
        tk.Button(popup, text="Close", command=popup.destroy).pack(pady=10)

    @staticmethod
    def show_success(title, message):
        """Show success message"""
        messagebox.showinfo(title, message)

    @staticmethod
    def show_info(title, message):
        """Show information message"""
        messagebox.showinfo(title, message)

    @staticmethod
    def show_error(title, message):
        """Show error message"""
        messagebox.showerror(title, message)

    @staticmethod
    def show_warning(title, message):
        """Show warning message"""
        messagebox.showwarning(title, message)
    
    @staticmethod
    def show_loading(parent, title="Processing", message="Please wait..."):
        """Show loading dialog with custom spinner"""
        return LoadingSpinner(parent, message)
    
    @staticmethod
    def hide_loading(loading_window):
        """Hide loading dialog"""
        if loading_window and loading_window.winfo_exists():
            loading_window.close()


class FileValidator:
    """File validation utilities"""
    
    @staticmethod
    def validate_file_exists(file_path, file_name):
        """Validate that file exists"""
        if not file_path or not file_path.strip():
            raise ValueError(f"Please select {file_name}.")
        
        if not os.path.exists(file_path):
            raise ValueError(f"{file_name} not found:\n{file_path}")
        
        return True
    
    @staticmethod
    def validate_directory_exists(dir_path, dir_name):
        """Validate that directory exists or can be created"""
        if not dir_path or not dir_path.strip():
            raise ValueError(f"Please select {dir_name}.")
        
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
            except Exception as e:
                raise ValueError(f"Cannot create {dir_name}:\n{str(e)}")
        
        return True
    
    @staticmethod
    def validate_date_format(date_str, field_name="Date"):
        """Validate date format"""
        if not date_str or date_str.strip() == "" or date_str == "DD/MM/YYYY":
            raise ValueError(f"Please select a valid {field_name.lower()}.")
        
        try:
            from datetime import datetime
            datetime.strptime(date_str, "%d/%m/%Y")
            return True
        except ValueError:
            raise ValueError(f"Please enter {field_name.lower()} in DD/MM/YYYY format.")


class WindowManager:
    """Window management utilities"""
    
    @staticmethod
    def setup_main_window(root, icon_path=None):
        """Setup main window properties"""
        root.title("PCM - Professional Clearing Member")
        
        # Get screen dimensions for responsive UI
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # Set window size to be responsive (80% of screen size with minimum dimensions)
        window_width = max(1200, int(screen_width * 0.8))
        window_height = max(800, int(screen_height * 0.8))
        
        # Center the window on screen
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        root.minsize(1000, 700)  # Set minimum size
        try:
            root.state('zoomed')
        except Exception:
            pass
        
        # Set icon
        if icon_path:
            try:
                root.iconbitmap(icon_path)
            except:
                pass  # Icon not found, continue without it


class DataProcessor:
    """Data processing utilities"""
    
    @staticmethod
    def filter_data_by_date(df, target_date):
        """Filter dataframe by date"""
        try:
            import pandas as pd
            # Try to identify date column
            date_columns = [col for col in df.columns if 'date' in col.lower() or 'Date' in col]
            
            if date_columns:
                date_col = date_columns[0]
                # Convert to datetime
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                # Filter by date
                filtered_df = df[df[date_col].dt.date == target_date.date()]
            else:
                # If no date column found, return all data
                filtered_df = df
                
            return filtered_df
            
        except Exception as e:
            return df


class Constants:
    """Application constants"""
    
    # Colors
    NAV_BG = '#2E4C3A'
    CONTENT_BG = '#73A070'
    HOME_BG = '#b8efba'
    PROCESSING_BG = '#A3C39E'
    
    # Button colors
    PRIMARY_BTN = '#27ae60'
    SECONDARY_BTN = '#3498db'
    INFO_BTN = '#73A070'
    DARK_BTN = '#2E4C3A'
    
    # Text colors
    PRIMARY_TEXT = '#2c3e50'
    SECONDARY_TEXT = '#73A070'
    
    # Fonts
    TITLE_FONT = ('Arial', 24, 'bold')
    SUBTITLE_FONT = ('Arial', 16)
    HEADER_FONT = ('Arial', 16, 'bold')
    LABEL_FONT = ('Arial', 12, 'bold')
    BUTTON_FONT = ('Arial', 12, 'bold')
    SMALL_FONT = ('Arial', 10)
    
    # Sizes
    MIN_WINDOW_WIDTH = 1000
    MIN_WINDOW_HEIGHT = 700
    ENTRY_WIDTH = 60
    BUTTON_PADX = 20
    BUTTON_PADY = 8
