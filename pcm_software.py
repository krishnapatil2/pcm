import io
import re
import sys
import pandas as pd
import os
import sqlite3
import glob
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, ttk
import traceback
import calendar
from datetime import datetime
import zipfile
import gzip
from tkcalendar import DateEntry
import cons_header
from db_manager import setup_database, insert_report
from physical_settlement_files import build_dict, segregate_excel_by_column


class PCMApplication:
    def __init__(self, root, db_path=None):
        self.root = root
        self.root.title("PCM - Professional Clearing Member")
        # Get screen dimensions for responsive UI
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set window size to be responsive (80% of screen size with minimum dimensions)
        window_width = max(1200, int(screen_width * 0.8))
        window_height = max(800, int(screen_height * 0.8))
        
        # Center the window on screen
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(1000, 700)  # Set minimum size
        
        # Set icon
        if getattr(sys, 'frozen', False):
            self.icon_path = os.path.join(sys._MEIPASS, "logo.ico")
        else:
            self.icon_path = os.path.abspath("logo.ico")

        try:
            self.root.iconbitmap(self.icon_path)
        except:
            pass  # Icon not found, continue without it
        
        # Database path
        self.db_path = db_path

        # Variables
        self.fno_path = tk.StringVar()
        self.mcx_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # Create main container
        self.main_container = tk.Frame(root)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create navigation bar
        self.create_navigation()
        
        # Create content area
        self.content_frame = tk.Frame(self.main_container, bg='#73A070')
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Initialize pages
        self.pages = {}
        self.create_home_page()
        self.create_fno_mcx_page()
        # Check and create database
        # self.setup_database()

        # Show home page by default
        self.show_page('home')
        # Show the startup popup

    def show_feature_popup(self, feature_type):
        popup = tk.Toplevel(self.root)
        try:
            popup.iconbitmap(self.icon_path)
        except:
            pass
        popup.title(f"About {feature_type}")
        popup.geometry("700x400")
        popup.grab_set()
        popup.transient(self.root)

        text = tk.Text(popup, wrap=tk.WORD, font=("Times New Roman", 12), padx=10, pady=10)
        text.pack(expand=True, fill=tk.BOTH)

        text.tag_configure("title", font=("Times New Roman", 14, "bold"), spacing3=10)
        text.tag_configure("bold", font=("Times New Roman", 12, "bold"))
        text.tag_configure("bullet", lmargin1=25, lmargin2=45, spacing3=5)

        if feature_type == "Monthly Float Report":
            text.insert(tk.END, "Average Monthly Float:\n\n", "bold")
            text.insert(tk.END, 
                "This application processes segregation wise report:\n"
                "- Upload NSE and MCX files to generate segregation reports based on the CP Code Excel file.\n"
                "- The output file shows total & average balances for each CP Code, along with a summary.\n"
                "- Missing dates are auto-filled with previous date data for each CP Code.\n"
                "- For reconciliation, CSV files from both folders are merged automatically.\n", 
                "bullet"
            )
        elif feature_type == "NMASS Allocation Report":
            text.insert(tk.END, "Client Allocation and Deallocation in Exchange – Daily Report:\n", "bold")
            text.insert(tk.END, 
                "- Attach both files: 'Client-Level Collaterals' and 'LEDGER'.\n"
                "- Compare the Client Code (CP CODE) values in LEDGER ('F&O Margin') with those in Client-Level Collaterals ('Cash Allocated (b)').\n"
                "- Calculation:\n"
                "   • FO-SEGMENT : TotalCollateral – Cash Allocated (b)\n"
                "   • CD-SEGMENT : TotalCollateral – Cash Allocated (b)\n"
                "- If the balance is positive, mark it as Upward (U).\n"
                "- If the balance is negative, mark it as Downward (D).\n"
                "- If the balance is zero, skip it in the report.\n"
                "- If the client code exists in LEDGER but not in Client-Level Collaterals, mark it as Upward (U) by default.\n"
                "- Final report format:\n"
                "    FO-SEGMENT → date, FO, Member_Code, , cp_code, , C, TotalCollateral,,,,,,, U/D\n"
                "    CD-SEGMENT → date, CD, Member_Code, , cp_code, , C, TotalCollateral,,,,,,, U/D\n", 
                "    All the down (D) records to be reflected first folowed by all the up (U) reords\n", 
                "    As cp_code 90072 corresponds to a TM, this TM code has been excluded.\n",
                "bullet"
            )
        elif feature_type == "Obligation Settlement":
            text.insert(tk.END, "Physical Settlement Report Generation:\n", "bold")
            text.insert(tk.END, 
                "- Upload the Obligation, STT, and Stamp Duty files.\n"
                "- The application will process these files to generate physical settlement reports.\n"
                "- Ensure that the files are in the correct format (CSV or Excel) for successful processing.\n"
                "- The output will be saved in the specified output folder.\n"
                "- Any errors during processing will be logged in an error log file within the output folder.\n",
                "bullet"
            )
        elif feature_type == "Segregation Report":
            text.insert(tk.END, "Segregation Report Generation:\n", "bold")
            text.insert(tk.END, 
                "- Upload all required files: Cash Collateral (CDS & FNO), Daily Margin Reports (NSECR & NSEFNO), CP Master Data (X & F), Collateral Violation Report, and Security Pledge file.\n"
                "- Enter the Date and CP PAN for the report.\n"
                "- The application will process all files and generate a comprehensive segregation report.\n"
                "- Supports both regular CSV and compressed (.gz) files for Security Pledge data.\n"
                "- The output includes all required columns with proper calculations and data mapping.\n"
                "- All input files and the generated report are packaged into a ZIP file for easy storage and backup.\n"
                "- The report follows the standard segregation format with FO and CD segments.\n",
                "bullet"
            )

        text.config(state=tk.DISABLED)
        tk.Button(popup, text="Close", command=popup.destroy).pack(pady=10)

    def log_error(self, output_dir, file_path, error):
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
            print(f"Failed to log error: {e}")

    def create_navigation(self):
        """Modern navigation bar with hover effects"""
        nav_frame = tk.Frame(self.main_container, bg='#2E4C3A', height=60)
        nav_frame.pack(fill=tk.X)
        nav_frame.pack_propagate(False)
        
        # Logo + Title
        title_label = tk.Label(nav_frame, text="PCM", font=('Arial', 20, 'bold'), 
                            bg='#2E4C3A', fg='white')
        title_label.pack(side=tk.LEFT, padx=20, pady=10)
        
        # Nav buttons frame
        nav_buttons_frame = tk.Frame(nav_frame, bg='#2E4C3A')
        nav_buttons_frame.pack(side=tk.RIGHT, padx=20, pady=10)
        
        # Button styling
        def on_enter(e): e.widget.config(bg="#5FA170")
        def on_leave(e): e.widget.config(bg="#73A070")
        
        # Home button
        home_btn = tk.Button(nav_buttons_frame, text="Home", font=('Arial', 12, 'bold'),
                            bg="#73A070", fg='white', relief=tk.FLAT, padx=20,
                            command=lambda: self.show_page('home'))
        home_btn.pack(side=tk.LEFT, padx=5)
        home_btn.bind("<Enter>", on_enter)
        home_btn.bind("<Leave>", on_leave)
        
        # Dropdown
        self.fno_mcx_var = tk.StringVar(value="Processing")
        fno_mcx_menu = tk.OptionMenu(nav_buttons_frame, self.fno_mcx_var,
                                    "Reports Dashboard",
                                    command=self.on_fno_mcx_select)
        fno_mcx_menu.config(font=('Arial', 12, 'bold'), bg="#A3C39E", fg='white', relief=tk.FLAT, padx=20)
        fno_mcx_menu.pack(side=tk.LEFT, padx=5)
        
    def on_fno_mcx_select(self, selection):
        """Handle Processing Files dropdown selection"""
        if selection == "Reports Dashboard":
            self.show_page('fno_mcx')
        self.fno_mcx_var.set("Processing")
    
    def show_page(self, page_name):
        """Show the selected page and hide others"""
        for page in self.pages.values():
            page.pack_forget()
        self.pages[page_name].pack(fill=tk.BOTH, expand=True)
        
    def create_home_page(self):
        """Modern home page with feature cards + functional buttons"""
        home_bg = "#b8efba"  # Company light green
        home_frame = tk.Frame(self.content_frame, bg=home_bg)
        
        # Welcome header
        welcome_label = tk.Label(home_frame, text="Welcome to PCM", 
                                font=('Arial', 24, 'bold'), bg=home_bg, fg='#2E4C3A')
        welcome_label.pack(pady=20)
        
        # Subtitle
        subtitle_label = tk.Label(home_frame, text="Professional Clearing Member", 
                                font=('Arial', 16), bg=home_bg, fg='#2E4C3A')
        subtitle_label.pack(pady=5)
        
        # Features frame
        features_frame = tk.Frame(home_frame, bg=home_bg)
        features_frame.pack(pady=30)

        # Features with working functionality
        features = [
            ("📊 Monthly Float Report", 
            "Merge NSE & MCX files and auto-fill missing dates.",
            lambda: self.show_feature_popup("Monthly Float Report"),
            lambda: self.redirect_to_feature("Monthly Float Report")),
            
            ("🧮 NMASS Allocation Report", 
            "Compare Cash Collateral & Client Allocation Files.",
            lambda: self.show_feature_popup("NMASS Allocation Report"),
            lambda: self.redirect_to_feature("NMASS Allocation Report")),
            
            ("📑 Physical Settlement", 
            "Generate physical settlement reports from obligation files.",
            lambda: self.show_feature_popup("Obligation Settlement"),
            lambda: self.redirect_to_feature("Obligation Settlement")),
            
            ("📋 Segregation Report", 
            "Generate comprehensive segregation reports with all required files.",
            lambda: self.show_feature_popup("Segregation Report"),
            lambda: self.redirect_to_feature("Segregation Report")),
        ]

        # Create cards
        for i, (title, desc, info_cmd, redirect_cmd) in enumerate(features):
            card = tk.Frame(features_frame, bg="white", bd=2, relief="groove")
            card.grid(row=0, column=i, padx=10, pady=10, sticky="n")
            card.configure(highlightbackground="#2E4C3A", highlightthickness=1)
            
            # Title and description
            tk.Label(card, text=title, font=('Arial', 14, 'bold'), bg="white", fg="#2E4C3A").pack(pady=8)
            tk.Label(card, text=desc, font=('Arial', 11), wraplength=200, justify="left", bg="white", fg="#2E4C3A").pack(pady=8)
            
            # Buttons with working functionality
            tk.Button(card, text="ℹ Info", font=('Arial', 11, 'bold'), bg="#73A070", fg="white", command=info_cmd).pack(pady=5)
            tk.Button(card, text="Click Here →", font=('Arial', 11, 'bold'), bg="#2E4C3A", fg="white", command=redirect_cmd).pack(pady=5)

        self.pages['home'] = home_frame

    ###############################################################

    # def create_home_page(self):
    #     """Create the home/welcome page"""
    #     home_bg = '#A3C39E'
    #     home_frame = tk.Frame(self.content_frame, bg=home_bg)
        
    #     # Welcome header
    #     welcome_label = tk.Label(home_frame, text="Welcome to PCM", 
    #                             font=('Arial', 24, 'bold'), bg=home_bg, fg='#2E4C3A')
    #     welcome_label.pack(pady=20)
        
    #     # Subtitle
    #     subtitle_label = tk.Label(home_frame, text="Professional Clearing Member (PCM)", 
    #                              font=('Arial', 16), bg=home_bg, fg='#73A070')
    #     subtitle_label.pack(pady=5)
        
    #     # Features frame
    #     features_frame = tk.Frame(home_frame, bg=home_bg)
    #     features_frame.pack(pady=30)

    #     # Define features with redirect function
    #     features = [
    #         ("📊 Monthly Float Report", "Process NSE & MCX data: merge files, auto-fill missing dates, and separate CP codes.",
    #         lambda: self.show_feature_popup("Monthly Float Report"),
    #         lambda: self.redirect_to_feature("Monthly Float Report")),
            
    #         ("🧮 NMASS Allocation Report", "Cash Collateral File comparison with NMASS Client Allocation File & NSE report generation.",
    #         lambda: self.show_feature_popup("NMASS Allocation Report"),
    #         lambda: self.redirect_to_feature("NMASS Allocation Report")),

    #         ("📑 Physical Settlement Files", "Based on obligation, stamp duty and STT files, generate physical settlement reports.",
    #         lambda: self.show_feature_popup("Obligation Settlement"),
    #         lambda: self.redirect_to_feature("Obligation Settlement")),
            
    #     ]

    #     # Place feature cards side by side using grid
    #     for col, (title, desc, info_cmd, redirect_cmd) in enumerate(features):
    #         feature_box = tk.Frame(features_frame, bg="white", bd=2, relief="groove", width=250, height=250)
    #         feature_box.grid(row=0, column=col, padx=10, pady=10, sticky="n")
    #         # feature_box.grid_propagate(False)  # Fix width & height

    #         # Title
    #         lbl = tk.Label(feature_box, text=title, font=('Arial', 14, 'bold'),
    #                     anchor="w", bg="white", fg='#2E4C3A')
    #         lbl.pack(anchor="w")

    #         # Description
    #         desc_lbl = tk.Label(feature_box, text=desc, font=('Arial', 11),
    #                             anchor="w", wraplength=220, justify="left",
    #                             bg="white", fg='#73A070')
    #         desc_lbl.pack(anchor="w", pady=5)

    #         # Buttons frame
    #         btn_frame = tk.Frame(feature_box, bg="white")
    #         btn_frame.pack(fill=tk.X, pady=5, side=tk.BOTTOM)

    #         info_btn = tk.Button(btn_frame, text="ℹ Info", font=('Arial', 11, 'bold'),
    #                             bg="#73A070", fg="white", relief=tk.RAISED,
    #                             command=info_cmd)
    #         info_btn.pack(side=tk.LEFT, padx=5)

    #         redirect_btn = tk.Button(btn_frame, text="Click Here →", font=('Arial', 11, 'bold'),
    #                                 bg="#2E4C3A", fg="white", relief=tk.RAISED,
    #                                 command=redirect_cmd)
    #         redirect_btn.pack(side=tk.RIGHT, padx=5)

    #     self.pages['home'] = home_frame

    def redirect_to_feature(self, tab_name):
        """Switch to fno_mcx page and then to the correct notebook tab"""
        # Show the notebook page
        self.show_page('fno_mcx')
        # Force update so the notebook is visible before selecting tab
        self.root.update_idletasks()

        # Now switch tab
        for i in range(self.notebook.index("end")):
            if self.notebook.tab(i, "text") == tab_name:
                self.notebook.select(i)
                break

    def create_fno_mcx_page(self):
        """Create the combined NSE/MCX and Ledger processing page"""
        fno_mcx_frame = tk.Frame(self.content_frame, bg='#A3C39E')
        bg_color = "#A3C39E"
        
        # Header
        header_label = tk.Label(fno_mcx_frame, text="Processing", 
                               font=('Arial', 16, 'bold'), bg='#A3C39E', fg='#2c3e50')
        header_label.pack(pady=8)
        # Style for notebook
        style = ttk.Style()
        style.theme_use('default')  # Make sure we can customize

        style.configure('Custom.TNotebook', background=bg_color, borderwidth=0)
        style.configure('Custom.TNotebook.Tab', background=bg_color, foreground='#2c3e50', padding=[10, 5])


        # Create notebook for tabs
        self.notebook = ttk.Notebook(fno_mcx_frame, style='Custom.TNotebook')
        self.notebook.pack(pady=4, padx=10, fill=tk.BOTH, expand=True)
        
        # FNO & MCX Tab
        fno_mcx_tab = tk.Frame(self.notebook, bg=bg_color,bd=2,relief="groove")
        self.notebook.add(fno_mcx_tab, text="Monthly Float Report")
        
        # Common entry style
        entry_width = 45  # Smaller width

        # FNO Folder
        fno_frame = tk.Frame(fno_mcx_tab, bg=bg_color)
        fno_frame.pack(pady=4, padx=10, fill=tk.X)
        
        tk.Label(fno_frame, text="NSE Segregation File:", font=('Arial', 11, 'bold'), 
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        tk.Entry(fno_frame, textvariable=self.fno_path, width=entry_width, font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(fno_frame, text="Browse NSE Segregation File", command=lambda: self.select_folder(self.fno_path),
                 bg='#3498db', fg='white', font=('Arial', 9)).pack(pady=4)
        
        # MCX Folder
        mcx_frame = tk.Frame(fno_mcx_tab, bg=bg_color)
        mcx_frame.pack(pady=13, padx=20, fill=tk.X)
        
        tk.Label(mcx_frame, text="MCX Segregation File:", font=('Arial', 12, 'bold'), 
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        tk.Entry(mcx_frame, textvariable=self.mcx_path, width=entry_width, font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(mcx_frame, text="Browse MCX Segregation File", command=lambda: self.select_folder(self.mcx_path),
                 bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)
        
        # Output Folder
        output_frame = tk.Frame(fno_mcx_tab, bg=bg_color)
        output_frame.pack(pady=13, padx=20, fill=tk.X)
        
        tk.Label(output_frame, text="Output Folder:", font=('Arial', 12, 'bold'), 
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        tk.Entry(output_frame, textvariable=self.output_path, width=entry_width, font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(output_frame, text="Browse Output Folder", command=lambda: self.select_folder(self.output_path),
                 bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)
        
        # Process button
        process_btn = tk.Button(fno_mcx_tab, text="🚀 Process Files", 
                               command=self.process_files, bg='#27ae60', fg='white', 
                               font=('Arial', 14, 'bold'), relief=tk.FLAT, padx=40, pady=8)
        process_btn.pack(pady=28)

        # NMASS Allocation Report Tab
        ledger_tab = tk.Frame(self.notebook, bg=bg_color,
                                bd=2,                # Border width
                                relief="groove",      # Border style: flat, raised, sunken, ridge, solid, groove
                              )
        self.notebook.add(ledger_tab, text="NMASS Allocation Report", padding=10)

        date_frame = tk.Frame(ledger_tab, bg=bg_color)
        date_frame.pack(pady=8, padx=20, fill=tk.X)

        tk.Label(date_frame, text="Date:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(side=tk.LEFT)

        self.date_var = tk.StringVar()

        # Date picker instead of normal Entry
        date_entry = DateEntry(
            date_frame,
            textvariable=self.date_var,
            date_pattern='dd/MM/yyyy',  # DD/MM/YYYY format
            width=15,
            font=('Arial', 10)
        )
        date_entry.pack(side=tk.LEFT, padx=(5, 15))

        
        # Dropdown for sheet selection
        tk.Label(date_frame, text="Sheet:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(side=tk.LEFT)
        
        self.sheet_var = tk.StringVar(value="FNO")
        sheet_options = ["FNO", "CD"]
        sheet_dropdown = tk.OptionMenu(date_frame, self.sheet_var, *sheet_options)
        sheet_dropdown.config(font=('Arial', 10), bg='white', relief=tk.RAISED)
        sheet_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Input Field 1
        input1_frame = tk.Frame(ledger_tab, bg=bg_color)
        input1_frame.pack(pady=8, padx=20, fill=tk.X)
        
        tk.Label(input1_frame, text="NMASS Client Allocation File:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        self.input1_var = tk.StringVar()
        tk.Entry(input1_frame, textvariable=self.input1_var, width=60,
                font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(input1_frame, text="Browse Client Allocation File", command=lambda: self.select_file(self.input1_var),
                bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)
        
        # Input Field 2
        input2_frame = tk.Frame(ledger_tab, bg=bg_color)
        input2_frame.pack(pady=8, padx=20, fill=tk.X)
        
        tk.Label(input2_frame, text="Cash Collateral File:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        self.input2_var = tk.StringVar()
        tk.Entry(input2_frame, textvariable=self.input2_var, width=60,
                font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(input2_frame, text="Browse Cash Collateral File", command=lambda: self.select_file(self.input2_var),
                bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)
        
        # Output Folder
        output_frame = tk.Frame(ledger_tab, bg=bg_color)
        output_frame.pack(pady=8, padx=20, fill=tk.X)
        
        tk.Label(output_frame, text="Output Folder:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        self.ledger_output_path = tk.StringVar()
        tk.Entry(output_frame, textvariable=self.ledger_output_path, width=60,
                font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(output_frame, text="Browse Output Folder", command=lambda: self.select_folder(self.ledger_output_path),
                bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)
        
        # Generate Button
        generate_btn = tk.Button(ledger_tab, text="🚀 Generate NMASS Allocation Report",
                                command=self.generate_ledger, bg='#27ae60', fg='white',
                                font=('Arial', 14, 'bold'), relief=tk.FLAT, padx=40, pady=8)
        generate_btn.pack(pady=28)

        #################################
        # --- Add Data Tab ---
        add_tab = tk.Frame(self.notebook, bg=bg_color, bd=2, relief="groove")
        self.notebook.add(add_tab, text="Obligation Settlement", padding=10)

        tk.Label(add_tab, text="Obligation Physical Settlement", font=('Arial', 14, 'bold'), bg=bg_color, fg='#2c3e50').pack(pady=10)

        self.obligation_path = tk.StringVar()
        self.stt_path = tk.StringVar()
        self.stamp_duty_path = tk.StringVar()
        self.output_path = tk.StringVar()

        form_frame = tk.Frame(add_tab, bg=bg_color)
        form_frame.pack(pady=10, padx=20, fill=tk.X)

        # Helper: add row
        def add_path_row(label_text, var, browse_cmd):
            frame = tk.Frame(form_frame, bg=bg_color)
            frame.pack(pady=6, padx=20, fill=tk.X)

            tk.Label(frame, text=label_text, font=('Arial', 12, 'bold'),
                    bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)

            tk.Entry(frame, textvariable=var, width=60,
                    font=('Arial', 10)).pack(pady=2, fill=tk.X)

            tk.Button(frame, text="Browse", command=lambda: browse_cmd(var),
                    bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=2)
        
        # Build rows
        add_path_row("Obligation File:", self.obligation_path, self.select_file)
        add_path_row("STT File:", self.stt_path, self.select_file)
        add_path_row("Stamp Duty File:", self.stamp_duty_path, self.select_file)
        add_path_row("Output Folder:", self.output_path, self.select_folder)

        tk.Button(add_tab, text="Generate Settlement Report", command=self.physical_settlement_processing, bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'), relief=tk.FLAT, padx=30, pady=6).pack(pady=20)

        #################################
        # --- Segregation Report Tab ---
        segregation_tab = tk.Frame(self.notebook, bg=bg_color, bd=2, relief="groove")
        self.notebook.add(segregation_tab, text="Segregation Report", padding=10)

        # Create a canvas and scrollbar for the segregation tab
        canvas = tk.Canvas(segregation_tab, bg=bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(segregation_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=bg_color)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bind mousewheel to canvas - make it work when hovering over the entire tab
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind mousewheel to the entire segregation tab area
        def bind_scroll_to_tab():
            segregation_tab.bind_all("<MouseWheel>", _on_mousewheel)
        
        def unbind_scroll_from_tab():
            segregation_tab.unbind_all("<MouseWheel>")
        
        # Bind when entering the tab
        segregation_tab.bind("<Enter>", lambda e: bind_scroll_to_tab())
        segregation_tab.bind("<Leave>", lambda e: unbind_scroll_from_tab())
        
        # Also bind to canvas for direct interaction
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # Add keyboard scrolling support
        def _on_key_press(event):
            if event.keysym == "Up":
                canvas.yview_scroll(-1, "units")
            elif event.keysym == "Down":
                canvas.yview_scroll(1, "units")
            elif event.keysym == "Page_Up":
                canvas.yview_scroll(-10, "units")
            elif event.keysym == "Page_Down":
                canvas.yview_scroll(10, "units")
            elif event.keysym == "Home":
                canvas.yview_moveto(0)
            elif event.keysym == "End":
                canvas.yview_moveto(1)
        
        # Bind keyboard events to the canvas
        canvas.bind("<KeyPress>", _on_key_press)
        canvas.focus_set()  # Make canvas focusable for keyboard events

        # Title
        tk.Label(scrollable_frame, text="Segregation Report Generation", font=('Arial', 14, 'bold'), 
                bg=bg_color, fg='#2c3e50').pack(pady=10)
        
        # Add scroll instruction
        instruction_label = tk.Label(scrollable_frame, 
                                   text="💡 Tip: Use mouse wheel, arrow keys, or scrollbar to navigate through all fields", 
                                   font=('Arial', 9, 'italic'), 
                                   bg=bg_color, fg='#666666')
        instruction_label.pack(pady=(0, 10))

        # Date and CP PAN fields
        date_pan_frame = tk.Frame(scrollable_frame, bg=bg_color)
        date_pan_frame.pack(pady=8, padx=20, fill=tk.X)

        # Date field
        tk.Label(date_pan_frame, text="Date:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(side=tk.LEFT)
        self.segregation_date_var = tk.StringVar()
        segregation_date_entry = DateEntry(
            date_pan_frame,
            textvariable=self.segregation_date_var,
            date_pattern='dd/MM/yyyy',
            width=15,
            font=('Arial', 10)
        )
        segregation_date_entry.pack(side=tk.LEFT, padx=(5, 20))

        # CP PAN field
        tk.Label(date_pan_frame, text="CP PAN:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(side=tk.LEFT)
        self.cp_pan_var = tk.StringVar()
        tk.Entry(date_pan_frame, textvariable=self.cp_pan_var, width=20,
                font=('Arial', 10)).pack(side=tk.LEFT, padx=5)

        # File selection frames
        file_frames = [
            ("Cash Collateral Files:", [
                ("CashCollateral_CDS", "self.cash_collateral_cds_var"),
                ("CashCollateral_FNO", "self.cash_collateral_fno_var")
            ]),
            ("Daily Margin Report Files:", [
                ("Daily Margin Report NSECR", "self.daily_margin_nsecr_var"),
                ("Daily Margin Report NSEFNO", "self.daily_margin_nsefno_var")
            ]),
            ("CP Master Data Files:", [
                ("X_CPMaster_data", "self.x_cp_master_var"),
                ("F_CPMaster_data", "self.f_cp_master_var")
            ]),
            ("Collateral Violation Report:", [
                ("Collateral Violation Report", "self.collateral_violation_var")
            ]),
            ("Security Pledge File:", [
                ("F_90123_SEC_PLEDGE", "self.sec_pledge_var")
            ])
        ]

        # Initialize variables
        self.cash_collateral_cds_var = tk.StringVar()
        self.cash_collateral_fno_var = tk.StringVar()
        self.daily_margin_nsecr_var = tk.StringVar()
        self.daily_margin_nsefno_var = tk.StringVar()
        self.x_cp_master_var = tk.StringVar()
        self.f_cp_master_var = tk.StringVar()
        self.collateral_violation_var = tk.StringVar()
        self.sec_pledge_var = tk.StringVar()
        self.segregation_output_var = tk.StringVar()

        # Create file selection sections
        for section_title, files in file_frames:
            section_frame = tk.Frame(scrollable_frame, bg=bg_color)
            section_frame.pack(pady=8, padx=20, fill=tk.X)
            
            tk.Label(section_frame, text=section_title, font=('Arial', 12, 'bold'),
                    bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
            
            for file_name, var_name in files:
                file_frame = tk.Frame(section_frame, bg=bg_color)
                file_frame.pack(pady=4, fill=tk.X)
                
                tk.Label(file_frame, text=f"  {file_name}:", font=('Arial', 10),
                        bg=bg_color, fg='#2c3e50').pack(side=tk.LEFT)
                
                var = getattr(self, var_name.replace('self.', '').replace('_var', '_var'))
                tk.Entry(file_frame, textvariable=var, width=60,
                        font=('Arial', 9)).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                
                tk.Button(file_frame, text="Browse", command=lambda v=var: self.select_file(v),
                        bg='#3498db', fg='white', font=('Arial', 9)).pack(side=tk.LEFT, padx=5)

        # Output folder
        output_frame = tk.Frame(scrollable_frame, bg=bg_color)
        output_frame.pack(pady=8, padx=20, fill=tk.X)
        
        tk.Label(output_frame, text="Output Folder:", font=('Arial', 12, 'bold'),
                bg=bg_color, fg='#2c3e50').pack(anchor=tk.W)
        tk.Entry(output_frame, textvariable=self.segregation_output_var, width=60,
                font=('Arial', 10)).pack(pady=4, fill=tk.X)
        tk.Button(output_frame, text="Browse Output Folder", command=lambda: self.select_folder(self.segregation_output_var),
                bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)

        # Generate Button
        generate_segregation_btn = tk.Button(scrollable_frame, text="🚀 Generate Segregation Report",
                                           command=self.generate_segregation_report, bg='#27ae60', fg='white',
                                           font=('Arial', 14, 'bold'), relief=tk.FLAT, padx=40, pady=8)
        generate_segregation_btn.pack(pady=20)
        
        # Add some bottom padding to ensure last element is visible
        bottom_padding = tk.Frame(scrollable_frame, bg=bg_color, height=20)
        bottom_padding.pack(fill=tk.X)
        
        self.pages['fno_mcx'] = fno_mcx_frame
    
    def physical_settlement_processing(self):
        """Process Obligation, STT, and Stamp Duty files to generate physical settlement reports"""
        try:
            # Read files
            obligation_file = self.obligation_path.get()
            stt_file = self.stt_path.get()
            stamp_duty_file = self.stamp_duty_path.get()
            output_folder = self.output_path.get()
            if not obligation_file or not stt_file or not stamp_duty_file or not output_folder:
                messagebox.showerror("Error", "Please select all input files and output folder.")
                return
            
            # Build STT dictionary
            stt_dict = build_dict(
                file_path=stt_file,
                key_cols=["BrkrOrCtdnPtcptId","TckrSymb", "FinInstrmId"],
                value_cols={
                    "Buy STT": "BuyDelvryTtlTaxs",
                    "Sell STT": "SellDelvryTtlTaxs"
                },
                filter_col="RptHdr",
                filter_val=40
            )

            stamp_duty_dict = build_dict(file_path=stamp_duty_file,
                key_cols=["BrkrOrCtdnPtcptId","TckrSymb", "FinInstrmId"],
                value_cols={
                    # "Sell Stamp Duty": "",
                    "Buy Stamp Duty": "BuyDlvryStmpDty"
                },
                filter_col="RptHdr",
                filter_val=40
            )
            
            output_file = os.path.join(output_folder, "Physical_Settlement_Report.xlsx")

            # Segregate obligation and update Buy/Sell STT from dictionary
            segregate_excel_by_column(
                excel_path=obligation_file,
                output_path=output_file,
                column_name="BrkrOrCtdnPtcptId",
                custom_header=cons_header.OBLIGATION_HEADER,
                update_dicts=[stt_dict, stamp_duty_dict]
            )

            # 🔹 Backup to ZIP
            dt = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"{cons_header.NSE_MEMBER_CODE}_PHYSICAL_SETTLEMENT_{dt}.zip"
            zip_path = os.path.join(output_folder, zip_filename)

            with zipfile.ZipFile(zip_path, 'w') as zipf:
                zipf.write(obligation_file, os.path.basename(obligation_file))
                zipf.write(stt_file, os.path.basename(stt_file))
                zipf.write(stamp_duty_file, os.path.basename(stamp_duty_file))
                zipf.write(output_file, os.path.basename(output_file))

            # 🔹 Insert ZIP into DB
            with open(zip_path, 'rb') as f:
                zip_blob = f.read()
            
            timstamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            insert_report(self.db_path, report_type="PHYSICAL_SETTLEMENT", created_at=timstamp, modified_at=timstamp, report_blob=zip_blob)

            # Success message
            messagebox.showinfo(
                "Success",
                f"✅ Physical Settlement Processing completed successfully!\n\n📁 Output Folder: {output_folder}\n💾 Backup stored in database."
            )

            return "Physical Settlement processed successfully."

        except Exception as e:
            self.log_error(output_folder, "Physical Settlement Processing", e)
            messagebox.showerror("Error", f"❌ Failed: {str(e)}")
            return None

    def show_results_page(self):
        """Show results page (placeholder for now)"""
        messagebox.showinfo("View Results", "Results viewing feature will be implemented in the next step.")

    def select_date(self):
        """Select date using a simple dialog (placeholder for now)"""
        messagebox.showinfo("Date Selection", "Date picker functionality will be implemented in the next step.\n\nFor now, please enter the date manually in DD/MM/YYYY format.")

    def generate_ledger(self):
        """Generate NMASS Allocation Report with actual processing logic"""
        # Get values from the form
        selected_date = self.date_var.get()
        selected_sheet = self.sheet_var.get()
        attachment1 = self.input1_var.get() # Client-Level Collaterals
        attachment2 = self.input2_var.get() # ledger
        output_path = self.ledger_output_path.get()
        
        # Validate inputs
        if selected_date == "DD/MM/YYYY" or not selected_date.strip():
            messagebox.showerror("Error", "Please select a valid date.")
            return
            
        if not attachment1.strip() or not attachment2.strip():
            messagebox.showerror("Error", "Please select both attachment files.")
            return
        
        if not output_path.strip():
            messagebox.showerror("Error", "Please select an output folder for the ledger.")
            return
        
        # Check if files exist
        if not os.path.exists(attachment1):
            messagebox.showerror("Error", f"Attachment 1 file not found:\n{attachment1}")
            return
            
        if not os.path.exists(attachment2):
            messagebox.showerror("Error", f"Attachment 2 file not found:\n{attachment2}")
            return
        
        # Check if output directory exists, if not create it
        if not os.path.exists(output_path):
            try:
                os.makedirs(output_path)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory:\n{str(e)}")
                return
        
        try:
            # Process the files
            result = self.process_ledger_files(attachment1, attachment2, selected_date, selected_sheet, output_path)
            
            if result:
                msg = f"✅ NMASS Allocation Report completed successfully!\n\n" \
                      f"📅 Selected Date: {selected_date}\n" \
                      f"📄 Selected Sheet: {selected_sheet}\n" \
                      f"📎 Attachment 1: {os.path.basename(attachment1)}\n" \
                      f"📎 Attachment 2: {os.path.basename(attachment2)}\n" \
                      f"📁 Output Folder: {output_path}\n\n" \
                      f"📊 Processing Results:\n{result}"
                
                messagebox.showinfo("Generate NMASS Allocation Report", msg)
            else:
                messagebox.showerror("Error", "Failed to process ledger files. Check the error logs.")
                
        except Exception as e:
            self.log_error(output_path, "An error occurred during processing", e)
            messagebox.showerror("Error", f"An error occurred during processing:\n{str(e)}")
    
    def read_file(self, file_path, selected_sheet=None, **kwargs):
        ext = os.path.splitext(file_path)[1].lower()  # get file extension
        
        if ext == ".csv":
            df = pd.read_csv(file_path, **kwargs)
        elif ext in [".xls", ".xlsx"]:
            df = pd.read_excel(file_path, sheet_name=selected_sheet or 0, **kwargs)
        else:
            raise ValueError(f"Unsupported file type: {ext}")
        
        # Drop rows where all columns are NaN
        df = df.dropna(how='all')
        
        return df

    def get_next_file_path(self, output_path, base_name, dt):
        """
        Generate the next available file path by incrementing T000X
        """
        pattern = re.compile(rf"{re.escape(base_name)}_ALLOC_{dt}\.T(\d+)$")

        max_num = 0
        for fname in os.listdir(output_path):
            match = pattern.match(fname)
            if match:
                num = int(match.group(1))
                if num > max_num:
                    max_num = num

        next_num = max_num + 1
        finame = f"{base_name}_ALLOC_{dt}.T{next_num:04d}"
        return os.path.normpath(os.path.join(output_path, finame))

    def build_segment_line(self, date, segment, member_code, cp_code, c_value, margin_value, status):
        return f"{date},{segment},{member_code},,{cp_code},,{c_value},{margin_value},,,,,,,{status}"

    def process_ledger_files(self, file1_path, file2_path, selected_date, selected_sheet, output_path):
        """Process ledger files and perform calculations"""
        try:
            df1 = self.read_file(file1_path)
            # df2 = self.read_file(file2_path, selected_sheet)
            ext = os.path.splitext(file2_path)[1].lower()

            if ext == ".csv":
                df2 = self.read_file(file2_path, header=9, usecols="B:K")
            elif ext in [".xls", ".xlsx"]:
                df2 = self.read_file(file2_path, header=9, usecols=[cons_header.CLIENT_CODE, cons_header.FO_MARGIN])
            else:
                raise ValueError(f"Unsupported file type: {ext}")
            
            df1[cons_header.TM_CP_CODE] = df1[cons_header.TM_CP_CODE].astype(str)
            df2[cons_header.CLIENT_CODE] = df2[cons_header.CLIENT_CODE].astype(str)

            data_dict = {
                        "file_1": dict(zip(df1[cons_header.TM_CP_CODE], df1[cons_header.CASH_COLLECTED])),
                        "file_2": dict(zip(df2[cons_header.CLIENT_CODE], df2[cons_header.FO_MARGIN]))
                        }
            
            
            dict1 = data_dict["file_1"]
            dict2 = data_dict["file_2"]

            formatted_date = datetime.strptime(selected_date, "%d/%m/%Y").strftime("%d-%b-%Y").upper()

            # Static values
            DATE = formatted_date

            dt = datetime.strptime(selected_date, "%d/%m/%Y").strftime("%d%m%Y")
            processed_lines = set()

            # finame = f'{NSE_MEMBER_CODE}_ALLOC_{dt}.T0001'
            # file_path = os.path.normpath(os.path.join(output_path, finame))
            # ✅ dynamically generate file path with incremental T000X
            file_path = self.get_next_file_path(output_path, cons_header.NSE_MEMBER_CODE, dt)

            lines_to_write = []

            for key in dict1:
                if key.lower() == "nan":  # skip string 'nan'
                    continue
                if key in dict2:
                    difference = dict2[key] - dict1[key]
                    
                    if difference > 0:
                        status = "U"
                    elif difference < 0:
                        status = "D"
                    else:
                        continue  # Skip if no change

                    line_fo = self.build_segment_line(DATE, cons_header.SEGMENTS[selected_sheet], cons_header.NSE_MEMBER_CODE, key, cons_header.C, dict2[key], status)

                    # Only write if it's unique
                    if line_fo not in processed_lines:
                        lines_to_write.append(line_fo)
                        processed_lines.add(line_fo)
            
            # Keys in dict2 but NOT in dict1 — always status "U"
            for key in dict2:
                if key.lower() == "nan":  # skip string 'nan'
                    continue
                if key not in dict1:
                    if float(dict2[key]) == 0:
                        continue
                    status = "U"
                    line_fo = self.build_segment_line(DATE, cons_header.SEGMENTS[selected_sheet], cons_header.NSE_MEMBER_CODE, key, cons_header.C, dict2[key], status)
                    if line_fo not in processed_lines:
                        lines_to_write.append(line_fo)
                        processed_lines.add(line_fo)
            
            # Sort so that 'D' comes before 'U'
            sorted_lines = sorted(lines_to_write, key=lambda x: x.strip().split(",")[-1])
            for i in sorted_lines:
                if i.split(",")[4] == '90072':
                    sorted_lines.remove(i)

            # >>> Write lines into report file here <<<
            with open(file_path, "w") as f:
                if lines_to_write:
                    f.write("\n".join(sorted_lines))
                else:
                    f.write("")

            # Create a ZIP including input files and output file
            zip_filename = f"{cons_header.NSE_MEMBER_CODE}_REPORT_{dt}.zip"
            zip_path = os.path.join(output_path, zip_filename)
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                zipf.write(file1_path, os.path.basename(file1_path))
                zipf.write(file2_path, os.path.basename(file2_path))
                zipf.write(file_path, os.path.basename(file_path))

            # Read ZIP as binary and insert into DB
            with open(zip_path, 'rb') as f:
                zip_blob = f.read()

            timstamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            insert_report(self.db_path, report_type=cons_header.LEDGER, created_at=timstamp, modified_at=timstamp, report_blob=zip_blob)
            summary = "Ledger processed."
            return summary

        except Exception as e:
            print(f"Error in process_ledger_files: {e}")
            self.log_error(output_path, "Error in process_ledger_files", e)
            return None

    def filter_data_by_date(self, df, target_date):
        """Filter dataframe by date"""
        try:
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
            print(f"Error filtering by date: {e}")
            return df

    def show_page(self, page_name):
        """Show the specified page"""
        # Hide all pages
        for page in self.pages.values():
            page.pack_forget()
        
        # Show the requested page
        if page_name in self.pages:
            self.pages[page_name].pack(fill=tk.BOTH, expand=True)

    def select_folder(self, var):
        """Select folder using file dialog"""
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            var.set(folder)

    def select_file(self, var):
        """Select file using file dialog"""
        file = filedialog.askopenfilename(title="Select File", 
                                        filetypes=[("All files", "*.*"), 
                                                 ("Excel files", "*.xlsx;*.xls"), 
                                                 ("CSV files", "*.csv"), 
                                                 ("Text files", "*.txt")])
        if file:
            var.set(file)

    def fill_missing_dates(self, df, error_log_path=None):
        """Fill missing dates by duplicating previous day's data."""
        try:
            df[cons_header.DATE] = pd.to_datetime(df[cons_header.DATE], dayfirst=True)
            df["YEAR"] = df[cons_header.DATE].dt.year
            df["MONTH"] = df[cons_header.DATE].dt.month

            filled_data = []
            status_messages = []

            for cp_code, cp_group in df.groupby(cons_header.CP_CODE, dropna=False):
                for (year, month), month_group in cp_group.groupby(["YEAR", "MONTH"]):
                    try:
                        year = int(year)
                        month = int(month)

                        days_in_month = calendar.monthrange(year, month)[1]
                        month_name = calendar.month_name[month]

                        all_days = pd.date_range(start=f"{year}-{month:02d}-01", periods=days_in_month)

                        existing_dates = set(month_group[cons_header.DATE])
                        missing_dates = [d for d in all_days if d not in existing_dates]

                        msg = ""
                        cp_code_display = "blankcpcode" if cp_code == "" or str(cp_code).lower() == "nan" else cp_code
                        if missing_dates:
                            missing_day_nums = ", ".join(str(d.day) for d in missing_dates)
                            msg = f"[INFO] {cp_code_display} → {month_name} {year}: Missing {len(missing_dates)} day(s) → Days: {missing_day_nums}"
                        else:
                            msg = f"[SUCCESS] {cp_code_display} → {month_name} {year}: ✅ All {days_in_month} days present."

                        status_messages.append(msg)
                        filled_month = month_group.copy()

                        for date in missing_dates:
                            prev_data = filled_month[filled_month[cons_header.DATE] < date]
                            if prev_data.empty:
                                continue
                            last_row = prev_data.sort_values(cons_header.DATE).iloc[-1].copy()
                            last_row[cons_header.DATE] = date
                            filled_month = pd.concat([filled_month, pd.DataFrame([last_row])])

                        filled_data.append(filled_month.sort_values(cons_header.DATE))

                    except Exception as e:
                        if error_log_path:
                            self.log_error(error_log_path, f"{cp_code} - {year}-{month:02d}", e)
                        return None, []

            final_df = pd.concat(filled_data, ignore_index=True)
            final_df.drop(columns=["YEAR", "MONTH"], inplace=True)
            return final_df, status_messages

        except Exception as e:
            if error_log_path:
                self.log_error(error_log_path, "fill_missing_dates", e)

    def merge_fno_and_mcx(self, FNO, MCX, out_path, error_log_path):
        """Merge FNO and MCX data"""
        try:
            all_dataframes = []

            try:
                for file in FNO:
                    df = pd.read_csv(file)
                    df['Source'] = 'FNO'
                    all_dataframes.append(df)
            except Exception as e:
                self.log_error(error_log_path, file, e)

            try:
                for file in MCX:
                    df = pd.read_csv(file)
                    df['Source'] = 'MCX'
                    all_dataframes.append(df)
            except Exception as e:
                self.log_error(error_log_path, file, e)

            if not all_dataframes:
                messagebox.showerror("Error", f"No valid CSV files found in FNO or MCX.\nSee log: {error_log_path}")
                return

            merged_df = pd.concat(all_dataframes, ignore_index=True)
            output_file = os.path.join(out_path, "merged_fno_mcx_data.xlsx")
            merged_df.to_excel(output_file, index=False)

        except Exception as e:
            self.log_error(error_log_path, "merge_fno_and_mcx", e)
            messagebox.showerror("Error", f"A fatal error occurred.\nSee log: {error_log_path}")

    def process_files(self):
        """Process FNO and MCX files"""
        FNO = self.fno_path.get()
        MCX = self.mcx_path.get()
        out_path = self.output_path.get()

        if not FNO or not MCX or not out_path:
            messagebox.showerror("Error", "Please select all folders before processing.")
            return
        
        error_log_path = os.path.join(out_path, "pcm_errors.txt")
        fno_files = glob.glob(os.path.join(FNO, "*.csv"))
        mcx_files = glob.glob(os.path.join(MCX, "*.csv"))

        fno_count = len(fno_files)
        mcx_count = len(mcx_files)
        
        self.merge_fno_and_mcx(fno_files, mcx_files, out_path, error_log_path)
        
        df_list = []

        for folder in [fno_files, mcx_files]:
            for file in folder:
                try:
                    temp_df = pd.read_csv(file)
                    temp_df.columns = temp_df.columns.str.strip()
                    df_list.append(temp_df[cons_header.columns_to_keep])
                except Exception as e:
                    self.log_error(error_log_path, file, e)
                    print(f"Error reading {file}: {e}")

        if not df_list:
            messagebox.showerror("Error", "No CSV files found or all failed to load.")
            return

        df = pd.concat(df_list, ignore_index=True)

        # Fill missing dates
        try:
            df_before_fill = len(df)
            df, messages = self.fill_missing_dates(df, error_log_path)
            df_after_fill = len(df)
            missing_filled = df_after_fill - df_before_fill
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill missing dates.\nSee log at:\n{error_log_path}")
            return

        summary_data = []
        df[cons_header.CP_CODE] = df[cons_header.CP_CODE].fillna("").astype(str).replace("nan", "").str.strip()
        cp_groups = list(df.groupby(cons_header.CP_CODE, dropna=False))
        
        for cp_code, group in cp_groups:
            cp_code_display = "blankcpcode" if cp_code == "" else cp_code
            total_val = group[cons_header.FINANCIAL_LEDGER_BALANCE].sum()
            avg_val = group[cons_header.FINANCIAL_LEDGER_BALANCE].mean()
            group[cons_header.CP_CODE] = "" if cp_code == "" else cp_code
            summary_data.append({
                "CP Code": cp_code_display,
                "Total": total_val,
                "Average": avg_val
            })

        # Write Excel file
        output_file = os.path.join(out_path, "cp_code_separate_sheets.xlsx")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            summary_df = pd.DataFrame(summary_data, columns=["CP Code", "Total", "Average"])
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            for cp_code, group in cp_groups:
                sheet_name = "blankcpcode" if cp_code == "" else cp_code[:31]
                group[cons_header.CP_CODE] = "" if cp_code == "" else cp_code
                group.to_excel(writer, sheet_name=sheet_name, index=False)

        # Final message
        msg = (
            f"✅ Excel created successfully!\n\n"
            f"📊 FNO Files Processed: {fno_count}\n"
            f"📊 MCX Files Processed: {mcx_count}\n"
            f"ℹ️ Missing Dates Filled: {missing_filled} rows\n"
            f"ℹ️ Monthly Status: Missing dates have been filled automatically. Please check the monthly_status.txt file.\n"
            f"📂 Reconciliation Note: Kindly verify and reconcile the final merged data with:\n"
            f"   - merged_fno_mcx_data.xlsx\n"
            f"   - cp_code_separate_sheets.xlsx.\n\n"
            f"   - And process for the Next Step\n"
            f"📁 Output File: {output_file}"
        )

        monthly_log_path = os.path.join(out_path, "monthly_status.txt")
        user_friendly_header = "ℹ️ Monthly Status: Missing dates have been filled automatically. Please check the summary below.\n\n"
        full_message = user_friendly_header + "\n".join(messages)

        with open(monthly_log_path, "w", encoding="utf-8") as f:
            f.write(full_message)
        
        # ------------- ZIP all output files -------------
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(out_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, out_path)
                    zipf.write(file_path, arcname)
        zip_blob = zip_buffer.getvalue()

        # ------------- INSERT into database -------------
        timstamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        insert_report(self.db_path, report_type=cons_header.NSE_AND_MCX, created_at=timstamp, modified_at=timstamp, report_blob=zip_blob)
        messagebox.showinfo("Process Complete", msg)

    def generate_segregation_report(self):
        """Generate Segregation Report with all required file processing"""
        try:
            # Get values from the form
            selected_date = self.segregation_date_var.get()
            cp_pan = self.cp_pan_var.get()
            cash_collateral_cds = self.cash_collateral_cds_var.get()
            cash_collateral_fno = self.cash_collateral_fno_var.get()
            daily_margin_nsecr = self.daily_margin_nsecr_var.get()
            daily_margin_nsefno = self.daily_margin_nsefno_var.get()
            x_cp_master = self.x_cp_master_var.get()
            f_cp_master = self.f_cp_master_var.get()
            collateral_violation = self.collateral_violation_var.get()
            sec_pledge = self.sec_pledge_var.get()
            output_path = self.segregation_output_var.get()
            
            # Validate inputs
            if selected_date == "DD/MM/YYYY" or not selected_date.strip():
                messagebox.showerror("Error", "Please select a valid date.")
                return
                
            if not cp_pan.strip():
                messagebox.showerror("Error", "Please enter CP PAN.")
                return
                
            # Check if all required files are selected
            required_files = [
                (cash_collateral_cds, "Cash Collateral CDS"),
                (cash_collateral_fno, "Cash Collateral FNO"),
                (daily_margin_nsecr, "Daily Margin Report NSECR"),
                (daily_margin_nsefno, "Daily Margin Report NSEFNO"),
                (x_cp_master, "X CP Master Data"),
                (f_cp_master, "F CP Master Data"),
                (collateral_violation, "Collateral Violation Report"),
                (sec_pledge, "Security Pledge File")
            ]
            
            missing_files = []
            for file_path, file_name in required_files:
                if not file_path.strip():
                    missing_files.append(file_name)
            
            if missing_files:
                messagebox.showerror("Error", f"Please select the following files:\n" + "\n".join(f"- {name}" for name in missing_files))
                return
            
            if not output_path.strip():
                messagebox.showerror("Error", "Please select an output folder.")
                return
            
            # Check if files exist
            for file_path, file_name in required_files:
                if not os.path.exists(file_path):
                    messagebox.showerror("Error", f"File not found: {file_name}\n{file_path}")
                    return
            
            # Check if output directory exists, if not create it
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                except Exception as e:
                    messagebox.showerror("Error", f"Cannot create output directory:\n{str(e)}")
                    return
            
            # Process the segregation report
            result = self.process_segregation_files(
                selected_date, cp_pan, cash_collateral_cds, cash_collateral_fno,
                daily_margin_nsecr, daily_margin_nsefno, x_cp_master, f_cp_master,
                collateral_violation, sec_pledge, output_path
            )
            
            if result:
                msg = f"✅ Segregation Report completed successfully!\n\n" \
                      f"📅 Selected Date: {selected_date}\n" \
                      f"🆔 CP PAN: {cp_pan}\n" \
                      f"📁 Output Folder: {output_path}\n\n" \
                      f"📊 Processing Results:\n{result}"
                
                messagebox.showinfo("Generate Segregation Report", msg)
            else:
                messagebox.showerror("Error", "Failed to process segregation files. Check the error logs.")
                
        except Exception as e:
            self.log_error(output_path if 'output_path' in locals() else ".", "An error occurred during segregation processing", e)
            messagebox.showerror("Error", f"An error occurred during processing:\n{str(e)}")

    def process_segregation_files(self, date, cp_pan, cash_collateral_cds, cash_collateral_fno,
                                daily_margin_nsecr, daily_margin_nsefno, x_cp_master, f_cp_master,
                                collateral_violation, sec_pledge, output_path):
        """Process all segregation files and generate the final report"""
        try:
            # Import segregation functions
            from segregation import read_file, write_file, calculate_final_effective_value, build_cp_lookup
            from CONSTANT_SEGREGATION import segregation_headers, A, B, C, D, E, F, G, H, I, J, K, L, O, P, AD, AV, AG, AW, AH, AX, BB, BD, BF
            
            # Format date for output
            formatted_date = datetime.strptime(date, "%d/%m/%Y").strftime("%d-%m-%Y")
            
            # Read CP Master files
            df_fo = read_file(f_cp_master)
            df_cd = read_file(x_cp_master)
            
            cp_codes_fo = df_fo["CP Code"].tolist()
            pan_fo = df_fo["PAN Number"].tolist()
            cp_codes_cd = df_cd["CP Code"].tolist()
            pan_cd = df_cd["PAN Number"].tolist()
            
            # Read Cash Collateral files
            df_cash_cds = read_file(cash_collateral_cds, header_row=9, usecols="B:I")
            df_cash_fno = read_file(cash_collateral_fno, header_row=9, usecols="B:I")
            
            cd_collateral_lookup = dict(zip(df_cash_cds["ClientCode"], df_cash_cds["TotalCollateral"]))
            fo_collateral_lookup = dict(zip(df_cash_fno["ClientCode"], df_cash_fno["TotalCollateral"]))
            
            # Read Daily Margin files
            df_margin_cds = read_file(daily_margin_nsecr, header_row=9, usecols="B:T")
            df_margin_fno = read_file(daily_margin_nsefno, header_row=9, usecols="B:T")
            
            cd_daily_margin_lookup = dict(zip(df_margin_cds["ClientCode"], df_margin_cds["Funds"]))
            fo_daily_margin_lookup = dict(zip(df_margin_fno["ClientCode"], df_margin_fno["Funds"]))
            
            # Read Collateral Violation Report
            df_violation = read_file(collateral_violation, header_row=9, usecols="B:H")
            collateral_violation_lookup = {}
            
            for _, row in df_violation.iterrows():
                client_code = row["ClientCode"]
                cash_eq = row["CashEquivalent"]
                non_cash = row["NonCash"]
                
                if client_code in collateral_violation_lookup:
                    collateral_violation_lookup[client_code]["CashEquivalent"] = cash_eq
                    collateral_violation_lookup[client_code]["NonCash"] = non_cash
                else:
                    collateral_violation_lookup[client_code] = {
                        "CashEquivalent": cash_eq,
                        "NonCash": non_cash
                    }
            
            # Step 1: Read raw file without header
            header_row = None

            # Step 1: Scan file for "GSEC" in first column
            with open(sec_pledge, "r", encoding="utf-8", errors="ignore") as f:
                for idx, line in enumerate(f):
                    first_col = line.split(",")[0].strip()  # only check first column
                    if first_col.upper() == "GSEC":
                        print(f"✅ Found 'GSEC' at line {idx}")
                        header_row = idx + 1  # header is next line
                        break

            if header_row is None:
                raise ValueError("'GSEC' not found in first column of file!")

            # Step 2: Read CSV using the detected header row
            df8 = pd.read_csv(sec_pledge, header=header_row, engine="python")

            # Step 5: Automatically strip column names
            df8.columns = df8.columns.str.strip()

            _sec_pledge_lookup = {}

            for _, row in df8.iterrows():
                client_code = row['Client/CP code']
                isin = row['ISIN']
                gross_value = row['GROSS VALUE']
                haircut = row['HAIRCUT']

                if not client_code or not isin:
                    continue  # skip incomplete rows

                key = f"{client_code}-{isin}"

                _sec_pledge_lookup[key] = {
                    "GROSS VALUE": gross_value,
                    "HAIRCUT": haircut
                }
            sec_pledge_cp_lookup = build_cp_lookup(_sec_pledge_lookup)

            # Generate report data
            data = []
            account_type = "C"
            
            # Process FO data
            for cp, pan_no in zip(cp_codes_fo, pan_fo):
                cv_lookup = collateral_violation_lookup.get(cp, {"CashEquivalent": 0, "NonCash": 0})
                row = {
                    A: formatted_date,
                    B: cp_pan,
                    C: cp_pan,
                    D: cp,
                    E: pan_no,
                    F: "",  # Client PAN
                    G: account_type,
                    H: "FO",
                    I: "",  # UCC Code
                    J: fo_collateral_lookup.get(cp, 0),
                    K: fo_daily_margin_lookup.get(cp, 0),
                    L: fo_daily_margin_lookup.get(cp, 0),
                    O: cv_lookup["CashEquivalent"],
                    P: cv_lookup["NonCash"],
                    BB: sec_pledge_cp_lookup.get(cp, 0),
                    BD: sec_pledge_cp_lookup.get(cp, 0),
                    BF: sec_pledge_cp_lookup.get(cp, 0)
                }
                
                # Duplicate values in other columns
                row[AD] = row[K]
                row[AV] = row[K]
                row[AG] = row[O]
                row[AW] = row[O]
                row[AH] = row[P]
                row[AX] = row[P]
                
                data.append(row)
            
            # Process CD data
            for cp, pan_no in zip(cp_codes_cd, pan_cd):
                cv_lookup = collateral_violation_lookup.get(cp, {"CashEquivalent": 0, "NonCash": 0})
                row = {
                    A: formatted_date,
                    B: cp_pan,
                    C: cp_pan,
                    D: cp,
                    E: pan_no,
                    F: "",  # Client PAN
                    G: account_type,
                    H: "CD",
                    I: "",  # UCC Code
                    J: cd_collateral_lookup.get(cp, 0),
                    K: cd_daily_margin_lookup.get(cp, 0),
                    L: cd_daily_margin_lookup.get(cp, 0),
                    O: cv_lookup["CashEquivalent"],
                    P: cv_lookup["NonCash"],
                }
                
                # Duplicate values in other columns
                row[AD] = row[K]
                row[AV] = row[K]
                row[AG] = row[O]
                row[AW] = row[O]
                row[AH] = row[P]
                row[AX] = row[P]
                
                data.append(row)
            
            # Write output file
            output_file = os.path.join(output_path, f"segregation_report_{formatted_date.replace('-', '')}.xlsx")
            write_file(output_file, data=data, header=segregation_headers)
            
            # Create ZIP file with all input files and output
            dt = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"SEGREGATION_REPORT_{dt}.zip"
            zip_path = os.path.join(output_path, zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                # Add input files
                input_files = [
                    (cash_collateral_cds, "CashCollateral_CDS.xls"),
                    (cash_collateral_fno, "CashCollateral_FNO.xls"),
                    (daily_margin_nsecr, "Daily_Margin_Report_NSECR.xls"),
                    (daily_margin_nsefno, "Daily_Margin_Report_NSEFNO.xls"),
                    (x_cp_master, "X_CPMaster_data.xlsx"),
                    (f_cp_master, "F_CPMaster_data.xlsx"),
                    (collateral_violation, "Collateral_Violation_Report.xls"),
                    (sec_pledge, "F_90123_SEC_PLEDGE_09092025_02.csv.gz")
                ]
                
                for file_path, arcname in input_files:
                    zipf.write(file_path, arcname)
                
                # Add output file
                zipf.write(output_file, os.path.basename(output_file))
            
            # Insert into database
            with open(zip_path, 'rb') as f:
                zip_blob = f.read()
            
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            insert_report(self.db_path, report_type="SEGREGATION_REPORT", created_at=timestamp, modified_at=timestamp, report_blob=zip_blob)
            
            return f"Segregation report generated successfully with {len(data)} records."
            
        except Exception as e:
            print(f"Error in process_segregation_files: {e}")
            self.log_error(output_path, "Error in process_segregation_files", e)
            return None

# Main application
if __name__ == "__main__":
    root = tk.Tk()
    db_path = setup_database()
    app = PCMApplication(root, db_path=db_path)
    root.mainloop()

# pyinstaller --onefile --windowed ^
# --icon="C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\dovminiproj\logo.ico" ^
# --add-data "C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\dovminiproj\logo.ico;." ^
# --hidden-import=win32com ^
# --hidden-import=win32com.client ^
# pcm_software.py

# rmdir /s /q build
# rmdir /s /q dist
# del *.spec
# ie4uinit.exe -ClearIconCache
# taskkill /IM explorer.exe /F
# start explorer.exe
