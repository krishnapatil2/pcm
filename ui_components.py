"""
UI Components for PCM Application
Separated UI components for better maintainability
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
from  CONSTANT_SEGREGATION import H , D, G

class BasePage:
    """Base class for all pages"""
    def __init__(self, parent, bg_color="#B5D1B1"):
        self.parent = parent
        self.bg_color = bg_color
        self.frame = tk.Frame(parent, bg=bg_color)
    
    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
    
    def pack_forget(self):
        self.frame.pack_forget()


class HomePage(BasePage):
    """Modern and elegant home page with company green theme"""
    def __init__(self, parent, on_feature_click, on_info_click):
        super().__init__(parent, "#dde9dd")  # Light green background
        self.on_feature_click = on_feature_click
        self.on_info_click = on_info_click
        self.create_widgets()
    
    def create_widgets(self):
        # Main container with gradient-like background
        main_container = tk.Frame(self.frame, bg=self.bg_color)
        main_container.pack(expand=True, fill=tk.BOTH, padx=30, pady=25)
        
        # Header section with modern styling
        header_frame = tk.Frame(main_container, bg='#ffffff', relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 25))
        
        # Welcome title with gradient effect
        welcome_label = tk.Label(header_frame, text="PCM Dashboard", 
                                font=('Segoe UI', 24, 'bold'), bg='#ffffff', fg='#1e293b')
        welcome_label.pack(pady=(20, 5))
        
        # Subtitle with better typography
        subtitle_label = tk.Label(header_frame, text="Professional Clearing Member • Report Processing Suite", 
                                font=('Segoe UI', 11), bg='#ffffff', fg='#64748b')
        subtitle_label.pack(pady=(0, 20))
        
        # Modern stats cards
        stats_container = tk.Frame(main_container, bg=self.bg_color)
        stats_container.pack(fill=tk.X, pady=(0, 25))
        
        stats_data = [
            ("📊", "4 Report Types", "Available"),
            ("🚀", "Quick Processing", "Fast & Reliable"),
            ("💾", "Auto Backup", "Database Storage")
        ]
        
        for i, (icon, title, desc) in enumerate(stats_data):
            stat_card = tk.Frame(stats_container, bg='#ffffff', relief=tk.FLAT, bd=0)
            stat_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10) if i < len(stats_data)-1 else 0)
            
            # Icon
            icon_label = tk.Label(stat_card, text=icon, font=('Segoe UI', 16), 
                                bg='#ffffff', fg='#3b82f6')
            icon_label.pack(pady=(15, 5))
            
            # Title
            title_label = tk.Label(stat_card, text=title, font=('Segoe UI', 10, 'bold'), 
                                 bg='#ffffff', fg='#1e293b')
            title_label.pack()
            
            # Description
            desc_label = tk.Label(stat_card, text=desc, font=('Segoe UI', 8), 
                                bg='#ffffff', fg='#64748b')
            desc_label.pack(pady=(2, 15))
        
        # Features in modern grid layout
        features_frame = tk.Frame(main_container, bg=self.bg_color)
        features_frame.pack(expand=True, fill=tk.BOTH)
        
        # Features data with company green theme
        features = [
            ("📊", "Monthly Float Report", "Process NSE & MCX files with intelligent date filling and comprehensive analysis", "Monthly Float Report", "#2d7d32"),  # Dark green
            ("🧮", "NMASS Allocation", "Compare client allocation files and generate detailed allocation reports", "NMASS Allocation Report", "#388e3c"),  # Medium green
            ("📑", "Physical Settlement", "Handle obligation, STT & stamp duty processing with automated workflows", "Obligation Settlement", "#4caf50"),  # Light green
            ("📋", "Segregation Report", "Generate comprehensive segregation reports with all required data sources", "Segregation Report", "#66bb6a"),  # Lighter green
        ]

        # Create modern feature cards in 2x2 grid
        for i, (icon, title, description, feature_name, color) in enumerate(features):
            row = i // 2
            col = i % 2
            
            # Create modern card with shadow effect
            card_frame = tk.Frame(features_frame, bg='#ffffff', relief=tk.FLAT, bd=0)
            card_frame.grid(row=row, column=col, padx=12, pady=8, sticky="nsew", ipadx=15, ipady=20)
            
            # Configure grid weights for responsive layout
            features_frame.grid_rowconfigure(row, weight=1)
            features_frame.grid_columnconfigure(col, weight=1)
            
            # Icon with colored background
            icon_frame = tk.Frame(card_frame, bg=color, relief=tk.FLAT, bd=0)
            icon_frame.pack(pady=(0, 15))
            
            icon_label = tk.Label(icon_frame, text=icon, font=('Segoe UI', 20), 
                                bg=color, fg='white', padx=15, pady=8)
            icon_label.pack()
            
            # Title with better typography
            title_label = tk.Label(card_frame, text=title, 
                                  font=('Segoe UI', 13, 'bold'), bg='#ffffff', fg='#1e293b')
            title_label.pack(pady=(0, 8))
            
            # Description with better formatting
            desc_label = tk.Label(card_frame, text=description, 
                                 font=('Segoe UI', 9), bg='#ffffff', fg='#64748b', 
                                 wraplength=180, justify=tk.CENTER)
            desc_label.pack(pady=(0, 15))
            
            # Modern button frame
            button_frame = tk.Frame(card_frame, bg='#ffffff')
            button_frame.pack(fill=tk.X, padx=5)
            
            # Info button with modern styling
            info_btn = tk.Button(button_frame, text="ℹ Info", font=('Segoe UI', 9, 'bold'), 
                               bg='#e2e8f0', fg='#475569', relief=tk.FLAT, padx=12, pady=6,
                               command=lambda f=feature_name: self.on_info_click(f),
                               cursor='hand2')
            info_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            # Main action button with gradient-like effect
            action_btn = tk.Button(button_frame, text="Open →", font=('Segoe UI', 10, 'bold'), 
                                 bg=color, fg='white', relief=tk.FLAT, padx=20, pady=6,
                                 command=lambda f=feature_name: self.on_feature_click(f),
                                 cursor='hand2')
            action_btn.pack(side=tk.RIGHT)
            
            # Enhanced hover effects
            def on_enter(e, frame=card_frame, btn1=info_btn, btn2=action_btn):
                frame.config(bg='#f8fafc')
                for child in frame.winfo_children():
                    if isinstance(child, tk.Label):
                        child.config(bg='#f8fafc')
                btn1.config(bg='#cbd5e1')
                btn2.config(bg=color)
            
            def on_leave(e, frame=card_frame, btn1=info_btn, btn2=action_btn):
                frame.config(bg='#ffffff')
                for child in frame.winfo_children():
                    if isinstance(child, tk.Label):
                        child.config(bg='#ffffff')
                btn1.config(bg='#e2e8f0')
                btn2.config(bg=color)
            
            card_frame.bind("<Enter>", on_enter)
            card_frame.bind("<Leave>", on_leave)
        
        # Modern footer with green theme
        footer_frame = tk.Frame(main_container, bg='#e8f5e8', relief=tk.FLAT, bd=0)
        footer_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Footer content
        footer_content = tk.Frame(footer_frame, bg='#e8f5e8')
        footer_content.pack(pady=12)
        
        quick_label = tk.Label(footer_content, text="💡 Quick Access: Use the 'Processing' menu above for all reports", 
                              font=('Segoe UI', 9), bg='#e8f5e8', fg='#2e7d32')
        quick_label.pack()


class CompactHomePage(BasePage):
    """Modern compact home page with company green theme"""
    def __init__(self, parent, on_feature_click, on_info_click):
        super().__init__(parent, '#f0f8f0')  # Light green background
        self.on_feature_click = on_feature_click
        self.on_info_click = on_info_click
        self.create_widgets()
    
    def create_widgets(self):
        # Main container with modern styling
        main_container = tk.Frame(self.frame, bg=self.bg_color)
        main_container.pack(expand=True, fill=tk.BOTH, padx=25, pady=20)
        
        # Header with modern design
        header_frame = tk.Frame(main_container, bg='#ffffff', relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title with better typography
        title_label = tk.Label(header_frame, text="PCM Reports", 
                              font=('Segoe UI', 18, 'bold'), bg='#ffffff', fg='#1e293b')
        title_label.pack(pady=(15, 5))
        
        # Subtitle
        subtitle_label = tk.Label(header_frame, text="Professional Clearing Member • Quick Access", 
                                 font=('Segoe UI', 10), bg='#ffffff', fg='#64748b')
        subtitle_label.pack(pady=(0, 15))
        
        # Features as modern horizontal cards
        features_frame = tk.Frame(main_container, bg=self.bg_color)
        features_frame.pack(expand=True, fill=tk.BOTH)
        
        # Features data with company green theme
        features = [
            ("📊", "Monthly Float", "NSE & MCX processing", "Monthly Float Report", "#2d7d32"),  # Dark green
            ("🧮", "NMASS Allocation", "Client allocation reports", "NMASS Allocation Report", "#388e3c"),  # Medium green
            ("📑", "Physical Settlement", "Obligation processing", "Obligation Settlement", "#4caf50"),  # Light green
            ("📋", "Segregation Report", "Comprehensive reports", "Segregation Report", "#66bb6a"),  # Lighter green
        ]

        # Create modern horizontal cards
        for i, (icon, title, desc, feature_name, color) in enumerate(features):
            # Modern card with subtle shadow
            card_frame = tk.Frame(features_frame, bg='#ffffff', relief=tk.FLAT, bd=0)
            card_frame.pack(fill=tk.X, pady=6, padx=5)
            
            # Left side - icon and content
            left_frame = tk.Frame(card_frame, bg='#ffffff')
            left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=12)
            
            # Icon with colored background
            icon_frame = tk.Frame(left_frame, bg=color, relief=tk.FLAT, bd=0)
            icon_frame.pack(side=tk.LEFT, padx=(0, 12))
            
            icon_label = tk.Label(icon_frame, text=icon, font=('Segoe UI', 14), 
                                bg=color, fg='white', padx=8, pady=6)
            icon_label.pack()
            
            # Text content
            text_frame = tk.Frame(left_frame, bg='#ffffff')
            text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            title_label = tk.Label(text_frame, text=title, 
                                  font=('Segoe UI', 12, 'bold'), bg='#ffffff', fg='#1e293b')
            title_label.pack(anchor=tk.W)
            
            desc_label = tk.Label(text_frame, text=desc, 
                                 font=('Segoe UI', 9), bg='#ffffff', fg='#64748b')
            desc_label.pack(anchor=tk.W)
            
            # Right side - buttons
            right_frame = tk.Frame(card_frame, bg='#ffffff')
            right_frame.pack(side=tk.RIGHT, padx=15, pady=12)
            
            # Info button with modern styling
            info_btn = tk.Button(right_frame, text="ℹ", font=('Segoe UI', 9, 'bold'), 
                               bg='#e2e8f0', fg='#475569', relief=tk.FLAT, padx=8, pady=4,
                               command=lambda f=feature_name: self.on_info_click(f),
                               cursor='hand2')
            info_btn.pack(side=tk.LEFT, padx=(0, 6))
            
            # Action button with color
            action_btn = tk.Button(right_frame, text="Open →", font=('Segoe UI', 10, 'bold'), 
                                 bg=color, fg='white', relief=tk.FLAT, padx=16, pady=4,
                                 command=lambda f=feature_name: self.on_feature_click(f),
                                 cursor='hand2')
            action_btn.pack(side=tk.LEFT)
            
            # Enhanced hover effects
            def on_enter(e, frame=card_frame, btn1=info_btn, btn2=action_btn):
                frame.config(bg='#f8fafc')
                for child in frame.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.config(bg='#f8fafc')
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, tk.Label):
                                grandchild.config(bg='#f8fafc')
                btn1.config(bg='#cbd5e1')
                btn2.config(bg=color)
            
            def on_leave(e, frame=card_frame, btn1=info_btn, btn2=action_btn):
                frame.config(bg='#ffffff')
                for child in frame.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.config(bg='#ffffff')
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, tk.Label):
                                grandchild.config(bg='#ffffff')
                btn1.config(bg='#e2e8f0')
                btn2.config(bg=color)
            
            card_frame.bind("<Enter>", on_enter)
            card_frame.bind("<Leave>", on_leave)
        
        # Modern footer with green theme
        footer_frame = tk.Frame(main_container, bg='#e8f5e8', relief=tk.FLAT, bd=0)
        footer_frame.pack(fill=tk.X, pady=(15, 0))
        
        footer_label = tk.Label(footer_frame, text="💡 Use 'Processing' menu above for all reports", 
                               font=('Segoe UI', 9), bg='#e8f5e8', fg='#2e7d32')
        footer_label.pack(pady=10)


class MinimalistHomePage(BasePage):
    """Ultra-minimalist home page with company green theme"""
    def __init__(self, parent, on_feature_click, on_info_click):
        super().__init__(parent, '#f0f8f0')  # Light green background
        self.on_feature_click = on_feature_click
        self.on_info_click = on_info_click
        self.create_widgets()
    
    def create_widgets(self):
        # Main container with modern styling
        main_container = tk.Frame(self.frame, bg=self.bg_color)
        main_container.pack(expand=True, fill=tk.BOTH, padx=20, pady=15)
        
        # Header - compact but modern
        header_frame = tk.Frame(main_container, bg='#ffffff', relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Title with modern typography
        title_label = tk.Label(header_frame, text="PCM", 
                              font=('Segoe UI', 16, 'bold'), bg='#ffffff', fg='#1e293b')
        title_label.pack(pady=(12, 5))
        
        # Subtitle
        subtitle_label = tk.Label(header_frame, text="Professional Clearing Member", 
                                 font=('Segoe UI', 9), bg='#ffffff', fg='#64748b')
        subtitle_label.pack(pady=(0, 12))
        
        # Features as modern list
        features_frame = tk.Frame(main_container, bg=self.bg_color)
        features_frame.pack(expand=True, fill=tk.BOTH)
        
        # Features data with company green theme
        features = [
            ("📊", "Monthly Float Report", "Monthly Float Report", "#2d7d32"),  # Dark green
            ("🧮", "NMASS Allocation Report", "NMASS Allocation Report", "#388e3c"),  # Medium green
            ("📑", "Obligation Settlement", "Obligation Settlement", "#4caf50"),  # Light green
            ("📋", "Segregation Report", "Segregation Report", "#66bb6a"),  # Lighter green
        ]

        # Create modern list layout
        for i, (icon, display_name, feature_name, color) in enumerate(features):
            # Modern row with subtle styling
            row_frame = tk.Frame(features_frame, bg='#ffffff', relief=tk.FLAT, bd=0)
            row_frame.pack(fill=tk.X, pady=2)
            
            # Left side - icon and name
            left_frame = tk.Frame(row_frame, bg='#ffffff')
            left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=12, pady=8)
            
            # Icon with colored background
            icon_frame = tk.Frame(left_frame, bg=color, relief=tk.FLAT, bd=0)
            icon_frame.pack(side=tk.LEFT, padx=(0, 10))
            
            icon_label = tk.Label(icon_frame, text=icon, font=('Segoe UI', 10), 
                                bg=color, fg='white', padx=6, pady=3)
            icon_label.pack()
            
            # Name
            name_label = tk.Label(left_frame, text=display_name, 
                                 font=('Segoe UI', 10, 'bold'), bg='#ffffff', fg='#1e293b')
            name_label.pack(side=tk.LEFT, padx=(0, 10))
            
            # Right side - buttons
            button_frame = tk.Frame(row_frame, bg='#ffffff')
            button_frame.pack(side=tk.RIGHT, padx=12, pady=6)
            
            # Info button - modern styling
            info_btn = tk.Button(button_frame, text="ℹ", font=('Segoe UI', 8, 'bold'), 
                               bg='#e2e8f0', fg='#475569', relief=tk.FLAT, padx=6, pady=2,
                               command=lambda f=feature_name: self.on_info_click(f),
                               cursor='hand2')
            info_btn.pack(side=tk.LEFT, padx=(0, 4))
            
            # Action button - modern styling
            action_btn = tk.Button(button_frame, text="→", font=('Segoe UI', 9, 'bold'), 
                                 bg=color, fg='white', relief=tk.FLAT, padx=10, pady=2,
                                 command=lambda f=feature_name: self.on_feature_click(f),
                                 cursor='hand2')
            action_btn.pack(side=tk.LEFT)
            
            # Enhanced hover effect
            def on_enter(e, frame=row_frame, btn1=info_btn, btn2=action_btn):
                frame.config(bg='#f8fafc')
                for child in frame.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.config(bg='#f8fafc')
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, tk.Label):
                                grandchild.config(bg='#f8fafc')
                btn1.config(bg='#cbd5e1')
                btn2.config(bg=color)
            
            def on_leave(e, frame=row_frame, btn1=info_btn, btn2=action_btn):
                frame.config(bg='#ffffff')
                for child in frame.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.config(bg='#ffffff')
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, tk.Label):
                                grandchild.config(bg='#ffffff')
                btn1.config(bg='#e2e8f0')
                btn2.config(bg=color)
            
            row_frame.bind("<Enter>", on_enter)
            row_frame.bind("<Leave>", on_leave)
        
        # Modern footer with green theme
        footer_frame = tk.Frame(main_container, bg='#e8f5e8', relief=tk.FLAT, bd=0)
        footer_frame.pack(fill=tk.X, pady=(10, 0))
        
        footer_label = tk.Label(footer_frame, text="💡 Use 'Processing' menu above for all reports", 
                               font=('Segoe UI', 8), bg='#e8f5e8', fg='#2e7d32')
        footer_label.pack(pady=8)


class NavigationBar:
    """Modern navigation bar component"""
    def __init__(self, parent, on_home_click, on_processing_select):
        self.parent = parent
        self.on_home_click = on_home_click
        self.on_processing_select = on_processing_select
        self.create_widgets()
    
    def create_widgets(self):
        # Modern navigation bar with company green theme
        nav_frame = tk.Frame(self.parent, bg='#2e7d32', height=65)  # Dark green
        nav_frame.pack(fill=tk.X)
        nav_frame.pack_propagate(False)
        
        # Logo + Title with modern styling
        logo_frame = tk.Frame(nav_frame, bg='#2e7d32')
        logo_frame.pack(side=tk.LEFT, padx=25, pady=15)
        
        title_label = tk.Label(logo_frame, text="PCM", font=('Segoe UI', 22, 'bold'), 
                            bg='#2e7d32', fg='white')
        title_label.pack(side=tk.LEFT)
        
        # Subtitle
        subtitle_label = tk.Label(logo_frame, text="Professional Clearing Member", 
                                font=('Segoe UI', 9), bg='#2e7d32', fg='#c8e6c9')  # Light green text
        subtitle_label.pack(side=tk.LEFT, padx=(10, 0), pady=(5, 0))
        
        # Nav buttons frame
        nav_buttons_frame = tk.Frame(nav_frame, bg='#2e7d32')
        nav_buttons_frame.pack(side=tk.RIGHT, padx=25, pady=15)
        
        # Modern button styling with green theme
        def on_enter(e): 
            e.widget.config(bg="#4caf50")  # Light green hover
        def on_leave(e): 
            e.widget.config(bg="#388e3c")  # Medium green
        
        # Home button with modern styling
        home_btn = tk.Button(nav_buttons_frame, text="🏠 Home", font=('Segoe UI', 11, 'bold'),
                            bg="#388e3c", fg='white', relief=tk.FLAT, padx=18, pady=8,
                            command=self.on_home_click, cursor='hand2')
        home_btn.pack(side=tk.LEFT, padx=(0, 8))
        home_btn.bind("<Enter>", on_enter)
        home_btn.bind("<Leave>", on_leave)
        
        # Modern dropdown with green theme
        self.fno_mcx_var = tk.StringVar(value="Processing")
        fno_mcx_menu = tk.OptionMenu(nav_buttons_frame, self.fno_mcx_var,
                                    "Reports Dashboard",
                                    command=self.on_processing_select)
        fno_mcx_menu.config(font=('Segoe UI', 11, 'bold'), bg="#4caf50", fg='white', 
                           relief=tk.FLAT, padx=18, pady=8, cursor='hand2')
        fno_mcx_menu.pack(side=tk.LEFT, padx=5)


class FileInputWidget:
    """Reusable file input widget"""
    def __init__(self, parent, label_text, var, is_folder=False, entry_width=60):
        self.parent = parent
        self.var = var
        self.is_folder = is_folder
        
        # Main frame
        self.frame = tk.Frame(parent, bg=parent['bg'])
        self.frame.pack(pady=8, padx=20, fill=tk.X)
        
        # Label
        tk.Label(self.frame, text=label_text, font=('Arial', 12, 'bold'),
                bg=parent['bg'], fg='#2c3e50').pack(anchor=tk.W)
        
        # Entry
        tk.Entry(self.frame, textvariable=var, width=entry_width,
                font=('Arial', 10)).pack(pady=4, fill=tk.X)
        
        # Browse button
        button_text = "Browse Folder" if is_folder else "Browse File"
        command = self.select_folder if is_folder else self.select_file
        tk.Button(self.frame, text=button_text, command=command,
                bg='#3498db', fg='white', font=('Arial', 10)).pack(pady=4)
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            self.var.set(folder)
    
    def select_file(self):
        file = filedialog.askopenfilename(title="Select File", 
                                        filetypes=[("All files", "*.*"), 
                                                 ("Excel files", "*.xlsx;*.xls"), 
                                                 ("CSV files", "*.csv"), 
                                                 ("Text files", "*.txt")])
        if file:
            self.var.set(file)


class DateInputWidget:
    """Reusable date input widget"""
    def __init__(self, parent, label_text, var, default_date=None):
        self.parent = parent
        self.var = var
        
        # Main frame
        self.frame = tk.Frame(parent, bg=parent['bg'])
        self.frame.pack(pady=8, padx=20, fill=tk.X)
        
        # Label
        tk.Label(self.frame, text=label_text, font=('Arial', 12, 'bold'),
                bg=parent['bg'], fg='#2c3e50').pack(side=tk.LEFT)
        
        # Date picker
        date_entry = DateEntry(
            self.frame,
            textvariable=var,
            date_pattern='dd/MM/yyyy',
            width=15,
            font=('Arial', 10)
        )
        if default_date:
            date_entry.set_date(default_date)
        date_entry.pack(side=tk.LEFT, padx=(5, 15))


class MonthlyFloatReportPage(BasePage):
    """Monthly Float Report page"""
    def __init__(self, parent, on_process_click):
        super().__init__(parent)
        self.on_process_click = on_process_click
        self.create_widgets()
    
    def create_widgets(self):
        # Header
        header_label = tk.Label(self.frame, text="Monthly Float Report", 
                               font=('Arial', 16, 'bold'), bg=self.bg_color, fg='#2c3e50')
        header_label.pack(pady=8)
        
        # File inputs
        self.fno_path = tk.StringVar()
        self.mcx_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        FileInputWidget(self.frame, "NSE Segregation File:", self.fno_path, is_folder=True)
        FileInputWidget(self.frame, "MCX Segregation File:", self.mcx_path, is_folder=True)
        FileInputWidget(self.frame, "Output Folder:", self.output_path, is_folder=True)
        
        # Process button
        process_btn = tk.Button(self.frame, text="🚀 Process Files", 
                               command=self.on_process_click, bg='#27ae60', fg='white', 
                               font=('Arial', 14, 'bold'), relief=tk.FLAT, padx=40, pady=8)
        process_btn.pack(pady=28)
    
    def get_values(self):
        return {
            'fno_path': self.fno_path.get(),
            'mcx_path': self.mcx_path.get(),
            'output_path': self.output_path.get()
        }


class NMASSAllocationPage(BasePage):
    """NMASS Allocation Report page"""
    def __init__(self, parent, on_generate_click):
        super().__init__(parent)
        self.on_generate_click = on_generate_click
        self.create_widgets()
    
    def create_widgets(self):
        # Header
        header_label = tk.Label(self.frame, text="NMASS Allocation Report", 
                               font=('Arial', 16, 'bold'), bg=self.bg_color, fg='#2c3e50')
        header_label.pack(pady=8)
        
        # Date and sheet selection
        date_sheet_frame = tk.Frame(self.frame, bg=self.bg_color)
        date_sheet_frame.pack(pady=8, padx=20, fill=tk.X)
        
        # Date
        self.date_var = tk.StringVar()
        DateInputWidget(date_sheet_frame, "Date:", self.date_var)
        
        # Sheet dropdown
        tk.Label(date_sheet_frame, text="Sheet:", font=('Arial', 12, 'bold'),
                bg=self.bg_color, fg='#2c3e50').pack(side=tk.LEFT)
        
        self.sheet_var = tk.StringVar(value="FNO")
        sheet_options = ["FNO", "CD"]
        sheet_dropdown = tk.OptionMenu(date_sheet_frame, self.sheet_var, *sheet_options)
        sheet_dropdown.config(font=('Arial', 10), bg='white', relief=tk.RAISED)
        sheet_dropdown.pack(side=tk.LEFT, padx=5)
        
        # File inputs
        self.input1_var = tk.StringVar()
        self.input2_var = tk.StringVar()
        self.output_path = tk.StringVar()
        
        FileInputWidget(self.frame, "NMASS Client Allocation File:", self.input1_var)
        FileInputWidget(self.frame, "Cash Collateral File:", self.input2_var)
        FileInputWidget(self.frame, "Output Folder:", self.output_path, is_folder=True)
        
        # Generate button
        generate_btn = tk.Button(self.frame, text="🚀 Generate NMASS Allocation Report",
                                command=self.on_generate_click, bg='#27ae60', fg='white',
                                font=('Arial', 14, 'bold'), relief=tk.FLAT, padx=40, pady=8)
        generate_btn.pack(pady=28)
    
    def get_values(self):
        return {
            'date': self.date_var.get(),
            'sheet': self.sheet_var.get(),
            'input1': self.input1_var.get(),
            'input2': self.input2_var.get(),
            'output_path': self.output_path.get()
        }


class ObligationSettlementPage(BasePage):
    """Obligation Settlement page"""
    def __init__(self, parent, on_generate_click):
        super().__init__(parent)
        self.on_generate_click = on_generate_click
        self.create_widgets()
    
    def create_widgets(self):
        # Header
        tk.Label(self.frame, text="Obligation Physical Settlement", 
                font=('Arial', 14, 'bold'), bg=self.bg_color, fg='#2c3e50').pack(pady=10)
        
        # File inputs
        self.obligation_path = tk.StringVar()
        self.stt_path = tk.StringVar()
        self.stamp_duty_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        FileInputWidget(self.frame, "Obligation File:", self.obligation_path)
        FileInputWidget(self.frame, "STT File:", self.stt_path)
        FileInputWidget(self.frame, "Stamp Duty File:", self.stamp_duty_path)
        FileInputWidget(self.frame, "Output Folder:", self.output_path, is_folder=True)
        
        # Generate button
        tk.Button(self.frame, text="Generate Settlement Report", 
                 command=self.on_generate_click, bg='#27ae60', fg='white',
                 font=('Arial', 12, 'bold'), relief=tk.FLAT, padx=30, pady=6).pack(pady=20)
    
    def get_values(self):
        return {
            'obligation_path': self.obligation_path.get(),
            'stt_path': self.stt_path.get(),
            'stamp_duty_path': self.stamp_duty_path.get(),
            'output_path': self.output_path.get()
        }


class SegregationReportPage(BasePage):
    """Segregation Report page with scrolling"""
    def __init__(self, parent, on_generate_click):
        super().__init__(parent)
        self.on_generate_click = on_generate_click
        self.create_widgets()
    
    def create_widgets(self):
        # Create scrollable frame
        canvas = tk.Canvas(self.frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.bg_color)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Store canvas reference for scrolling
        self.canvas = canvas
        self.scrollable_frame = scrollable_frame
        
        # Bind mousewheel - improved scrolling
        def _on_mousewheel(event):
            # Check if the scroll region is larger than the canvas
            if canvas.bbox("all"):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_scroll_to_tab():
            self.frame.bind_all("<MouseWheel>", _on_mousewheel)
        
        def unbind_scroll_from_tab():
            self.frame.unbind_all("<MouseWheel>")
        
        # Bind scrolling events
        self.frame.bind("<Enter>", lambda e: bind_scroll_to_tab())
        self.frame.bind("<Leave>", lambda e: unbind_scroll_from_tab())
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        # Store the mousewheel function for binding to child widgets
        self._mousewheel_handler = _on_mousewheel
        
        # Update scroll region when content changes
        def update_scroll_region():
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Schedule scroll region update
        self.frame.after(100, update_scroll_region)
        
        # Title
        tk.Label(scrollable_frame, text="Segregation Report Generation", 
                font=('Arial', 14, 'bold'), bg=self.bg_color, fg='#2c3e50').pack(pady=10)
        
        # Add scroll instruction
        instruction_label = tk.Label(scrollable_frame, 
                                   text="💡 Tip: Use mouse wheel, arrow keys, or scrollbar to navigate through all fields", 
                                   font=('Arial', 9, 'italic'), 
                                   bg=self.bg_color, fg='#666666')
        instruction_label.pack(pady=(0, 10))

        # Date and CP PAN fields
        yesterday = datetime.now() - timedelta(days=1)
        date_pan_frame = tk.Frame(scrollable_frame, bg=self.bg_color)
        date_pan_frame.pack(pady=8, padx=20, fill=tk.X)

        # Date field
        self.segregation_date_var = tk.StringVar()
        DateInputWidget(date_pan_frame, "Date:", self.segregation_date_var, yesterday)

        # CP PAN field
        tk.Label(date_pan_frame, text="Trading member PAN:", font=('Arial', 12, 'bold'),
                bg=self.bg_color, fg='#2c3e50').pack(side=tk.LEFT)
        self.cp_pan_var = tk.StringVar(value="AACCO4820B")  # Default value
        tk.Entry(date_pan_frame, textvariable=self.cp_pan_var, width=20,
                font=('Arial', 10)).pack(side=tk.LEFT, padx=5)

        # Initialize all file variables
        self._init_file_variables()
        
        # Create file selection sections
        self._create_file_sections(scrollable_frame)
        
        # Master Records Form
        self._create_master_records_form(scrollable_frame)
        
        # Output folder
        FileInputWidget(scrollable_frame, "Output Folder:", self.segregation_output_var, is_folder=True)

        # Generate Button
        generate_segregation_btn = tk.Button(scrollable_frame, text="🚀 Generate Segregation Report",
                                           command=self.on_generate_click, bg='#27ae60', fg='white',
                                           font=('Arial', 14, 'bold'), relief=tk.FLAT, padx=40, pady=8)
        generate_segregation_btn.pack(pady=20)
        
        # Add some bottom padding
        bottom_padding = tk.Frame(scrollable_frame, bg=self.bg_color, height=20)
        bottom_padding.pack(fill=tk.X)
    
    def _init_file_variables(self):
        """Initialize all file variables"""
        self.cash_collateral_cds_var = tk.StringVar()
        self.cash_collateral_fno_var = tk.StringVar()
        self.daily_margin_nsecr_var = tk.StringVar()
        self.daily_margin_nsefno_var = tk.StringVar()
        self.x_cp_master_var = tk.StringVar()
        self.f_cp_master_var = tk.StringVar()
        self.collateral_valuation_cds_var = tk.StringVar()
        self.collateral_valuation_fno_var = tk.StringVar()
        self.sec_pledge_var = tk.StringVar()
        self.cash_with_ncl_var = tk.StringVar()
        self.santom_file_var = tk.StringVar()
        self.extra_records_file = tk.StringVar()
        self.segregation_output_var = tk.StringVar()
        
        # Master Records Form variables
        self.account_type_var = tk.StringVar(value="P")  # Default to P
        self.cp_code_var = tk.StringVar()  # Separate variable for CP Code
        self.segment_var = tk.StringVar(value="FO")
        self.av_value_var = tk.StringVar()
        self.master_records_data = []  # Store records in JSON format
        self.json_file_path = "master_records.json"  # Master JSON file path
        self.selected_record_id = None  # Track selected record for updates
        self.table_type_var = tk.StringVar(value="AV_Records")  # Table type identifier
        
        # Dynamic table configuration - FUTURE-PROOF!
        self.table_configurations = {
            "AV_Records": {
                "columns": ["Account Type", "Segment", "AV Value"],
                "column_widths": [120, 120, 200],
                "fields": [
                    {"name": "Account Type", "type": "combobox", "values": ["P", "C"], "var": "account_type_var", "default": "P"},
                    {"name": "Segment", "type": "combobox", "values": ["CD", "CM", "CO", "FO"], "var": "segment_var", "default": "FO"},
                    {"name": "AV Value", "type": "entry", "var": "av_value_var", "default": ""}
                ],
                "data_mapping": {"account_type": 0, "segment": 1, "av_value": 2}
            },
            "AT_Records": {
                "columns": ["CP Code", "Segment", "AT Value"],
                "column_widths": [120, 120, 200],
                "fields": [
                    {"name": "CP Code", "type": "entry", "var": "cp_code_var", "default": "", "placeholder": "e.g., ICICI4343KL54"},
                    {"name": "Segment", "type": "combobox", "values": ["CD", "CM", "CO", "FO"], "var": "segment_var", "default": "FO"},
                    {"name": "AT Value", "type": "entry", "var": "av_value_var", "default": ""}
                ],
                "data_mapping": {"cp_code": 0, "segment": 1, "at_value": 2}
            }
            # Future table types can be added here easily!
            # Example for future expansion:
            # "NEW_TABLE_TYPE": {
            #     "columns": ["Field1", "Field2", "Field3", "Field4"],
            #     "column_widths": [100, 150, 120, 180],
            #     "fields": [
            #         {"name": "Field1", "type": "combobox", "values": ["A", "B", "C"], "var": "field1_var", "default": "A"},
            #         {"name": "Field2", "type": "entry", "var": "field2_var", "default": "", "placeholder": "Enter value"},
            #         {"name": "Field3", "type": "combobox", "values": ["X", "Y", "Z"], "var": "field3_var", "default": "X"},
            #         {"name": "Field4", "type": "entry", "var": "field4_var", "default": ""}
            #     ],
            #     "data_mapping": {"field1": 0, "field2": 1, "field3": 2, "field4": 3}
            # }
        }
    
    def _create_file_sections(self, parent):
        """Create file selection sections"""
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
            ("Collateral Valuation Report:", [
                ("Collateral Valuation Report CDS", "self.collateral_valuation_cds_var"),
                ("Collateral Valuation Report FNO", "self.collateral_valuation_fno_var")
            ]),
            ("Gsec File:", [
                ("Gsec File", "self.sec_pledge_var")
            ]),
            #  Add manual input before Santom file
            ("Santom File:", [
                ("SANTOM_FILE", "self.santom_file_var"),
                ("Cash with NCL (PROP)", "self.cash_with_ncl_var")
            ]),
            ("Extra Records:", [
                ("Extra_Records_File", "self.extra_records_file"),
            ])
        ]

        for section_title, files in file_frames:
            section_frame = tk.Frame(parent, bg=self.bg_color)
            section_frame.pack(pady=8, padx=20, fill=tk.X)
            
            tk.Label(section_frame, text=section_title, font=('Arial', 12, 'bold'),
                    bg=self.bg_color, fg='#2c3e50').pack(anchor=tk.W)
            
            for file_name, var_name in files:
                file_frame = tk.Frame(section_frame, bg=self.bg_color)
                file_frame.pack(pady=4, fill=tk.X)
                
                tk.Label(file_frame, text=f"  {file_name}:", font=('Arial', 10),
                        bg=self.bg_color, fg='#2c3e50').pack(side=tk.LEFT)
                
                var = getattr(self, var_name.replace('self.', '').replace('_var', '_var'))

                if "cash_with_ncl" in var_name:   # ✅ Manual text input
                    tk.Entry(file_frame, textvariable=var, width=20,
                            font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
                else:
                    tk.Entry(file_frame, textvariable=var, width=60,
                            font=('Arial', 9)).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                    
                    tk.Button(file_frame, text="Browse", command=lambda v=var: self._select_file(v),
                            bg='#3498db', fg='white', font=('Arial', 9)).pack(side=tk.LEFT, padx=5)
    
    def _create_master_records_form(self, parent):
        """Create the dynamic master records form"""
        # Main form frame
        form_frame = tk.Frame(parent, bg=self.bg_color, relief=tk.RIDGE, bd=2)
        form_frame.pack(pady=15, padx=20, fill=tk.X)
        
        # Title
        tk.Label(form_frame, text="📝 Master Records Form", font=('Arial', 14, 'bold'),
                bg=self.bg_color, fg='#2c3e50').pack(pady=(10, 5))
        
        # Controls frame
        controls_frame = tk.Frame(form_frame, bg=self.bg_color)
        controls_frame.pack(pady=10, padx=10, fill=tk.X)
        
        # Table Type Selector (Row 0)
        tk.Label(controls_frame, text="Table Type:", font=('Arial', 11, 'bold'),
                bg=self.bg_color, fg='#2c3e50').grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # Dynamic table type selector - automatically gets all configured table types
        table_types = list(self.table_configurations.keys())
        table_type_combo = ttk.Combobox(controls_frame, textvariable=self.table_type_var,
                                      values=table_types, state='readonly', width=15)
        table_type_combo.grid(row=0, column=1, padx=(0, 20), sticky=tk.W)
        table_type_combo.bind('<<ComboboxSelected>>', self._on_table_type_change)
        
        # First Column (Account Type / CP Code) - Dynamic based on table type
        self.first_column_label = tk.Label(controls_frame, text="Account Type:", font=('Arial', 11, 'bold'),
                                          bg=self.bg_color, fg='#2c3e50')
        self.first_column_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        
        # Combobox for Account Type (AV Records)
        self.account_type_combo = ttk.Combobox(controls_frame, textvariable=self.account_type_var,
                                             values=['P', 'C'], state='readonly', width=15)
        self.account_type_combo.grid(row=1, column=1, padx=(0, 20), sticky=tk.W, pady=(5, 0))
        self.account_type_combo.set('P')  # Default to P
        
        # Text Entry for CP Code (AT Records) - initially hidden
        self.cp_code_entry = tk.Entry(controls_frame, textvariable=self.cp_code_var, width=18, font=('Arial', 10))
        self.cp_code_entry.grid(row=1, column=1, padx=(0, 20), sticky=tk.W, pady=(5, 0))
        self.cp_code_entry.grid_remove()  # Hide initially
        
        # Segment Dropdown (Row 1)
        tk.Label(controls_frame, text="Segment:", font=('Arial', 11, 'bold'),
                bg=self.bg_color, fg='#2c3e50').grid(row=1, column=2, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        
        segment_combo = ttk.Combobox(controls_frame, textvariable=self.segment_var,
                                   values=['CD', 'CM', 'CO', 'FO'], state='readonly', width=15)
        segment_combo.grid(row=1, column=3, padx=(0, 20), sticky=tk.W, pady=(5, 0))
        segment_combo.set('FO')  # Default to FO
        
        
        # Data table frame
        table_frame = tk.Frame(form_frame, bg=self.bg_color)
        table_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # Table title
        tk.Label(table_frame, text="📊 Master Records:", font=('Arial', 12, 'bold'),
                bg=self.bg_color, fg='#2c3e50').pack(anchor=tk.W, pady=(0, 5))
        
        # Create Treeview for data table (will be updated based on table type)
        self.table_columns = ('Account Type', 'Segment', 'AV Value')
        self.records_tree = ttk.Treeview(table_frame, columns=self.table_columns, show='headings', height=6)
        
        # Define initial headings manually (will be updated later)
        for col in self.table_columns:
            self.records_tree.heading(col, text=col)
            self.records_tree.column(col, width=140, anchor='center')
        
        # Scrollbar for table
        table_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.records_tree.yview)
        self.records_tree.configure(yscrollcommand=table_scrollbar.set)
        
        # Pack table and scrollbar
        self.records_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Value input frame
        input_frame = tk.Frame(form_frame, bg=self.bg_color)
        input_frame.pack(pady=5, padx=10, fill=tk.X)
        
        # Dynamic label that changes based on table type
        self.value_label = tk.Label(input_frame, text="AV Value:", font=('Arial', 11, 'bold'),
                                  bg=self.bg_color, fg='#2c3e50')
        self.value_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.av_value_var = tk.StringVar()
        self.av_value_entry = tk.Entry(input_frame, textvariable=self.av_value_var, width=30, font=('Arial', 10))
        self.av_value_entry.pack(side=tk.LEFT, padx=5)
        
        # Bind mousewheel to entry field (will be done after form creation)
        
        # Add Record button (next to AV Value)
        self.add_record_btn = tk.Button(input_frame, text="➕ Add Record", command=self._add_record_to_table,
                                       bg='#27ae60', fg='white', font=('Arial', 10, 'bold'), relief=tk.FLAT)
        self.add_record_btn.pack(side=tk.LEFT, padx=10)
        
        # Update Record button (initially hidden, next to AV Value)
        self.update_record_btn = tk.Button(input_frame, text="✏️ Update Record", command=self._update_record,
                                         bg='#f39c12', fg='white', font=('Arial', 10, 'bold'), relief=tk.FLAT)
        self.update_record_btn.pack(side=tk.LEFT, padx=10)
        self.update_record_btn.pack_forget()  # Hide initially
        
        # Cancel Update button (initially hidden, next to AV Value)
        self.cancel_update_btn = tk.Button(input_frame, text="❌ Cancel", command=self._cancel_update,
                                         bg='#95a5a6', fg='white', font=('Arial', 10, 'bold'), relief=tk.FLAT)
        self.cancel_update_btn.pack(side=tk.LEFT, padx=5)
        self.cancel_update_btn.pack_forget()  # Hide initially
        
        # Buttons frame (for delete and clear operations)
        buttons_frame = tk.Frame(form_frame, bg=self.bg_color)
        buttons_frame.pack(pady=10, padx=10, fill=tk.X)
        
        # Delete Selected button
        delete_record_btn = tk.Button(buttons_frame, text="🗑️ Delete Selected", command=self._delete_selected_record,
                                    bg='#e74c3c', fg='white', font=('Arial', 10, 'bold'), relief=tk.FLAT)
        delete_record_btn.pack(side=tk.LEFT, padx=5)
        
        # Clear All button
        clear_all_btn = tk.Button(buttons_frame, text="🗑️ Clear All", command=self._clear_all_records,
                                bg='#c0392b', fg='white', font=('Arial', 10, 'bold'), relief=tk.FLAT)
        clear_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key to add record
        self.av_value_entry.bind('<Return>', lambda e: self._add_or_update_record())
        
        # Bind table selection event
        self.records_tree.bind('<<TreeviewSelect>>', self._on_record_select)
        
        # Instructions
        instructions = tk.Label(form_frame, 
                              text="AV col name 'Fixed deposit receipt (FDR) placed with NCL', AT col name 'Cash placed with NCL'\n💡 Select Table Type and fill the form, then click 'Add Record'.\nClick on a record to edit it. All data is automatically saved to master JSON file.",
                              font=('Arial', 9, 'italic'), bg=self.bg_color, fg='#666666')
        instructions.pack(pady=(5, 10))
        
        # Setup table columns after all UI elements are created
        self._setup_table_columns()
        
        # Load existing records from JSON file
        self._load_records_from_json()
        
        # Final scroll region update after all content is loaded
        self.frame.after(200, self._update_scroll_region)
        
        # Bind mousewheel to all form widgets after everything is created
        self.frame.after(300, lambda: self._bind_mousewheel_to_all_children(self.scrollable_frame))
    
    def _update_scroll_region(self):
        """Update the scroll region to accommodate all content"""
        if hasattr(self, 'canvas'):
            self.canvas.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _bind_mousewheel_to_widget(self, widget):
        """Bind mousewheel event to a specific widget"""
        if hasattr(self, '_mousewheel_handler'):
            widget.bind("<MouseWheel>", self._mousewheel_handler)
    
    def _bind_mousewheel_to_all_children(self, parent):
        """Recursively bind mousewheel to all child widgets"""
        if hasattr(self, '_mousewheel_handler'):
            try:
                parent.bind("<MouseWheel>", self._mousewheel_handler)
            except:
                pass  # Some widgets might not support binding
            
            # Recursively bind to all children
            for child in parent.winfo_children():
                self._bind_mousewheel_to_all_children(child)
    
    def _setup_table_columns(self):
        """Setup table columns based on table type - FULLY DYNAMIC"""
        table_type = self.table_type_var.get()
        
        # Get configuration for current table type
        if table_type not in self.table_configurations:
            print(f"Warning: Unknown table type '{table_type}'")
            return
        
        config = self.table_configurations[table_type]
        
        # Set table columns dynamically
        self.table_columns = tuple(config["columns"])
        column_widths = config["column_widths"]
        
        # Update value label dynamically (last field name)
        last_field = config["fields"][-1]["name"]
        if hasattr(self, 'value_label'):
            self.value_label.config(text=f"{last_field}:")
        
        # Update first column label dynamically
        first_field = config["fields"][0]["name"]
        if hasattr(self, 'first_column_label'):
            self.first_column_label.config(text=f"{first_field}:")
        
        # Show/hide appropriate input widgets based on field type
        first_field_config = config["fields"][0]
        if hasattr(self, 'account_type_combo') and hasattr(self, 'cp_code_entry'):
            if first_field_config["type"] == "combobox":
                self.account_type_combo.grid()  # Show combobox
                self.cp_code_entry.grid_remove()  # Hide text entry
                # Update combobox values if specified
                if "values" in first_field_config:
                    self.account_type_combo.config(values=first_field_config["values"])
            else:  # entry type
                self.account_type_combo.grid_remove()  # Hide combobox
                self.cp_code_entry.grid()  # Show text entry
        
        # Update treeview columns
        self.records_tree.config(columns=self.table_columns)
        
        # Setup headings and column widths dynamically
        for i, col in enumerate(self.table_columns):
            self.records_tree.heading(col, text=col)
            width = column_widths[i] if i < len(column_widths) else 140
            self.records_tree.column(col, width=width, anchor='center')
    
    def _on_table_type_change(self, event=None):
        """Handle table type change"""
        # Cancel any ongoing update operation first
        if self.selected_record_id is not None:
            self._cancel_update()
        
        # Clear existing records from display
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)
        
        # Setup new table columns
        self._setup_table_columns()
        
        # Clear form inputs and reset to defaults
        self._clear_inputs()
        
        # Reset segment to default FO for both table types
        self.segment_var.set("FO")
        
        # Reload records for the selected table type
        self._load_records_from_json()
    
    def _load_records_from_json(self):
        """Load existing records from JSON file"""
        import json
        import os
        
        if os.path.exists(self.json_file_path):
            try:
                with open(self.json_file_path, 'r') as f:
                    all_data = json.load(f)
                
                # Get current table type
                table_type = self.table_type_var.get()
                
                # Load records for the current table type
                self.master_records_data = all_data.get(table_type, [])
                
                # Populate the table
                for record in self.master_records_data:
                    if table_type == "AV_Records":
                        values = (
                            record.get(G, ''),
                            record.get(H, ''),
                            record.get('av_value', '')
                        )
                    else:  # AT_Records
                        values = (
                            record.get(D, ''),
                            record.get(H, ''),
                            record.get('at_value', '')
                        )
                    self.records_tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"Error loading JSON file: {e}")
                self.master_records_data = []
    
    def _save_records_to_json(self):
        """Save records to master JSON file"""
        import json
        import os
        
        try:
            # Load existing master data
            all_data = {}
            if os.path.exists(self.json_file_path):
                with open(self.json_file_path, 'r') as f:
                    all_data = json.load(f)
            
            # Update data for current table type
            table_type = self.table_type_var.get()
            all_data[table_type] = self.master_records_data
            
            # Save back to file
            with open(self.json_file_path, 'w') as f:
                json.dump(all_data, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save records: {e}")
    
    def _add_or_update_record(self):
        """Add new record or update existing record"""
        if self.selected_record_id is not None:
            self._update_record()
        else:
            self._add_record_to_table()
    
    def _add_record_to_table(self):
        """Add a record to the data table"""
        table_type = self.table_type_var.get()
        segment = self.segment_var.get()
        value = self.av_value_var.get().strip()
        
        # Validation
        if not segment:
            messagebox.showwarning("Warning", "Please select a Segment!")
            return
            
        if not value:
            value_type = "AV Value" if table_type == "AV_Records" else "AT Value"
            messagebox.showwarning("Warning", f"Please enter a {value_type}!")
            return
        
        if table_type == "AV_Records":
            account_type = self.account_type_var.get()
            if not account_type:
                messagebox.showwarning("Warning", "Please select an Account Type!")
                return
            
            # Check for duplicate combination of Account Type + Segment
            if self._check_duplicate_record(account_type, segment, table_type):
                messagebox.showwarning("Duplicate Record", 
                                     f"A record with Account Type '{account_type}' and Segment '{segment}' already exists!")
                return
            
            # Add to data table
            self.records_tree.insert('', 'end', values=(account_type, segment, value))

            # Add to JSON data
            record = {
                'id': len(self.master_records_data),
                 G : account_type,
                 H : segment,
                'av_value': value,
                'table_type': table_type
            }
        else:  # AT_Records
            cp_code = self.cp_code_var.get().strip()  # Get CP Code from separate variable
            if not cp_code:
                messagebox.showwarning("Warning", "Please enter a CP Code (e.g., ICICI4343KL54)!")
                return
            
            # Check for duplicate combination of CP Code + Segment
            if self._check_duplicate_record(cp_code, segment, table_type):
                messagebox.showwarning("Duplicate Record", 
                                     f"A record with CP Code '{cp_code}' and Segment '{segment}' already exists!")
                return
            
            # Add to data table
            self.records_tree.insert('', 'end', values=(cp_code, segment, value))
            
            # Add to JSON data
            record = {
                'id': len(self.master_records_data),
                 D : cp_code,
                 H : segment,
                'at_value': value,
                'table_type': table_type
            }
        
        self.master_records_data.append(record)
        
        # Save to JSON file
        self._save_records_to_json()
        
        # Clear the inputs
        self._clear_inputs()
        
        # Update scroll region
        self.frame.after(50, self._update_scroll_region)
        
        messagebox.showinfo("Success", "Record added successfully!")
    
    def _check_duplicate_record(self, first_field, segment, table_type, exclude_id=None):
        """Check if a record with the same key combination already exists"""
        for record in self.master_records_data:
            # Skip the record being updated (for edit mode)
            if exclude_id is not None and record.get('id') == exclude_id:
                continue
            
            if table_type == "AV_Records":
                # Check for duplicate Account Type + Segment combination
                if (record.get(G) == first_field and 
                    record.get(H) == segment):
                    return True
            else:  # AT_Records
                # Check for duplicate CP Code + Segment combination
                if (record.get(D) == first_field and 
                    record.get(H) == segment):
                    return True
        return False
    
    def _on_record_select(self, event):
        """Handle record selection for editing"""
        selected_item = self.records_tree.selection()
        if not selected_item:
            return
        
        # Get the selected record data
        item = self.records_tree.item(selected_item[0])
        values = item['values']
        
        if values:
            # Fill the form with selected record data based on table type
            table_type = self.table_type_var.get()
            if table_type == "AV_Records":
                self.account_type_var.set(values[0])  # Account Type
            else:  # AT_Records
                self.cp_code_var.set(values[0])  # CP Code
            
            self.segment_var.set(values[1])
            self.av_value_var.set(values[2])
            
            # Get the record ID
            item_index = self.records_tree.index(selected_item[0])
            self.selected_record_id = item_index
            
            # Show update buttons, hide add button
            self.add_record_btn.pack_forget()
            self.update_record_btn.pack(side=tk.LEFT, padx=10)
            self.cancel_update_btn.pack(side=tk.LEFT, padx=5)
    
    def _update_record(self):
        """Update the selected record"""
        if self.selected_record_id is None:
            return
        
        table_type = self.table_type_var.get()
        segment = self.segment_var.get()
        value = self.av_value_var.get().strip()
        
        if table_type == "AV_Records":
            account_type = self.account_type_var.get()
            # Validation
            if not account_type or not segment or not value:
                messagebox.showwarning("Warning", "Please fill all fields!")
                return
            
            # Check for duplicate combination (exclude current record being updated)
            current_record_id = self.master_records_data[self.selected_record_id].get('id') if 0 <= self.selected_record_id < len(self.master_records_data) else None
            if self._check_duplicate_record(account_type, segment, table_type, exclude_id=current_record_id):
                messagebox.showwarning("Duplicate Record", 
                                     f"A record with Account Type '{account_type}' and Segment '{segment}' already exists!")
                return
            
            # Update in data
            if 0 <= self.selected_record_id < len(self.master_records_data):
                self.master_records_data[self.selected_record_id].update({
                    G : account_type,
                    H : segment,
                    'av_value': value
                })
            
            # Update in table
            selected_item = self.records_tree.selection()[0]
            self.records_tree.item(selected_item, values=(account_type, segment, value))
            
        else:  # AT_Records
            cp_code = self.cp_code_var.get().strip()
            # Validation
            if not cp_code or not segment or not value:
                messagebox.showwarning("Warning", "Please fill all fields!")
                return
            
            # Check for duplicate combination (exclude current record being updated)
            current_record_id = self.master_records_data[self.selected_record_id].get('id') if 0 <= self.selected_record_id < len(self.master_records_data) else None
            if self._check_duplicate_record(cp_code, segment, table_type, exclude_id=current_record_id):
                messagebox.showwarning("Duplicate Record", 
                                     f"A record with CP Code '{cp_code}' and Segment '{segment}' already exists!")
                return
            
            # Update in data
            if 0 <= self.selected_record_id < len(self.master_records_data):
                self.master_records_data[self.selected_record_id].update({
                    D : cp_code,
                    H : segment,
                    'at_value': value
                })
            
            # Update in table
            selected_item = self.records_tree.selection()[0]
            self.records_tree.item(selected_item, values=(cp_code, segment, value))
        
        # Save to JSON file
        self._save_records_to_json()
        
        # Reset form
        self._cancel_update()
        messagebox.showinfo("Success", "Record updated successfully!")
    
    def _cancel_update(self):
        """Cancel update operation"""
        self.selected_record_id = None
        self._clear_inputs()
        
        # Show add button, hide update buttons
        self.update_record_btn.pack_forget()
        self.cancel_update_btn.pack_forget()
        self.add_record_btn.pack(side=tk.LEFT, padx=10)
        
        # Clear selection
        self.records_tree.selection_remove(self.records_tree.selection())
    
    def _delete_selected_record(self):
        """Delete selected record from the table"""
        selected_item = self.records_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a record to delete!")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this record?"):
            # Get the index of the selected item
            item_index = self.records_tree.index(selected_item[0])
            
            # Remove from tree
            self.records_tree.delete(selected_item[0])
            
            # Remove from JSON data
            if 0 <= item_index < len(self.master_records_data):
                self.master_records_data.pop(item_index)
                
                # Update IDs for remaining records
                for i, record in enumerate(self.master_records_data):
                    record['id'] = i
            
            # Save to JSON file
            self._save_records_to_json()
            
            # Cancel any ongoing update
            self._cancel_update()
            messagebox.showinfo("Success", "Record deleted successfully!")
    
    def _clear_all_records(self):
        """Clear all records from the table"""
        if not self.master_records_data:
            messagebox.showinfo("Info", "No records to clear!")
            return
            
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all records?\nThis will delete the JSON file!"):
            # Clear tree
            for item in self.records_tree.get_children():
                self.records_tree.delete(item)
            
            # Clear JSON data
            self.master_records_data.clear()
            
            # Save empty data (or delete file)
            self._save_records_to_json()
            
            # Cancel any ongoing update
            self._cancel_update()
            messagebox.showinfo("Success", "All records cleared!")
    
    def _clear_inputs(self):
        """Clear all input fields"""
        self.account_type_var.set("P")  # Reset to default
        self.cp_code_var.set("")  # Clear CP Code
        self.segment_var.set("FO")  # Reset to default FO
        self.av_value_var.set("")
        self.av_value_entry.focus()
    
    def _get_master_records_data(self):
        """Get all master records data in JSON format (both AV and AT records)"""
        return self.master_records_data
    
    def _select_file(self, var):
        """Select file using file dialog"""
        file = filedialog.askopenfilename(title="Select File", 
                                        filetypes=[("All files", "*.*"), 
                                                 ("Excel files", "*.xlsx;*.xls"), 
                                                 ("CSV files", "*.csv"), 
                                                 ("Text files", "*.txt")])
        if file:
            var.set(file)
    
    def _show_file_error(self, file_label, file_path, error_message):
        """Show user-friendly error message for file reading issues"""
        messagebox.showerror(
            "File Reading Error",
            f"❌ Error reading file: {file_label}\n\n"
            f"File path: {file_path}\n\n"
            f"Please check if the correct file is attached and the file format is valid.\n\n"
            f"Expected: {file_label} file\n\n"
            f"Technical details: {str(error_message)}"
        )
    
    def get_values(self):
        return {
            'date': self.segregation_date_var.get(),
            'cp_pan': self.cp_pan_var.get(),
            'cash_collateral_cds': self.cash_collateral_cds_var.get(),
            'cash_collateral_fno': self.cash_collateral_fno_var.get(),
            'daily_margin_nsecr': self.daily_margin_nsecr_var.get(),
            'daily_margin_nsefno': self.daily_margin_nsefno_var.get(),
            'x_cp_master': self.x_cp_master_var.get(),
            'f_cp_master': self.f_cp_master_var.get(),
            'collateral_valuation_cds': self.collateral_valuation_cds_var.get(),
            'collateral_valuation_fno': self.collateral_valuation_fno_var.get(),
            'sec_pledge': self.sec_pledge_var.get(),
            'cash_with_ncl': self.cash_with_ncl_var.get(),   #  Added here
            'santom_file': self.santom_file_var.get(),
            'extra_records': self.extra_records_file.get(),
            'output_path': self.segregation_output_var.get()
        }