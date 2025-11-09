"""
Client Position Page with CP Code CRUD Management
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os

def load_passwords(json_file):
    """Load CP code → password mapping from JSON"""
    with open(json_file, "r") as f:
        data = json.load(f)
    return {item["cp_code"]: item["password"] for item in data}


class ClientPositionPage:
    """Client Position Report page with CP Code Management"""
    
    # Default CP codes (hardcoded) - used for initial creation and reset
    DEFAULT_CP_CODES = [
        {"cp_code": "DBSBK0000033", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000036", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000038", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000041", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000042", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000044", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000043", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000049", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000050", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000051", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000052", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0000475", "password": "AACCO2383D", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000057", "password": "AARCA6399L", "mode": "7z", "add_total": False},
        {"cp_code": "ORBIS0000721", "password": "AAHCD1353P", "mode": "7z", "add_total": False},
        {"cp_code": "90072", "password": "AAVCS8275R", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0005102", "password": "AAHCD2926Q", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000389", "password": "AAICB6686J", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000074", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000077", "password": "AAHCC9973C", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0000718", "password": "AADCF4059J", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000079", "password": "AAICD0896J", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000178", "password": "AAICD1968M", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000179", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000458", "password": "AAICD2891H", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000459", "password": "AAICD3412C", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0005162", "password": "AAHCD7958E", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000189", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001080", "password": "AADCF4059J", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0005402", "password": "AAGCE6929F", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000494", "password": "AAWCA0001C", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000214", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000216", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000488", "password": "AAICD7821M", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000480", "password": "AAICD6359G", "mode": "7z", "add_total": False},
        {"cp_code": "OHMDO0000001", "password": "AACCA2197H", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000540", "password": "AAJCD5624K", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000229", "password": "AAJCD6205G", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000535", "password": "AAJCD4991K", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000192", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000217", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000246", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000247", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "OHMDO0000002", "password": "AAECC6885A", "mode": "7z", "add_total": False},
        {"cp_code": "A3854825I", "password": "AADPD3226A", "mode": "7z", "add_total": False},
        {"cp_code": "OHMDO0000003", "password": "AABCH5042B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000232", "password": "AAGCD0792B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000231", "password": "AAJCD6048F", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0005101", "password": "AAHCD5316Q", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000228", "password": "AAJCD6049E", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000499", "password": "AAICD9377Q", "mode": "7z", "add_total": False},
        {"cp_code": "BNPP00000475", "password": "AAICD5720H", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0005164", "password": "AABCI6920P", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001453", "password": "AAKTA2588D", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000276", "password": "AAKCD4198F", "mode": "7z", "add_total": False},
        {"cp_code": "ORBIS0007660", "password": "AAICD5720H", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001482", "password": "AAKCT7857R", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001479", "password": "AAKCT7815R", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001480", "password": "AAKCT7812J", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001481", "password": "AAKCT7855P", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000289", "password": "AAKCT7813K", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000286", "password": "AAKCT7814Q", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000288", "password": "AAKCT7856Q", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000287", "password": "AAKCT7854N", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000285", "password": "AAKCD6244Q", "mode": "7z", "add_total": False},
        {"cp_code": "OHMDO0000004", "password": "AAHCD0576E", "mode": "7z", "add_total": False},
        {"cp_code": "U8399487I", "password": "AGBPJ7494C", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0005271", "password": "AADCC2348D", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001329", "password": "AAATF9362J", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000299", "password": "AAKCD7324B", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000300", "password": "AAKCD7414N", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000354", "password": "AALCD1141K", "mode": "7z", "add_total": False},
        {"cp_code": "ECASL0000538", "password": "ABMCS3199L", "mode": "7z", "add_total": False},
        {"cp_code": "V6148852I", "password": "ARJPK3191L", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000353", "password": "AALCD1140J", "mode": "7z", "add_total": False},
        {"cp_code": "EP286656I", "password": "AQCPA0001F", "mode": "7z", "add_total": False},
        {"cp_code": "S8274288I", "password": "ATJPD3231G", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000348", "password": "AAETD0225G", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000356", "password": "AAETG6735E", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001706", "password": "AALTA0989R", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0006265", "password": "AAMCC1280F", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000416", "password": "AAGCL1184K", "mode": "7z", "add_total": False},
        {"cp_code": "Z5739353I", "password": "BEHPB0322L", "mode": "7z", "add_total": False},
        {"cp_code": "ECASL0000724", "password": "ABITS1146C", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000360", "password": "AALCT2614F", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000361", "password": "AALCT2613C", "mode": "7z", "add_total": False},
        {"cp_code": "DBNK00009051", "password": "AACTM3577A", "mode": "7z", "add_total": False},
        {"cp_code": "DBNK00009053", "password": "AACTM3577A", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000380", "password": "AALCD2920N", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000381", "password": "AALCD2919D", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000397", "password": "AALCD3274D", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000456", "password": "AAGCL3082L", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000398", "password": "AAJCN6787F", "mode": "7z", "add_total": False},
        {"cp_code": "ICICI0006090", "password": "AAJCN6787F", "mode": "7z", "add_total": False},
        {"cp_code": "KOTBK0001688", "password": "AAKCT7857R", "mode": "7z", "add_total": False},
        {"cp_code": "DBNK00009050", "password": "AACTM3577A", "mode": "7z", "add_total": False},
        {"cp_code": "DBNK00009529", "password": "AACTM3577A", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000458", "password": "AALCD7016M", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000455", "password": "AALCD6828M", "mode": "7z", "add_total": False},
        {"cp_code": "DBSBK0000444", "password": "AALCD5956G", "mode": "7z", "add_total": False},
    ]
    
    def __init__(self, parent, on_process_click, bg_color="#B5D1B1"):
        self.parent = parent
        self.bg_color = bg_color
        self.frame = tk.Frame(parent, bg=bg_color)
        self.on_process_click = on_process_click
        self.cp_codes_data = []  # Store CP codes with checkbox states and passwords
        self.master_json_path = "master_passwords.json"
        self.cash_collateral_path = tk.StringVar()
        self.create_widgets()
        self.load_cp_codes_from_json()  # Load on init
    
    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
    
    def pack_forget(self):
        self.frame.pack_forget()
    
    def create_widgets(self):
        # Header
        header_label = tk.Label(self.frame, text="Client Position Report", 
                               font=('Arial', 16, 'bold'), bg=self.bg_color, fg='#2c3e50')
        header_label.pack(pady=8)
        
        # Main container with two columns
        main_container = tk.Frame(self.frame, bg=self.bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Left side - File inputs and process
        left_frame = tk.Frame(main_container, bg=self.bg_color, width=350)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False, padx=(0, 10))
        left_frame.pack_propagate(False)
        
        # File inputs section
        file_section = tk.LabelFrame(left_frame, text="📁 File Selection", 
                                     font=('Arial', 11, 'bold'), bg=self.bg_color, 
                                     fg='#2c3e50', padx=10, pady=10)
        file_section.pack(fill=tk.X, pady=(0, 15))
        
        self.client_position_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # Client Position File
        tk.Label(file_section, text="Client Position File:", font=('Arial', 10, 'bold'), 
                bg=self.bg_color, fg='#2c3e50').pack(anchor='w', pady=(5, 2))
        
        file_frame1 = tk.Frame(file_section, bg=self.bg_color)
        file_frame1.pack(fill=tk.X, pady=(0, 10))
        
        tk.Entry(file_frame1, textvariable=self.client_position_path, 
                font=('Arial', 9), width=25).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(file_frame1, text="📂", command=lambda: self._browse_file(self.client_position_path),
                 bg='#3498db', fg='white', font=('Arial', 9, 'bold'), 
                 relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=(5, 0))
        
        # Optional Cash Collateral file
        tk.Label(file_section, text="Cash Collateral File (Optional):", font=('Arial', 10, 'bold'),
                bg=self.bg_color, fg='#2c3e50').pack(anchor='w', pady=(5, 2))

        file_frame3 = tk.Frame(file_section, bg=self.bg_color)
        file_frame3.pack(fill=tk.X, pady=(0, 10))

        tk.Entry(file_frame3, textvariable=self.cash_collateral_path,
                font=('Arial', 9), width=25).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(file_frame3, text="📂", command=lambda: self._browse_file(self.cash_collateral_path),
                 bg='#3498db', fg='white', font=('Arial', 9, 'bold'),
                 relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=(5, 0))

        # Output Folder (placed last)
        tk.Label(file_section, text="Output Folder:", font=('Arial', 10, 'bold'), 
                bg=self.bg_color, fg='#2c3e50').pack(anchor='w', pady=(5, 2))
        
        file_frame2 = tk.Frame(file_section, bg=self.bg_color)
        file_frame2.pack(fill=tk.X)
        
        tk.Entry(file_frame2, textvariable=self.output_path, 
                font=('Arial', 9), width=25).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(file_frame2, text="📂", command=lambda: self._browse_folder(self.output_path),
                 bg='#3498db', fg='white', font=('Arial', 9, 'bold'), 
                 relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=(5, 0))
        
        # Process button
        process_btn = tk.Button(left_frame, text="🚀 Process Client Position\nReport", 
                               command=self.on_process_click, bg='#27ae60', fg='white', 
                               font=('Arial', 12, 'bold'), relief=tk.FLAT, 
                               padx=20, pady=15, wraplength=200)
        process_btn.pack(pady=20, fill=tk.X)
        
        # Info section
        info_frame = tk.LabelFrame(left_frame, text="ℹ️ Information", 
                                  font=('Arial', 10, 'bold'), bg=self.bg_color, 
                                  fg='#2c3e50', padx=10, pady=10)
        info_frame.pack(fill=tk.BOTH, expand=True)
        
        info_text = (
            "• Select CP codes from the right panel\n"
            "• Double-click to edit\n"
            "• Click checkbox to select/deselect\n"
            "• Use CRUD buttons to manage\n"
            "• Save changes to JSON file"
        )
        tk.Label(info_frame, text=info_text, font=('Arial', 9), 
                bg=self.bg_color, fg='#34495e', justify=tk.LEFT).pack(anchor='w')
        
        # Right side - CP Code Management
        right_frame = tk.Frame(main_container, bg=self.bg_color)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # CP Code Management Section Header
        header_frame = tk.Frame(right_frame, bg=self.bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(header_frame, text="CP Code Management", 
                font=('Arial', 14, 'bold'), bg=self.bg_color, fg='#2c3e50').pack(side=tk.LEFT)
        
        # CRUD Buttons
        crud_frame = tk.Frame(header_frame, bg=self.bg_color)
        crud_frame.pack(side=tk.RIGHT)
        
        tk.Button(crud_frame, text="➕ Add", command=self.add_cp_code,
                 bg='#3498db', fg='white', font=('Arial', 9, 'bold'), 
                 relief=tk.FLAT, padx=12, pady=5).pack(side=tk.LEFT, padx=2)
        
        tk.Button(crud_frame, text="✏️ Edit", command=self.edit_cp_code,
                 bg='#f39c12', fg='white', font=('Arial', 9, 'bold'), 
                 relief=tk.FLAT, padx=12, pady=5).pack(side=tk.LEFT, padx=2)
        
        tk.Button(crud_frame, text="🗑️ Delete", command=self.delete_cp_code,
                 bg='#e74c3c', fg='white', font=('Arial', 9, 'bold'), 
                 relief=tk.FLAT, padx=12, pady=5).pack(side=tk.LEFT, padx=2)
        
        # Search and filter
        search_frame = tk.Frame(right_frame, bg=self.bg_color)
        search_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(search_frame, text="🔍", font=('Arial', 12), 
                bg=self.bg_color).pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', self._on_search)
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, 
                               font=('Arial', 10), width=30)
        search_entry.pack(side=tk.LEFT, padx=(0, 15))
        
        # Selection count label
        self.selection_label = tk.Label(search_frame, text="Selected: 0 / 0", 
                                       font=('Arial', 10, 'bold'), 
                                       bg=self.bg_color, fg='#27ae60')
        self.selection_label.pack(side=tk.RIGHT)
        
        # Datatable frame with scrollbar (limited height to show buttons below)
        table_container = tk.Frame(right_frame, bg=self.bg_color)
        table_container.pack(fill=tk.BOTH, expand=True)
        
        table_frame = tk.Frame(table_container, bg='white', relief=tk.SUNKEN, bd=2, height=380)
        table_frame.pack(fill=tk.BOTH, expand=False, side=tk.TOP)
        table_frame.pack_propagate(False)  # Maintain fixed height
        
        # Create Treeview for datatable
        columns = ('select', 'cp_code', 'password', 'mode', 'add_total')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', 
                                height=15, selectmode='browse')
        
        # Define column headings
        self.tree.heading('select', text='☐')
        self.tree.heading('cp_code', text='CP Code')
        self.tree.heading('password', text='Password')
        self.tree.heading('mode', text='Mode')
        self.tree.heading('add_total', text='Total')
        
        # Define column widths
        self.tree.column('select', width=50, anchor='center')
        self.tree.column('cp_code', width=150, anchor='w')
        self.tree.column('password', width=150, anchor='w')
        self.tree.column('mode', width=80, anchor='center')
        self.tree.column('add_total', width=70, anchor='center')
        
        # Add scrollbars
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout for tree and scrollbars
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Bind click event for checkbox toggle
        self.tree.bind('<Button-1>', self._on_tree_click)
        self.tree.bind('<Double-1>', self.edit_cp_code)
        
        # Separator line
        separator = tk.Frame(table_container, bg='#d0d0d0', height=1)
        separator.pack(fill=tk.X, pady=(8, 0), side=tk.TOP)
        
        # Bulk action buttons (placed below table)
        button_frame = tk.Frame(table_container, bg=self.bg_color)
        button_frame.pack(fill=tk.X, pady=(8, 5), side=tk.TOP)
        
        tk.Button(button_frame, text="✓ Select All", command=self._select_all,
                 bg='#3498db', fg='white', font=('Arial', 10), 
                 relief=tk.FLAT, padx=15, pady=6).pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(button_frame, text="✗ Deselect All", command=self._deselect_all,
                 bg='#95a5a6', fg='white', font=('Arial', 10), 
                 relief=tk.FLAT, padx=15, pady=6).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(button_frame, text="⚙ Update Mode/Total", command=self.bulk_update_all,
                 bg='#d35400', fg='white', font=('Arial', 10, 'bold'),
                 relief=tk.FLAT, padx=15, pady=6).pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(button_frame, text="🔄 Reset to Default", command=self.reset_to_default,
                 bg='#9b59b6', fg='white', font=('Arial', 10), 
                 relief=tk.FLAT, padx=15, pady=6).pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(button_frame, text="💾 Save Changes", command=self.save_to_json,
                 bg='#16a085', fg='white', font=('Arial', 10, 'bold'), 
                 relief=tk.FLAT, padx=20, pady=6).pack(side=tk.RIGHT)
    
    def _browse_file(self, var):
        """Browse for file"""
        from tkinter import filedialog
        filename = filedialog.askopenfilename(
            title="Select Client Position File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
    
    def _browse_folder(self, var):
        """Browse for folder"""
        from tkinter import filedialog
        foldername = filedialog.askdirectory(title="Select Output Folder")
        if foldername:
            var.set(foldername)
    
    def load_cp_codes_from_json(self):
        """Load CP codes from master_passwords.json - creates with defaults if not exists"""
        try:
            # If file doesn't exist, create it with default CP codes
            if not os.path.exists(self.master_json_path):
                self._create_default_json_file()
                return  # _create_default_json_file will load data
            
            # File exists, load it
            with open(self.master_json_path, 'r') as f:
                data = json.load(f)
            
            self.cp_codes_data = []
            
            # Handle array format
            if isinstance(data, list):
                for item in data:
                    cp_code_value = item.get('cp_code', '')
                    # Ensure CP code is stored as string
                    if cp_code_value is not None:
                        cp_code_value = str(cp_code_value)
                    
                    self.cp_codes_data.append({
                        'cp_code': cp_code_value,
                        'password': item.get('password', '123'),
                        'mode': item.get('mode', '7z'),
                        'add_total': item.get('add_total', False),
                        'selected': False  # Default unchecked
                    })
            # Handle object format
            elif isinstance(data, dict):
                for cp_code, config in data.items():
                    # Ensure CP code is stored as string
                    cp_code_str = str(cp_code) if cp_code is not None else ''
                    
                    self.cp_codes_data.append({
                        'cp_code': cp_code_str,
                        'password': config.get('password', '123'),
                        'mode': config.get('mode', '7z'),
                        'add_total': config.get('add_total', False),
                        'selected': False  # Default unchecked
                    })
            
            self._refresh_tree()
        
        except Exception as e:
            messagebox.showerror("❌ Error", f"Failed to load CP codes:\n{str(e)}")
    
    def _create_default_json_file(self):
        """Create master_passwords.json with default CP codes"""
        try:
            with open(self.master_json_path, 'w') as f:
                json.dump(self.DEFAULT_CP_CODES, f, indent=2)
            
            # Load the defaults into memory
            self.cp_codes_data = []
            for item in self.DEFAULT_CP_CODES:
                self.cp_codes_data.append({
                    'cp_code': item['cp_code'],
                    'password': item['password'],
                    'mode': item['mode'],
                    'add_total': item['add_total'],
                    'selected': False
                })
            
            self._refresh_tree()
            # File created silently without showing message
        
        except Exception as e:
            messagebox.showerror("❌ Error", f"Failed to create default JSON file:\n{str(e)}")
    
    def reset_to_default(self):
        """Reset all CP codes to hardcoded defaults"""
        result = messagebox.askyesno(
            "🔄 Confirm Reset to Default",
            f"This will:\n\n"
            f"• Delete all current CP codes\n"
            f"• Restore {len(self.DEFAULT_CP_CODES)} default CP codes\n"
            f"• Overwrite {self.master_json_path}\n\n"
            f"⚠️ This action cannot be undone!\n\n"
            f"Do you want to continue?"
        )
        
        if result:
            try:
                # Write defaults to JSON file
                with open(self.master_json_path, 'w') as f:
                    json.dump(self.DEFAULT_CP_CODES, f, indent=2)
                
                # Load defaults into memory
                self.cp_codes_data = []
                for item in self.DEFAULT_CP_CODES:
                    self.cp_codes_data.append({
                        'cp_code': item['cp_code'],
                        'password': item['password'],
                        'mode': item['mode'],
                        'add_total': item['add_total'],
                        'selected': False
                    })
                
                self._refresh_tree()
                messagebox.showinfo("✅ Reset Complete", 
                    f"Successfully reset to {len(self.DEFAULT_CP_CODES)} default CP codes!")
            
            except Exception as e:
                messagebox.showerror("❌ Error", f"Failed to reset to defaults:\n{str(e)}")
    
    def _auto_save_to_json(self):
        """Auto-save current CP codes data to JSON file (silent)"""
        try:
            # Convert to array format
            data = []
            for item in self.cp_codes_data:
                data.append({
                    'cp_code': item['cp_code'],
                    'password': item['password'],
                    'mode': item['mode'],
                    'add_total': item['add_total']
                })
            
            with open(self.master_json_path, 'w') as f:
                json.dump(data, f, indent=2)
        
        except Exception as e:
            messagebox.showerror("❌ Error", f"Failed to auto-save CP codes:\n{str(e)}")
    
    def save_to_json(self):
        """Manually save current CP codes data to JSON file"""
        try:
            # Convert to array format
            data = []
            for item in self.cp_codes_data:
                data.append({
                    'cp_code': item['cp_code'],
                    'password': item['password'],
                    'mode': item['mode'],
                    'add_total': item['add_total']
                })
            
            with open(self.master_json_path, 'w') as f:
                json.dump(data, f, indent=2)
            
            messagebox.showinfo("✅ Success", 
                f"Successfully saved {len(data)} CP codes to\n{self.master_json_path}")
        
        except Exception as e:
            messagebox.showerror("❌ Error", f"Failed to save CP codes:\n{str(e)}")
    
    def add_cp_code(self):
        """Add new CP code"""
        dialog = tk.Toplevel(self.frame)
        dialog.title("➕ Add New CP Code")
        dialog.geometry("450x280")
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f8f0')
        
        # Center the dialog
        dialog.transient(self.frame)
        dialog.grab_set()
        
        # Form fields
        tk.Label(dialog, text="CP Code:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=0, column=0, sticky='w', padx=20, pady=12)
        cp_code_entry = tk.Entry(dialog, font=('Arial', 10), width=35)
        cp_code_entry.grid(row=0, column=1, padx=20, pady=12)
        cp_code_entry.focus()
        
        tk.Label(dialog, text="Password:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=1, column=0, sticky='w', padx=20, pady=12)
        password_entry = tk.Entry(dialog, font=('Arial', 10), width=35)
        password_entry.insert(0, "123")
        password_entry.grid(row=1, column=1, padx=20, pady=12)
        
        tk.Label(dialog, text="Mode:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=2, column=0, sticky='w', padx=20, pady=12)
        mode_var = tk.StringVar(value="7z")
        mode_frame = tk.Frame(dialog, bg='#f0f8f0')
        mode_frame.grid(row=2, column=1, sticky='w', padx=20, pady=12)
        tk.Radiobutton(mode_frame, text="ZIP", variable=mode_var, value="zip", 
                      bg='#f0f8f0', font=('Arial', 10)).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="7z", variable=mode_var, value="7z", 
                      bg='#f0f8f0', font=('Arial', 10)).pack(side=tk.LEFT, padx=10)
        
        tk.Label(dialog, text="Add Total Row:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=3, column=0, sticky='w', padx=20, pady=12)
        add_total_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dialog, text="Yes, add total row", variable=add_total_var, 
                      bg='#f0f8f0', font=('Arial', 10)).grid(row=3, column=1, 
                      sticky='w', padx=20, pady=12)
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg='#f0f8f0')
        btn_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        def save_new():
            cp_code = cp_code_entry.get().strip()
            if not cp_code:
                messagebox.showwarning("⚠️ Validation Error", "CP Code cannot be empty!", parent=dialog)
                return
            
            # Check if CP code already exists
            if any(item['cp_code'] == cp_code for item in self.cp_codes_data):
                messagebox.showwarning("⚠️ Duplicate", 
                    f"CP Code '{cp_code}' already exists!", parent=dialog)
                return
            
            self.cp_codes_data.append({
                'cp_code': cp_code,
                'password': password_entry.get().strip() or '123',
                'mode': mode_var.get(),
                'add_total': add_total_var.get(),
                'selected': False
            })
            
            self._refresh_tree_and_focus(cp_code, reset_filter=True, deselect_all=True)
            dialog.destroy()
            
            # Auto-save to JSON file
            self._auto_save_to_json()
            
            messagebox.showinfo("✅ Success", 
                f"Added CP Code: {cp_code}\n\nAutomatically saved to {self.master_json_path}")
        
        tk.Button(btn_frame, text="💾 Save", command=save_new,
                 bg='#27ae60', fg='white', font=('Arial', 10, 'bold'), 
                 relief=tk.FLAT, padx=25, pady=8).pack(side=tk.LEFT, padx=5)
        
        def cancel_add():
            self._refresh_tree_and_focus(reset_filter=True, deselect_all=True)
            dialog.destroy()
        
        tk.Button(btn_frame, text="✖ Cancel", command=cancel_add,
                 bg='#95a5a6', fg='white', font=('Arial', 10), 
                 relief=tk.FLAT, padx=25, pady=8).pack(side=tk.LEFT, padx=5)
    
    def edit_cp_code(self, event=None):
        """Edit selected CP code"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("⚠️ No Selection", "Please select a CP code to edit.")
            return
        
        item = self.tree.item(selection[0])
        values = item['values']
        cp_code = values[1]
        
        # Find the CP code data - ensure both are strings for comparison
        cp_code_str = str(cp_code) if cp_code is not None else ''
        cp_data = next((item for item in self.cp_codes_data if str(item['cp_code']) == cp_code_str), None)
        if not cp_data:
            # Debug: Show available CP codes and types
            available_codes = [item['cp_code'] for item in self.cp_codes_data]
            cp_code_type = type(cp_code).__name__
            available_types = [type(item['cp_code']).__name__ for item in self.cp_codes_data[:5]]
            
            messagebox.showerror("❌ Error", 
                f"CP code '{cp_code}' (type: {cp_code_type}) not found in data.\n\n"
                f"Looking for: '{cp_code}'\n"
                f"Available CP codes: {', '.join(available_codes[:10])}{'...' if len(available_codes) > 10 else ''}\n"
                f"Data types: {', '.join(available_types)}")
            return
        
        dialog = tk.Toplevel(self.frame)
        dialog.title(f"✏️ Edit CP Code: {cp_code}")
        dialog.geometry("450x280")
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f8f0')
        
        dialog.transient(self.frame)
        dialog.grab_set()
        
        # Form fields (CP Code is read-only)
        tk.Label(dialog, text="CP Code:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=0, column=0, sticky='w', padx=20, pady=12)
        cp_code_label = tk.Label(dialog, text=cp_code, font=('Arial', 10), 
                                 bg='#e8f5e9', relief=tk.SUNKEN, width=33, anchor='w', padx=5)
        cp_code_label.grid(row=0, column=1, padx=20, pady=12)
        
        tk.Label(dialog, text="Password:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=1, column=0, sticky='w', padx=20, pady=12)
        password_entry = tk.Entry(dialog, font=('Arial', 10), width=35)
        password_entry.insert(0, cp_data['password'])
        password_entry.grid(row=1, column=1, padx=20, pady=12)
        password_entry.focus()
        
        tk.Label(dialog, text="Mode:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=2, column=0, sticky='w', padx=20, pady=12)
        mode_var = tk.StringVar(value=cp_data['mode'])
        mode_frame = tk.Frame(dialog, bg='#f0f8f0')
        mode_frame.grid(row=2, column=1, sticky='w', padx=20, pady=12)
        tk.Radiobutton(mode_frame, text="ZIP", variable=mode_var, value="zip", 
                      bg='#f0f8f0', font=('Arial', 10)).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="7z", variable=mode_var, value="7z", 
                      bg='#f0f8f0', font=('Arial', 10)).pack(side=tk.LEFT, padx=10)
        
        tk.Label(dialog, text="Add Total Row:", font=('Arial', 10, 'bold'), 
                bg='#f0f8f0').grid(row=3, column=0, sticky='w', padx=20, pady=12)
        add_total_var = tk.BooleanVar(value=cp_data['add_total'])
        tk.Checkbutton(dialog, text="Yes, add total row", variable=add_total_var, 
                      bg='#f0f8f0', font=('Arial', 10)).grid(row=3, column=1, 
                      sticky='w', padx=20, pady=12)
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg='#f0f8f0')
        btn_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        def save_changes():
            cp_data['password'] = password_entry.get().strip() or '123'
            cp_data['mode'] = mode_var.get()
            cp_data['add_total'] = add_total_var.get()
            
            self._refresh_tree_and_focus(cp_code, reset_filter=True, deselect_all=True)
            dialog.destroy()
            
            # Auto-save to JSON file
            self._auto_save_to_json()
            
            messagebox.showinfo("✅ Success", 
                f"Updated CP Code: {cp_code}\n\nAutomatically saved to {self.master_json_path}")
        
        tk.Button(btn_frame, text="💾 Update", command=save_changes,
                 bg='#f39c12', fg='white', font=('Arial', 10, 'bold'), 
                 relief=tk.FLAT, padx=25, pady=8).pack(side=tk.LEFT, padx=5)
        
        def cancel_edit():
            self._refresh_tree_and_focus(reset_filter=True, deselect_all=True)
            dialog.destroy()
        
        tk.Button(btn_frame, text="✖ Cancel", command=cancel_edit,
                 bg='#95a5a6', fg='white', font=('Arial', 10), 
                 relief=tk.FLAT, padx=25, pady=8).pack(side=tk.LEFT, padx=5)
    
    def delete_cp_code(self):
        """Delete selected CP code"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("⚠️ No Selection", "Please select a CP code to delete.")
            return
        
        item = self.tree.item(selection[0])
        cp_code = item['values'][1]
        
        result = messagebox.askyesno("🗑️ Confirm Delete", 
                                     f"Are you sure you want to delete?\n\nCP Code: {cp_code}")
        
        if result:
            self.cp_codes_data = [item for item in self.cp_codes_data 
                                 if item['cp_code'] != cp_code]
            self._refresh_tree()
            
            # Auto-save to JSON file
            self._auto_save_to_json()
            
            messagebox.showinfo("✅ Success", 
                f"Deleted CP Code: {cp_code}\n\nAutomatically saved to {self.master_json_path}")
    
    def _refresh_tree(self, filter_text=None):
        """Refresh the tree view with current data"""
        if filter_text is None:
            filter_text = self.search_var.get()
        
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Filter data if search text provided
        if filter_text:
            filtered_data = [d for d in self.cp_codes_data 
                           if filter_text.lower() in d['cp_code'].lower() or
                              filter_text.lower() in d['password'].lower()]
        else:
            filtered_data = self.cp_codes_data
        
        # Add items to tree
        for data in filtered_data:
            checkbox = '☑' if data['selected'] else '☐'
            add_total_text = '✓' if data['add_total'] else '✗'
            
            self.tree.insert('', tk.END, values=(
                checkbox, 
                data['cp_code'], 
                data['password'],
                data['mode'].upper(),
                add_total_text
            ), tags=('selected' if data['selected'] else 'unselected',))
        
        # Configure tags for visual feedback
        self.tree.tag_configure('selected', background='#e8f5e9', font=('Arial', 9, 'bold'))
        self.tree.tag_configure('unselected', background='white', font=('Arial', 9))
        
        # Update selection count
        selected_count = sum(1 for d in self.cp_codes_data if d['selected'])
        total_count = len(self.cp_codes_data)
        self.selection_label.config(text=f"Selected: {selected_count} / {total_count}")
    
    def _refresh_tree_and_focus(self, cp_code=None, reset_filter=False, deselect_all=False):
        """Refresh tree with options to reset filter, focus a CP code, and clear selections."""

        if deselect_all:
            for data in self.cp_codes_data:
                data['selected'] = False

        if reset_filter:
            self.search_var.set('')
            # Ensure table refreshes immediately even if trace callbacks are delayed
            self._refresh_tree('')
        else:
            self._refresh_tree(self.search_var.get())
            if cp_code is not None and not deselect_all:
                self._focus_tree_item(cp_code)
            return

        # When reset_filter=True, the trace callback already refreshed the tree.
        if cp_code is not None and not deselect_all:
            self._focus_tree_item(cp_code)
    
    def _focus_tree_item(self, cp_code):
        """Ensure the specified CP code row remains focused after refresh"""
        cp_code_str = str(cp_code) if cp_code is not None else ''
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if values and str(values[1]) == cp_code_str:
                self.tree.selection_set(item)
                self.tree.focus(item)
                self.tree.see(item)
                break
    
    def _on_tree_click(self, event):
        """Handle tree click to toggle checkbox"""
        # Don't handle single clicks if this might be part of a double-click
        if hasattr(self, '_click_time'):
            current_time = event.time
            if current_time - self._click_time < 300:  # Less than 300ms, likely part of double-click
                return
        
        self._click_time = event.time
        
        region = self.tree.identify_region(event.x, event.y)
        if region == 'cell':
            item = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)
            
            if item:
                values = self.tree.item(item, 'values')
                cp_code = values[1]
                
                # Toggle if clicking on select column or CP code column
                if column in ('#1', '#2'):  # Select column or CP Code column
                    cp_code_str = str(cp_code) if cp_code is not None else ''
                    for data in self.cp_codes_data:
                        if str(data['cp_code']) == cp_code_str:
                            data['selected'] = not data['selected']
                            break
                    
                    self._refresh_tree(self.search_var.get())
                    self._focus_tree_item(cp_code_str)
    
    def _on_search(self, *args):
        """Handle search text change"""
        search_text = self.search_var.get()
        self._refresh_tree(search_text)
    
    def bulk_update_all(self):
        """Bulk update mode and total settings for every CP code"""
        if not self.cp_codes_data:
            messagebox.showinfo("ℹ️ No Records", "There are no CP codes available to update.")
            return
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("⚙ Update Mode / Total For All")
        dialog.geometry("360x220")
        dialog.resizable(False, False)
        dialog.configure(bg='#fdf7f0')
        
        dialog.transient(self.frame)
        dialog.grab_set()
        
        # Mode selection controls
        apply_mode_var = tk.BooleanVar(value=True)
        tk.Checkbutton(dialog, text="Update Mode", variable=apply_mode_var,
                       bg='#fdf7f0', font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=2,
                                                                       sticky='w', padx=20, pady=(20, 6))
        
        tk.Label(dialog, text="Mode:", font=('Arial', 10),
                 bg='#fdf7f0').grid(row=1, column=0, sticky='w', padx=20, pady=4)
        
        first_mode = (self.cp_codes_data[0].get('mode') or '7z').lower()
        mode_var = tk.StringVar(value='zip' if first_mode == 'zip' else '7z')
        
        mode_frame = tk.Frame(dialog, bg='#fdf7f0')
        mode_frame.grid(row=1, column=1, sticky='w', padx=20, pady=4)
        mode_zip_btn = tk.Radiobutton(mode_frame, text="ZIP", variable=mode_var, value="zip",
                                      bg='#fdf7f0', font=('Arial', 10))
        mode_zip_btn.pack(side=tk.LEFT, padx=5)
        mode_7z_btn = tk.Radiobutton(mode_frame, text="7z", variable=mode_var, value="7z",
                                     bg='#fdf7f0', font=('Arial', 10))
        mode_7z_btn.pack(side=tk.LEFT, padx=5)
        
        # Total row controls
        apply_total_var = tk.BooleanVar(value=True)
        tk.Checkbutton(dialog, text="Update Total Setting", variable=apply_total_var,
                       bg='#fdf7f0', font=('Arial', 10, 'bold')).grid(row=2, column=0, columnspan=2,
                                                                       sticky='w', padx=20, pady=(12, 6))
        
        tk.Label(dialog, text="Add Total Row:", font=('Arial', 10),
                 bg='#fdf7f0').grid(row=3, column=0, sticky='w', padx=20, pady=4)
        
        add_total_var = tk.BooleanVar(value=bool(self.cp_codes_data[0].get('add_total')))
        total_checkbox = tk.Checkbutton(dialog, text="Yes, include total row", variable=add_total_var,
                                        bg='#fdf7f0', font=('Arial', 10))
        total_checkbox.grid(row=3, column=1, sticky='w', padx=20, pady=4)
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg='#fdf7f0')
        btn_frame.grid(row=4, column=0, columnspan=2, pady=25)
        
        def toggle_mode_state():
            state = tk.NORMAL if apply_mode_var.get() else tk.DISABLED
            mode_zip_btn.configure(state=state)
            mode_7z_btn.configure(state=state)
        
        def toggle_total_state():
            state = tk.NORMAL if apply_total_var.get() else tk.DISABLED
            total_checkbox.configure(state=state)
        
        def apply_changes():
            if not apply_mode_var.get() and not apply_total_var.get():
                messagebox.showwarning("⚠️ No Update Selected", "Enable Mode and/or Total before applying changes.")
                return
            
            summary_parts = []
            
            if apply_mode_var.get():
                new_mode = mode_var.get()
                for data in self.cp_codes_data:
                    data['mode'] = new_mode
                summary_parts.append(f"mode '{new_mode.upper()}'")
            
            if apply_total_var.get():
                new_total = add_total_var.get()
                for data in self.cp_codes_data:
                    data['add_total'] = new_total
                summary_parts.append(f"total row {'ON' if new_total else 'OFF'}")
            
            self._refresh_tree(self.search_var.get())
            dialog.destroy()
            
            # Auto-save to JSON file
            self._auto_save_to_json()
            
            messagebox.showinfo("✅ Update Complete",
                                f"Updated all {len(self.cp_codes_data)} CP code(s) → {', '.join(summary_parts)}.")
        
        # Initialize states
        toggle_mode_state()
        toggle_total_state()
        apply_mode_var.trace_add('write', lambda *_: toggle_mode_state())
        apply_total_var.trace_add('write', lambda *_: toggle_total_state())
        
        def cancel_changes():
            dialog.destroy()
        
        tk.Button(btn_frame, text="💾 Apply", command=apply_changes,
                 bg='#27ae60', fg='white', font=('Arial', 10, 'bold'),
                 relief=tk.FLAT, padx=25, pady=8).pack(side=tk.LEFT, padx=8)
        
        tk.Button(btn_frame, text="✖ Cancel", command=cancel_changes,
                 bg='#95a5a6', fg='white', font=('Arial', 10),
                 relief=tk.FLAT, padx=25, pady=8).pack(side=tk.LEFT, padx=8)
        
    def _select_all(self):
        """Select all CP codes"""
        for data in self.cp_codes_data:
            data['selected'] = True
        self._refresh_tree(self.search_var.get())
    
    def _deselect_all(self):
        """Deselect all CP codes"""
        for data in self.cp_codes_data:
            data['selected'] = False
        self._refresh_tree(self.search_var.get())
    
    def get_selected_cp_codes(self):
        """Get list of selected CP codes with their settings"""
        return [d for d in self.cp_codes_data if d['selected']]
    
    def get_values(self):
        selected = self.get_selected_cp_codes()
        return {
            'client_position_path': self.client_position_path.get(),
            'output_path': self.output_path.get(),
            'cash_collateral_path': self.cash_collateral_path.get(),
            'selected_cp_codes': [d['cp_code'] for d in selected],
            'cp_codes_config': {d['cp_code']: {
                'password': d['password'],
                'mode': d['mode'],
                'add_total': d['add_total']
            } for d in selected}
        }

