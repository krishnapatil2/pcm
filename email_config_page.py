"""
Email Configuration Page for PCM Application
Comprehensive email configuration interface with normal and compact views
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from email_sender import EmailSender
from utils import Constants


class EmailConfigPage:
    """Email configuration page with normal view"""
    
    def __init__(self, parent):
        self.parent = parent
        self.email_sender = EmailSender()
        self.attachments = []
        
        # Create main frame with company legacy color
        self.frame = tk.Frame(parent, bg=Constants.PROCESSING_BG)
        
        # Create normal view
        self._create_normal_view()
    
    def _create_normal_view(self):
        """Create normal (full) view of email configuration"""
        # Main container with scrolling
        self._create_scrollable_frame()
        
        # Header
        header_frame = tk.Frame(self.scrollable_frame, bg=Constants.PROCESSING_BG, relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = tk.Label(header_frame, text="📧 Email Configuration", 
                              font=Constants.HEADER_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT)
        title_label.pack(pady=(15, 5))
        
        subtitle_label = tk.Label(header_frame, text="Configure email settings and send emails", 
                               font=Constants.SMALL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.SECONDARY_TEXT)
        subtitle_label.pack(pady=(0, 15))
        
        # Configuration sections
        self._create_smtp_config_section()
        self._create_email_composition_section()
        self._create_attachments_section()
        self._create_action_buttons_section()
        
        # Load current configuration
        self._load_configuration()
        
        # Bind mousewheel to all child widgets for better scrolling coverage
        self.frame.after(100, self._bind_mousewheel_to_all_children)
    
    
    def _create_scrollable_frame(self):
        """Create scrollable frame for normal view"""
        # Create canvas and scrollbar
        self.canvas = tk.Canvas(self.frame, bg=Constants.PROCESSING_BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=Constants.PROCESSING_BG)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel
        def _on_mousewheel(event):
            if self.canvas.bbox("all"):
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind mousewheel to canvas
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # Bind mousewheel to scrollable frame
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        # Bind mousewheel to the main frame for better coverage
        self.frame.bind("<MouseWheel>", _on_mousewheel)
        
        # Store the mousewheel function for binding to child widgets
        self._mousewheel_handler = _on_mousewheel
    
    def _bind_mousewheel_to_all_children(self):
        """Recursively bind mousewheel to all child widgets for better scrolling coverage"""
        if hasattr(self, '_mousewheel_handler'):
            self._bind_mousewheel_to_widget(self.frame)
    
    def _bind_mousewheel_to_widget(self, widget):
        """Bind mousewheel event to a specific widget and all its children"""
        if hasattr(self, '_mousewheel_handler'):
            try:
                widget.bind("<MouseWheel>", self._mousewheel_handler)
            except:
                pass  # Some widgets might not support binding
            
            # Recursively bind to all children
            for child in widget.winfo_children():
                self._bind_mousewheel_to_widget(child)
    
    def _on_tls_change(self):
        """Handle TLS checkbox change - make it mutually exclusive with SSL"""
        if self.use_tls_var.get():
            # If TLS is checked, uncheck SSL
            self.use_ssl_var.set(False)
    
    def _on_ssl_change(self):
        """Handle SSL checkbox change - make it mutually exclusive with TLS"""
        if self.use_ssl_var.get():
            # If SSL is checked, uncheck TLS
            self.use_tls_var.set(False)
    
    def _create_smtp_config_section(self):
        """Create SMTP configuration section"""
        smtp_frame = tk.Frame(self.scrollable_frame, bg=Constants.PROCESSING_BG, relief=tk.RIDGE, bd=1)
        smtp_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Section title
        title_label = tk.Label(smtp_frame, text="🔧 SMTP Configuration", 
                              font=Constants.LABEL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT)
        title_label.pack(pady=(15, 10))
        
        # Configuration grid
        config_grid = tk.Frame(smtp_frame, bg=Constants.PROCESSING_BG)
        config_grid.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        # SMTP Server
        tk.Label(config_grid, text="SMTP Server:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.smtp_server_var = tk.StringVar()
        smtp_entry = tk.Entry(config_grid, textvariable=self.smtp_server_var, width=30,
                             font=Constants.SMALL_FONT)
        smtp_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # SMTP Port
        tk.Label(config_grid, text="Port:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=0, column=2, sticky=tk.W, padx=(0, 10), pady=5)
        self.smtp_port_var = tk.StringVar()
        port_entry = tk.Entry(config_grid, textvariable=self.smtp_port_var, width=10,
                            font=Constants.SMALL_FONT)
        port_entry.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # Email Address
        tk.Label(config_grid, text="Email Address:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.email_address_var = tk.StringVar()
        email_entry = tk.Entry(config_grid, textvariable=self.email_address_var, width=30,
                             font=Constants.SMALL_FONT)
        email_entry.grid(row=1, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # Email Password
        tk.Label(config_grid, text="Password:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=1, column=2, sticky=tk.W, padx=(0, 10), pady=5)
        self.email_password_var = tk.StringVar()
        password_entry = tk.Entry(config_grid, textvariable=self.email_password_var, width=20,
                                font=Constants.SMALL_FONT, show='*')
        password_entry.grid(row=1, column=3, sticky=tk.W, pady=5)
        
        # Security options
        security_frame = tk.Frame(smtp_frame, bg=Constants.PROCESSING_BG)
        security_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        self.use_tls_var = tk.BooleanVar(value=True)
        tls_check = tk.Checkbutton(security_frame, text="Use TLS", variable=self.use_tls_var,
                                 font=Constants.SMALL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT,
                                 command=self._on_tls_change)
        tls_check.pack(side=tk.LEFT, padx=(0, 20))
        
        self.use_ssl_var = tk.BooleanVar(value=False)
        ssl_check = tk.Checkbutton(security_frame, text="Use SSL", variable=self.use_ssl_var,
                                font=Constants.SMALL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT,
                                command=self._on_ssl_change)
        ssl_check.pack(side=tk.LEFT, padx=(0, 20))
        
        # Test connection button
        test_btn = tk.Button(security_frame, text="🔍 Test Connection", 
                           command=self._test_connection, bg=Constants.SECONDARY_BTN, fg='white',
                           font=Constants.BUTTON_FONT, relief=tk.FLAT, padx=15, pady=5)
        test_btn.pack(side=tk.RIGHT)
    
    def _create_email_composition_section(self):
        """Create email composition section"""
        comp_frame = tk.Frame(self.scrollable_frame, bg=Constants.PROCESSING_BG, relief=tk.RIDGE, bd=1)
        comp_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Section title
        title_label = tk.Label(comp_frame, text="✉️ Email Composition", 
                              font=Constants.LABEL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT)
        title_label.pack(pady=(15, 10))
        
        # Composition fields
        comp_grid = tk.Frame(comp_frame, bg=Constants.PROCESSING_BG)
        comp_grid.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        # To field
        tk.Label(comp_grid, text="To:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.to_var = tk.StringVar()
        to_entry = tk.Entry(comp_grid, textvariable=self.to_var, width=70,
                           font=Constants.SMALL_FONT)
        to_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # CC field
        tk.Label(comp_grid, text="CC:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.cc_var = tk.StringVar()
        cc_entry = tk.Entry(comp_grid, textvariable=self.cc_var, width=70,
                          font=Constants.SMALL_FONT)
        cc_entry.grid(row=1, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # BCC field
        tk.Label(comp_grid, text="BCC:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.bcc_var = tk.StringVar()
        bcc_entry = tk.Entry(comp_grid, textvariable=self.bcc_var, width=70,
                           font=Constants.SMALL_FONT)
        bcc_entry.grid(row=2, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # Subject field
        tk.Label(comp_grid, text="Subject:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.subject_var = tk.StringVar()
        subject_entry = tk.Entry(comp_grid, textvariable=self.subject_var, width=70,
                               font=Constants.SMALL_FONT)
        subject_entry.grid(row=3, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # Body field
        tk.Label(comp_grid, text="Body:", font=Constants.LABEL_FONT,
                bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT).grid(row=4, column=0, sticky=tk.NW, padx=(0, 10), pady=5)
        self.body_text = tk.Text(comp_grid, width=70, height=10, font=Constants.SMALL_FONT,
                               wrap=tk.WORD)
        self.body_text.grid(row=4, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        
        # Body scrollbar
        body_scrollbar = ttk.Scrollbar(comp_grid, orient="vertical", command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=body_scrollbar.set)
        body_scrollbar.grid(row=4, column=2, sticky=tk.NS, pady=5)
        
        # Bind mousewheel to text widget for scrolling
        if hasattr(self, '_mousewheel_handler'):
            self.body_text.bind("<MouseWheel>", self._mousewheel_handler)
    
    def _create_attachments_section(self):
        """Create attachments section"""
        attach_frame = tk.Frame(self.scrollable_frame, bg=Constants.PROCESSING_BG, relief=tk.RIDGE, bd=1)
        attach_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Section title
        title_label = tk.Label(attach_frame, text="📎 Attachments", 
                              font=Constants.LABEL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT)
        title_label.pack(pady=(15, 10))
        
        # Attachments list
        attach_content = tk.Frame(attach_frame, bg=Constants.PROCESSING_BG)
        attach_content.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        # Listbox for attachments
        self.attachments_listbox = tk.Listbox(attach_content, height=4, font=Constants.SMALL_FONT)
        self.attachments_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Bind mousewheel to listbox for scrolling
        if hasattr(self, '_mousewheel_handler'):
            self.attachments_listbox.bind("<MouseWheel>", self._mousewheel_handler)
        
        # Attachment buttons
        attach_buttons = tk.Frame(attach_content, bg=Constants.PROCESSING_BG)
        attach_buttons.pack(side=tk.RIGHT)
        
        add_btn = tk.Button(attach_buttons, text="➕ Add", command=self._add_attachment,
                          bg=Constants.PRIMARY_BTN, fg='white', font=Constants.BUTTON_FONT,
                          relief=tk.FLAT, padx=10, pady=3)
        add_btn.pack(pady=(0, 5))
        
        remove_btn = tk.Button(attach_buttons, text="➖ Remove", command=self._remove_attachment,
                             bg='#e74c3c', fg='white', font=Constants.BUTTON_FONT,
                             relief=tk.FLAT, padx=10, pady=3)
        remove_btn.pack()
    
    def _create_action_buttons_section(self):
        """Create action buttons section"""
        action_frame = tk.Frame(self.scrollable_frame, bg=Constants.PROCESSING_BG, relief=tk.RIDGE, bd=1)
        action_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Section title
        title_label = tk.Label(action_frame, text="🚀 Actions", 
                              font=Constants.LABEL_FONT, bg=Constants.PROCESSING_BG, fg=Constants.PRIMARY_TEXT)
        title_label.pack(pady=(15, 10))
        
        # Action buttons
        buttons_frame = tk.Frame(action_frame, bg=Constants.PROCESSING_BG)
        buttons_frame.pack(pady=(0, 15))
        
        # Save config button
        save_btn = tk.Button(buttons_frame, text="💾 Save Configuration", 
                           command=self._save_configuration, bg=Constants.SECONDARY_BTN, fg='white',
                           font=Constants.BUTTON_FONT, relief=tk.FLAT, padx=20, pady=8)
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Send email button
        send_btn = tk.Button(buttons_frame, text="📤 Send Email", 
                           command=self._send_email, bg=Constants.PRIMARY_BTN, fg='white',
                           font=Constants.BUTTON_FONT, relief=tk.FLAT, padx=20, pady=8)
        send_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Load defaults button
        load_btn = tk.Button(buttons_frame, text="🔄 Load Defaults", 
                           command=self._load_defaults, bg='#95a5a6', fg='white',
                           font=Constants.BUTTON_FONT, relief=tk.FLAT, padx=20, pady=8)
        load_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear all button
        clear_btn = tk.Button(buttons_frame, text="🗑️ Clear All", 
                            command=self._clear_all, bg='#e74c3c', fg='white',
                            font=Constants.BUTTON_FONT, relief=tk.FLAT, padx=20, pady=8)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    
    
    
    def _load_configuration(self):
        """Load current email configuration"""
        try:
            config = self.email_sender.get_config()
            
            # SMTP settings
            self.smtp_server_var.set(config.get('smtp_server', ''))
            self.smtp_port_var.set(str(config.get('smtp_port', '')))
            self.email_address_var.set(config.get('email_address', ''))
            self.email_password_var.set(config.get('email_password', ''))
            self.use_tls_var.set(config.get('use_tls', True))
            self.use_ssl_var.set(config.get('use_ssl', False))
            
            # Email composition
            self.to_var.set(config.get('default_to', ''))
            self.cc_var.set(config.get('default_cc', ''))
            self.bcc_var.set(config.get('default_bcc', ''))
            
            # Set dynamic subject with current date and client code placeholder
            current_date = self._get_current_date_formatted()
            default_subject = f"Open Position Report (PS04)_*cp cpde* - {current_date}"
            self.subject_var.set(default_subject)
            
            # Set default body with dynamic date
            default_body = self._get_default_body()
            self.body_text.delete(1.0, tk.END)
            self.body_text.insert(1.0, default_body)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")
    
    def _get_current_date_formatted(self):
        """Get current date in DD MMM YYYY format"""
        from datetime import datetime
        return datetime.now().strftime("%d %b %Y")
    
    def _get_default_body(self):
        """Get default email body with current date"""
        current_date = self._get_current_date_formatted()
        return f"""Dear Sir/Madam,

Please find attached PS04 as at end of day {current_date}.

The password for opening the 'ZIP' file will be your PAN."""
    
    def _save_configuration(self, show_message=True):
        """Save email configuration"""
        try:
            # Update email sender configuration
            self.email_sender.update_config(
                smtp_server=self.smtp_server_var.get(),
                smtp_port=int(self.smtp_port_var.get()) if self.smtp_port_var.get() else 587,
                email_address=self.email_address_var.get(),
                email_password=self.email_password_var.get(),
                default_to=self.to_var.get(),
                default_cc=self.cc_var.get(),
                default_bcc=self.bcc_var.get(),
                default_subject=self.subject_var.get(),
                default_body=self.body_text.get(1.0, tk.END).strip(),
                use_tls=self.use_tls_var.get(),
                use_ssl=self.use_ssl_var.get()
            )
            
            if show_message:
                messagebox.showinfo("Success", "Configuration saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
    
    def _test_connection(self):
        """Test SMTP connection"""
        try:
            # Save current config first
            self._save_configuration(show_message=False)
            
            # Test connection
            success, message = self.email_sender.test_connection()
            
            if success:
                messagebox.showinfo("Connection Test", f"✅ {message}")
            else:
                messagebox.showerror("Connection Test", f"❌ {message}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Connection test failed: {str(e)}")
    
    def _send_email(self):
        """Send email"""
        try:
            # Save current config first
            self._save_configuration(show_message=False)
            
            # Get email content
            to = self.to_var.get().strip()
            subject = self.subject_var.get().strip()
            body = self.body_text.get(1.0, tk.END).strip()
            cc = self.cc_var.get().strip()
            bcc = self.bcc_var.get().strip()
            
            # Validate required fields
            if not to:
                messagebox.showerror("Error", "Please enter recipient email address")
                return
            
            if not subject:
                messagebox.showerror("Error", "Please enter email subject")
                return
            
            if not body:
                messagebox.showerror("Error", "Please enter email body")
                return
            
            # Show loading message
            self._show_loading_message()
            
            # Send email in a separate thread to avoid blocking UI
            import threading
            threading.Thread(target=self._send_email_thread, args=(to, subject, body, cc, bcc), daemon=True).start()
                
        except Exception as e:
            self._hide_loading_message()
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")
    
    def _send_email_thread(self, to, subject, body, cc, bcc):
        """Send email in separate thread"""
        try:
            # Send email
            success, message = self.email_sender.send_email(
                to=to,
                subject=subject,
                body=body,
                cc=cc if cc else None,
                bcc=bcc if bcc else None,
                attachments=self.attachments if self.attachments else None
            )
            
            # Hide loading message and show result
            self.frame.after(0, self._hide_loading_message)
            
            if success:
                self.frame.after(0, lambda: self._show_success_message(message))
            else:
                self.frame.after(0, lambda: messagebox.showerror("Email Failed", f"❌ {message}"))
                
        except Exception as e:
            self.frame.after(0, self._hide_loading_message)
            self.frame.after(0, lambda: messagebox.showerror("Error", f"Failed to send email: {str(e)}"))
    
    def _show_loading_message(self):
        """Show loading message overlay"""
        # Create loading overlay
        self.loading_frame = tk.Frame(self.frame, bg='#A3C39E', relief=tk.RAISED, bd=2)
        self.loading_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Loading content
        loading_content = tk.Frame(self.loading_frame, bg='#ffffff', relief=tk.RAISED, bd=2)
        loading_content.pack(padx=20, pady=20)
        
        # Loading icon and text
        loading_label = tk.Label(loading_content, text="📧 Sending Email...", 
                               font=Constants.LABEL_FONT, bg='#ffffff', fg=Constants.PRIMARY_TEXT)
        loading_label.pack(pady=(15, 10))
        
        # Please wait message
        wait_label = tk.Label(loading_content, text="Please wait while we send your email", 
                            font=Constants.SMALL_FONT, bg='#ffffff', fg=Constants.SECONDARY_TEXT)
        wait_label.pack(pady=(0, 15))
        
        # Disable send button to prevent multiple sends
        self._disable_send_button()
    
    def _hide_loading_message(self):
        """Hide loading message overlay"""
        if hasattr(self, 'loading_frame'):
            self.loading_frame.destroy()
            delattr(self, 'loading_frame')
        
        # Re-enable send button
        self._enable_send_button()
    
    def _show_success_message(self, message):
        """Show success message"""
        messagebox.showinfo("Email Sent Successfully", f"✅ {message}")
    
    def _disable_send_button(self):
        """Disable send button during email sending"""
        # Find and disable send button
        for widget in self.frame.winfo_children():
            self._disable_button_recursive(widget, "📤 Send Email")
    
    def _enable_send_button(self):
        """Enable send button after email sending"""
        # Find and enable send button
        for widget in self.frame.winfo_children():
            self._enable_button_recursive(widget, "📤 Send Email")
    
    def _disable_button_recursive(self, widget, button_text):
        """Recursively find and disable button with specific text"""
        if isinstance(widget, tk.Button) and widget.cget('text') == button_text:
            widget.config(state='disabled')
        for child in widget.winfo_children():
            self._disable_button_recursive(child, button_text)
    
    def _enable_button_recursive(self, widget, button_text):
        """Recursively find and enable button with specific text"""
        if isinstance(widget, tk.Button) and widget.cget('text') == button_text:
            widget.config(state='normal')
        for child in widget.winfo_children():
            self._enable_button_recursive(child, button_text)
    
    def _add_attachment(self):
        """Add file attachment"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select file to attach",
                filetypes=[
                    ("All files", "*.*"),
                    ("PDF files", "*.pdf"),
                    ("Word documents", "*.doc;*.docx"),
                    ("Excel files", "*.xls;*.xlsx"),
                    ("Images", "*.jpg;*.jpeg;*.png;*.gif"),
                    ("Text files", "*.txt")
                ]
            )
            
            if file_path:
                self.attachments.append(file_path)
                filename = os.path.basename(file_path)
                self.attachments_listbox.insert(tk.END, filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add attachment: {str(e)}")
    
    def _remove_attachment(self):
        """Remove selected attachment"""
        try:
            selection = self.attachments_listbox.curselection()
            if selection:
                index = selection[0]
                self.attachments_listbox.delete(index)
                self.attachments.pop(index)
            else:
                messagebox.showwarning("Warning", "Please select an attachment to remove")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove attachment: {str(e)}")
    
    def _load_defaults(self):
        """Load default configuration"""
        try:
            # Reset to default values
            self.smtp_server_var.set("smtp.gmail.com")
            self.smtp_port_var.set("587")
            self.email_address_var.set("")
            self.email_password_var.set("")
            self.use_tls_var.set(True)
            self.use_ssl_var.set(False)
            
            self.to_var.set("")
            self.cc_var.set("")
            self.bcc_var.set("")
            self.subject_var.set("")
            self.body_text.delete(1.0, tk.END)
            
            # Clear attachments
            self.attachments_listbox.delete(0, tk.END)
            self.attachments.clear()
            
            messagebox.showinfo("Success", "Default configuration loaded!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load defaults: {str(e)}")
    
    def _clear_all(self):
        """Clear all fields"""
        try:
            if messagebox.askyesno("Confirm", "Are you sure you want to clear all fields?"):
                self.smtp_server_var.set("")
                self.smtp_port_var.set("")
                self.email_address_var.set("")
                self.email_password_var.set("")
                
                self.to_var.set("")
                self.cc_var.set("")
                self.bcc_var.set("")
                self.subject_var.set("")
                self.body_text.delete(1.0, tk.END)
                
                # Clear attachments
                self.attachments_listbox.delete(0, tk.END)
                self.attachments.clear()
                
                messagebox.showinfo("Success", "All fields cleared!")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear fields: {str(e)}")
    
    def pack(self, **kwargs):
        """Pack the frame"""
        self.frame.pack(**kwargs)
    
    def pack_forget(self):
        """Pack forget the frame"""
        self.frame.pack_forget()
    
    def get_values(self):
        """Get current values for external use"""
        return {
            'smtp_server': self.smtp_server_var.get(),
            'smtp_port': self.smtp_port_var.get(),
            'email_address': self.email_address_var.get(),
            'email_password': self.email_password_var.get(),
            'to': self.to_var.get(),
            'cc': self.cc_var.get(),
            'bcc': self.bcc_var.get(),
            'subject': self.subject_var.get(),
            'body': self.body_text.get(1.0, tk.END).strip(),
            'attachments': self.attachments.copy()
        }
