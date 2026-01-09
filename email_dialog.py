import tkinter as tk
from tkinter import messagebox
import tempfile
import zipfile
import os
import sys
from outlook_email import send_outlook_email


class EmailDialog(tk.Toplevel):
    """Dialog window for sending email with file attachments via Outlook."""
    
    def __init__(self, parent, zip_path, file_list, default_subject=None, default_body=None, default_email=None):
        """
        Initialize the email dialog.
        
        Args:
            parent: Parent window
            zip_path (str): Path to the ZIP file containing the files
            file_list (list): List of filenames in the ZIP
            default_subject (str, optional): Default subject for the email
            default_body (str, optional): Default body text for the email
            default_email (str, optional): Default recipient email address
        """
        super().__init__(parent)
        self.parent = parent
        self.zip_path = zip_path
        self.file_list = file_list
        self.temp_dir = None
        
        self.title("Send Email via Outlook")
        self.geometry("500x500")
        self.configure(bg="#ecf0f1")
        self.resizable(True, True)
        self.minsize(500, 450)
        
        # Load icon
        self._load_icon()
        
        # Center the dialog on screen
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Default values - use provided values or fall back to defaults
        self.default_subject = default_subject if default_subject is not None else "Daily_Trade_File"
        self.default_email = default_email if default_email is not None else ""
        self.default_body = default_body if default_body is not None else ""
        
        # Store checkboxes for file selection (initialize before _create_widgets)
        self.file_vars = {}
        
        # Create UI
        self._create_widgets()
        
        # Set focus to canvas after widgets are created so arrow keys work immediately
        self.after(100, self._focus_attachments_box)
        
        # Bind arrow keys to dialog window so they work globally
        self.bind("<Up>", self._on_dialog_arrow_up)
        self.bind("<Down>", self._on_dialog_arrow_down)
        
        # Set focus to canvas after widgets are created so arrow keys work immediately
        self.after(100, self._focus_attachments_box)
    
    def _create_widgets(self):
        """Create UI widgets for the email dialog."""
        # Title
        title_label = tk.Label(
            self, 
            text="Send Email via Outlook", 
            font=("Arial", 14, "bold"), 
            bg="#ecf0f1", 
            fg="#2c3e50"
        )
        title_label.pack(pady=8)
        
        # Subject field
        subject_frame = tk.Frame(self, bg="#ecf0f1")
        subject_frame.pack(fill="x", padx=15, pady=4)
        tk.Label(
            subject_frame, 
            text="Subject:", 
            font=("Arial", 10), 
            bg="#ecf0f1", 
            fg="#2c3e50",
            width=10,
            anchor="w"
        ).pack(side="left")
        self.subject_var = tk.StringVar(value=self.default_subject)
        subject_entry = tk.Entry(subject_frame, textvariable=self.subject_var, width=40, font=("Arial", 9))
        subject_entry.pack(side="left", padx=5)
        
        # Email field - supports multiple emails (comma or semicolon separated)
        # Users can copy-paste multiple email addresses in this field
        email_frame = tk.Frame(self, bg="#ecf0f1")
        email_frame.pack(fill="x", padx=15, pady=4)
        tk.Label(
            email_frame, 
            text="Email To:", 
            font=("Arial", 10), 
            bg="#ecf0f1", 
            fg="#2c3e50",
            width=10,
            anchor="w"
        ).pack(side="left")
        self.email_var = tk.StringVar(value=self.default_email)
        email_entry = tk.Entry(email_frame, textvariable=self.email_var, width=40, font=("Arial", 9))
        email_entry.pack(side="left", padx=5)
        
        # Helper label for multiple emails (small text below the entry)
        email_help_frame = tk.Frame(self, bg="#ecf0f1")
        email_help_frame.pack(fill="x", padx=15, pady=(0, 4))
        tk.Label(
            email_help_frame,
            text="Note: Multiple emails can be pasted here (comma or semicolon separated)",
            font=("Arial", 8),
            bg="#ecf0f1",
            fg="#7f8c8d",
            anchor="w"
        ).pack(side="left", padx=(85, 0))  # Align with email entry field
        
        # Body field - compact
        body_frame = tk.Frame(self, bg="#ecf0f1")
        body_frame.pack(fill="x", padx=15, pady=4)
        tk.Label(
            body_frame, 
            text="Body:", 
            font=("Arial", 10), 
            bg="#ecf0f1", 
            fg="#2c3e50",
            width=10,
            anchor="nw"
        ).pack(side="left", anchor="n", pady=(3, 0))
        
        body_text_frame = tk.Frame(body_frame, bg="#ffffff", highlightthickness=1, highlightbackground="#bdc3c7", height=50)
        body_text_frame.pack(side="left", fill="both", expand=True, padx=5)
        body_text_frame.pack_propagate(False)
        
        self.body_text = tk.Text(
            body_text_frame,
            width=40,
            height=2,
            font=("Arial", 9),
            wrap=tk.WORD,
            bg="#ffffff",
            fg="#2c3e50",
            relief="flat",
            padx=3,
            pady=3
        )
        self.body_text.pack(fill="both", expand=True)
        
        # Set default body text if provided
        if self.default_body:
            self.body_text.insert("1.0", self.default_body)
        
        # Email sending option checkbox
        email_option_frame = tk.Frame(self, bg="#ecf0f1")
        email_option_frame.pack(fill="x", padx=15, pady=4)
        self.separate_emails_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            email_option_frame,
            text="Send each file in a separate email",
            variable=self.separate_emails_var,
            font=("Arial", 9),
            bg="#ecf0f1",
            fg="#2c3e50",
            selectcolor="#ecf0f1"
        ).pack(side="left")
        
        # File selection label
        file_label_frame = tk.Frame(self, bg="#ecf0f1")
        file_label_frame.pack(fill="x", padx=15, pady=(6, 3))
        tk.Label(
            file_label_frame, 
            text="Attachments:", 
            font=("Arial", 10, "bold"), 
            bg="#ecf0f1", 
            fg="#2c3e50"
        ).pack(side="left")
        
        # File selection frame with scrollbar - fixed height
        file_selection_frame = tk.Frame(self, bg="#ecf0f1", height=150)
        file_selection_frame.pack(fill="x", padx=15, pady=3)
        file_selection_frame.pack_propagate(False)
        
        # Create canvas and scrollbar for file list - cursor set directly on widgets (simple approach like dashboard.py)
        canvas = tk.Canvas(file_selection_frame, bg="#ffffff", highlightthickness=1, highlightbackground="#bdc3c7", height=150, cursor="hand2")
        scrollbar = tk.Scrollbar(file_selection_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#ffffff", cursor="hand2")  # Also set cursor directly on frame (like dashboard.py)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Enable keyboard arrow key scrolling
        def _on_arrow_up(event):
            canvas.yview_scroll(-1, "units")
            return "break"  # Prevent default behavior
        
        def _on_arrow_down(event):
            canvas.yview_scroll(1, "units")
            return "break"  # Prevent default behavior
        
        # Make canvas focusable and bind arrow keys
        canvas.config(takefocus=True)
        canvas.bind("<Up>", _on_arrow_up)
        canvas.bind("<Down>", _on_arrow_down)
        canvas.bind("<Button-1>", lambda e: canvas.focus_set())  # Focus on click
        
        # Store handlers and canvas reference for use in other bindings
        self._arrow_up_handler = _on_arrow_up
        self._arrow_down_handler = _on_arrow_down
        self.attachments_canvas = canvas  # Store canvas reference for focus and global bindings
        
        # Enable mouse wheel scrolling and cursor on canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            # CRITICAL: Always set cursor to hand2 when scrolling (works even in empty space)
            canvas.config(cursor="hand2")
        
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.config(cursor="hand2")  # Set cursor when entering canvas area
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            # Don't reset cursor immediately - check if mouse is still in canvas area
            # The cursor should stay hand2 as long as mouse is anywhere in attachments area
        
        canvas.bind("<Enter>", _bind_to_mousewheel)
        canvas.bind("<Leave>", _unbind_from_mousewheel)
        
        # Bind mouse motion on canvas to ensure cursor stays
        def _on_canvas_motion(event):
            canvas.config(cursor="hand2")
        
        canvas.bind("<Motion>", _on_canvas_motion)
        
        # Bind mouse motion on scrollable frame to change canvas cursor - critical for empty spaces
        def _on_frame_motion(event):
            # Always set canvas cursor when mouse moves over frame (even empty space)
            canvas.config(cursor="hand2")
        
        # Also bind to frame enter/leave to maintain cursor
        def _on_frame_enter(event):
            canvas.config(cursor="hand2")
        
        def _on_frame_leave(event):
            # When leaving frame, check if we're still in canvas area
            # If yes, keep cursor; if no, reset it
            try:
                x, y = self.winfo_pointerx(), self.winfo_pointery()
                widget_x = canvas.winfo_rootx()
                widget_y = canvas.winfo_rooty()
                widget_width = canvas.winfo_width()
                widget_height = canvas.winfo_height()
                
                # If mouse is still within canvas bounds, keep cursor
                if widget_x <= x <= widget_x + widget_width and widget_y <= y <= widget_y + widget_height:
                    canvas.config(cursor="hand2")
                else:
                    canvas.config(cursor="")
            except:
                # If calculation fails, keep cursor as hand2 to be safe
                canvas.config(cursor="hand2")
        
        # Aggressively bind to frame - these should fire for ANY mouse movement in frame
        scrollable_frame.bind("<Motion>", _on_frame_motion)  # Motion events on frame (empty space)
        scrollable_frame.bind("<Enter>", _on_frame_enter)
        scrollable_frame.bind("<Leave>", _on_frame_leave)
        
        # Also bind to all mouse events on frame to ensure cursor is always hand2
        scrollable_frame.bind("<Button-1>", lambda e: canvas.config(cursor="hand2"))
        scrollable_frame.bind("<ButtonRelease-1>", lambda e: canvas.config(cursor="hand2"))
        
        # Bind arrow keys to scrollable_frame as well (so they work when frame has focus)
        scrollable_frame.bind("<Up>", lambda e: self._arrow_up_handler(e))
        scrollable_frame.bind("<Down>", lambda e: self._arrow_down_handler(e))
        scrollable_frame.bind("<Button-1>", lambda e: canvas.focus_set())  # Focus canvas on click
        
        # Add checkboxes for each file
        for filename in self.file_list:
            var = tk.BooleanVar(value=True)  # Default: all files selected
            self.file_vars[filename] = var
            checkbox = tk.Checkbutton(
                scrollable_frame,
                text=filename,
                variable=var,
                font=("Arial", 9),
                bg="#ffffff",
                fg="#2c3e50",
                selectcolor="#ffffff",
                anchor="w",
                padx=5,
                pady=2,
                cursor="hand2"  # Set cursor directly on checkbox (works for both box and text)
            )
            checkbox.pack(fill="x", padx=8, pady=1)
            
            # Bind ALL events to ensure canvas cursor is ALWAYS hand2 when in attachments area
            def _on_checkbox_motion(event):
                # Always set canvas cursor when mouse moves over checkbox
                canvas.config(cursor="hand2")
            
            def _on_checkbox_enter(event):
                canvas.config(cursor="hand2")
            
            def _on_checkbox_leave(event):
                # When leaving checkbox, immediately set canvas cursor to hand2
                # The frame motion handler will also fire, ensuring cursor stays hand2
                canvas.config(cursor="hand2")
                # Also trigger frame motion to ensure cursor is set even if moving to empty space
                try:
                    x, y = self.winfo_pointerx(), self.winfo_pointery()
                    # Convert to canvas coordinates to check if still in canvas
                    canvas_x = canvas.winfo_rootx()
                    canvas_y = canvas.winfo_rooty()
                    if canvas_x <= x <= canvas_x + canvas.winfo_width() and canvas_y <= y <= canvas_y + canvas.winfo_height():
                        canvas.config(cursor="hand2")
                except:
                    canvas.config(cursor="hand2")
            
            # Bind to checkbox with aggressive cursor setting - use add="+" to not override
            checkbox.bind("<Motion>", _on_checkbox_motion, add="+")
            checkbox.bind("<Enter>", _on_checkbox_enter, add="+")
            checkbox.bind("<Leave>", _on_checkbox_leave, add="+")
            # Also bind to Button-1, ButtonRelease-1, etc. to maintain cursor
            checkbox.bind("<Button-1>", lambda e: canvas.config(cursor="hand2"))
            checkbox.bind("<ButtonRelease-1>", lambda e: canvas.config(cursor="hand2"))
            # Bind arrow keys to checkboxes for scrolling
            checkbox.bind("<Up>", lambda e: self._arrow_up_handler(e))
            checkbox.bind("<Down>", lambda e: self._arrow_down_handler(e))
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Buttons frame - ensure it's always visible at bottom
        buttons_frame = tk.Frame(self, bg="#ecf0f1", height=60)
        buttons_frame.pack(side="bottom", fill="x", padx=15, pady=(8, 12))
        buttons_frame.pack_propagate(False)
        
        # Cancel button
        cancel_btn = tk.Button(
            buttons_frame,
            text="Cancel",
            command=self._on_cancel,
            bg="#95a5a6",
            fg="white",
            relief="flat",
            padx=18,
            pady=6,
            font=("Arial", 10, "bold"),
            cursor="hand2"
        )
        cancel_btn.pack(side="right", padx=(8, 0))
        
        # Send Mail button
        send_btn = tk.Button(
            buttons_frame,
            text="Send Mail",
            command=self._on_send,
            bg="#27ae60",
            fg="white",
            relief="flat",
            padx=18,
            pady=6,
            font=("Arial", 10, "bold"),
            cursor="hand2"
        )
        send_btn.pack(side="right", padx=(0, 8))
    
    def _focus_attachments_box(self):
        """Set focus to attachments canvas so arrow keys work immediately."""
        if hasattr(self, 'attachments_canvas'):
            try:
                self.attachments_canvas.focus_set()
            except:
                pass
    
    def _on_dialog_arrow_up(self, event):
        """Handle arrow up key when pressed anywhere in dialog (except text fields)."""
        # Only scroll if not in a text entry field (subject, email, body)
        focused_widget = self.focus_get()
        if focused_widget and isinstance(focused_widget, (tk.Entry, tk.Text)):
            # Let text widgets handle arrow keys normally
            return None
        # Otherwise, scroll the attachments
        if hasattr(self, '_arrow_up_handler'):
            return self._arrow_up_handler(event)
        return None
    
    def _on_dialog_arrow_down(self, event):
        """Handle arrow down key when pressed anywhere in dialog (except text fields)."""
        # Only scroll if not in a text entry field (subject, email, body)
        focused_widget = self.focus_get()
        if focused_widget and isinstance(focused_widget, (tk.Entry, tk.Text)):
            # Let text widgets handle arrow keys normally
            return None
        # Otherwise, scroll the attachments
        if hasattr(self, '_arrow_down_handler'):
            return self._arrow_down_handler(event)
        return None
    
    def _on_cancel(self):
        """Handle cancel button click."""
        self._cleanup_temp_dir()
        self.destroy()
    
    def _cleanup_temp_dir(self):
        """Clean up temporary directory."""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
            except:
                pass
    
    def _extract_files_from_zip(self, selected_filenames, output_dir):
        """
        Extract specific files from a ZIP archive.
        
        Args:
            selected_filenames (list): List of filenames to extract
            output_dir (str): Directory to extract files to
        
        Returns:
            list: List of paths to extracted files, or empty list on error
        """
        extracted_paths = []
        
        try:
            os.makedirs(output_dir, exist_ok=True)
            
            # Extract selected files from ZIP
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_file_list = zip_ref.namelist()
                
                # Extract only selected files
                for filename in selected_filenames:
                    if filename in zip_file_list:
                        extracted_path = os.path.join(output_dir, filename)
                        zip_ref.extract(filename, output_dir)
                        extracted_paths.append(extracted_path)
                    else:
                        messagebox.showwarning("Warning", f"File '{filename}' not found in ZIP archive.")
            
            return extracted_paths
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract files from ZIP: {str(e)}")
            return []
    
    def _parse_emails(self, email_text):
        """
        Parse email addresses from text input.
        Supports multiple formats:
        - One email per line
        - Comma-separated emails
        - Semicolon-separated emails
        - Mixed formats
        
        Args:
            email_text (str): Raw email input text
            
        Returns:
            list: List of valid email addresses (stripped and cleaned)
        """
        if not email_text:
            return []
        
        # Replace semicolons with commas for consistent parsing
        email_text = email_text.replace(";", ",")
        
        # Split by both commas and newlines
        emails = []
        for line in email_text.split("\n"):
            for email in line.split(","):
                email = email.strip()
                if email:  # Only add non-empty emails
                    emails.append(email)
        
        # Remove duplicates while preserving order
        seen = set()
        unique_emails = []
        for email in emails:
            if email not in seen:
                seen.add(email)
                unique_emails.append(email)
        
        return unique_emails
    
    def _validate_email(self, email):
        """
        Basic email validation.
        
        Args:
            email (str): Email address to validate
            
        Returns:
            bool: True if email appears valid, False otherwise
        """
        if not email or not email.strip():
            return False
        
        email = email.strip()
        
        # Basic validation: must contain @ and have a domain with at least one dot
        if "@" not in email:
            return False
        
        parts = email.split("@")
        if len(parts) != 2:
            return False
        
        local, domain = parts
        if not local or not domain:
            return False
        
        # Domain must contain at least one dot
        if "." not in domain:
            return False
        
        return True
    
    def _on_send(self):
        """Handle send mail button click."""
        # Get subject and email
        subject = self.subject_var.get().strip()
        email_text = self.email_var.get().strip()
        
        # Validate inputs
        if not subject:
            messagebox.showwarning("Validation Error", "Please enter a subject.")
            return
        
        if not email_text:
            messagebox.showwarning("Validation Error", "Please enter at least one email address.")
            return
        
        # Parse multiple emails
        email_list = self._parse_emails(email_text)
        
        if not email_list:
            messagebox.showwarning("Validation Error", "Please enter at least one valid email address.")
            return
        
        # Validate all emails
        invalid_emails = []
        for email in email_list:
            if not self._validate_email(email):
                invalid_emails.append(email)
        
        if invalid_emails:
            messagebox.showwarning(
                "Validation Error", 
                f"Please enter valid email addresses.\n\nInvalid emails:\n" + "\n".join(invalid_emails)
            )
            return
        
        # Join emails with semicolon for Outlook (Outlook uses semicolon as separator)
        email = "; ".join(email_list)
        
        # Get selected files
        selected_files = [filename for filename, var in self.file_vars.items() if var.get()]
        
        if not selected_files:
            messagebox.showwarning("Validation Error", "Please select at least one file to send.")
            return
        
        # Extract selected files from ZIP to temporary directory
        try:
            self.temp_dir = tempfile.mkdtemp()
            extracted_paths = self._extract_files_from_zip(selected_files, self.temp_dir)
            
            if not extracted_paths:
                self._cleanup_temp_dir()
                messagebox.showerror("Error", "Failed to extract files from ZIP.")
                return
            
            # Get body text from text widget
            body = self.body_text.get("1.0", tk.END).strip()
            
            # Check if sending separate emails or one email
            send_separate = self.separate_emails_var.get()
            
            if send_separate:
                # Send each file in a separate email
                success_count = 0
                failed_count = 0
                
                for i, file_path in enumerate(extracted_paths):
                    # Use original subject as entered (no filename modification)
                    success = send_outlook_email(
                        recipient=email,
                        subject=subject,
                        body=body,
                        attachment_paths=[file_path]  # Single file per email
                    )
                    
                    if success:
                        success_count += 1
                    else:
                        failed_count += 1
                
                # Show summary
                if failed_count == 0:
                    messagebox.showinfo(
                        "Success",
                        f"Successfully processed {success_count} email(s).\n\n"
                        f"Each file was sent in a separate email.\n"
                        f"If Outlook opened the emails, please review and send them manually."
                    )
                    self.destroy()
                else:
                    messagebox.showwarning(
                        "Partial Success",
                        f"Processed {success_count} email(s) successfully.\n"
                        f"Failed to process {failed_count} email(s).\n\n"
                        f"Please check Outlook configuration."
                    )
                    self._cleanup_temp_dir()
            else:
                # Send all files in one email
                
                # Send email via Outlook
                success = send_outlook_email(
                    recipient=email,
                    subject=subject,
                    body=body,
                    attachment_paths=extracted_paths
                )
                
                if success:
                    messagebox.showinfo(
                        "Success", 
                        f"Email processed successfully with {len(extracted_paths)} attachment(s).\n\n"
                        f"If Outlook opened the email, please review and send it manually.\n"
                        f"If email was sent automatically, it has been delivered."
                    )
                    # Don't cleanup temp dir immediately - Outlook might need the files
                    # Cleanup will happen when dialog is destroyed
                    self.destroy()
                else:
                    self._cleanup_temp_dir()
                    messagebox.showerror("Error", "Failed to send email. Please check Outlook configuration.")
        
        except Exception as e:
            self._cleanup_temp_dir()
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")
    
    def _get_icon_path(self):
        """Get the icon path for the email dialog - works for both development and compiled EXE"""
        if getattr(sys, 'frozen', False):
            # Running as compiled EXE
            return os.path.join(sys._MEIPASS, "outlook.png")
        else:
            # Running as script - try multiple possible locations
            possible_paths = [
                "outlook.png",  # Same directory (root)
                "../outlook.png",  # Parent directory
                os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "outlook.png")  # Root directory
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return os.path.abspath(path)
            return None
    
    def _load_icon(self):
        """Load outlook.png icon for the dialog"""
        try:
            icon_path = self._get_icon_path()
            if icon_path and os.path.exists(icon_path):
                # Use iconphoto for PNG files (works better than iconbitmap)
                try:
                    from PIL import Image, ImageTk
                    icon_image = Image.open(icon_path)
                    icon_photo = ImageTk.PhotoImage(icon_image)
                    self.iconphoto(False, icon_photo)
                except Exception:
                    pass
        except Exception:
            pass
    
    def destroy(self):
        """Override destroy to clean up temporary directory."""
        self._cleanup_temp_dir()
        super().destroy()

