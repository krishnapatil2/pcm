"""
Email Sender Module for PCM Application
Dynamic email functionality with comprehensive configuration support
"""

import smtplib
import ssl
import json
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication
from datetime import datetime
import traceback


class EmailSender:
    """Dynamic email sender class with comprehensive functionality"""
    
    def __init__(self, config_file="email_config.json"):
        self.config_file = config_file
        self.config = self._load_config()
        self.smtp_server = None
        self.smtp_port = None
        self.email_address = None
        self.email_password = None
        
    def _load_config(self):
        """Load email configuration from JSON file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # Create default configuration
                default_config = {
                    "smtp_server": "smtp.gmail.com",
                    "smtp_port": 587,
                    "email_address": "",
                    "email_password": "",
                    "default_to": "",
                    "default_subject": "",
                    "default_cc": "",
                    "default_bcc": "",
                    "default_body": "",
                    "use_tls": True,
                    "use_ssl": False,
                    "timeout": 30,
                    "max_retries": 3,
                    "retry_delay": 5
                }
                self._save_config(default_config)
                return default_config
        except Exception as e:
            print(f"Error loading email config: {e}")
            return {}
    
    def _save_config(self, config=None):
        """Save email configuration to JSON file"""
        try:
            config_to_save = config or self.config
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving email config: {e}")
    
    def update_config(self, **kwargs):
        """Update email configuration dynamically"""
        try:
            for key, value in kwargs.items():
                if key in self.config:
                    self.config[key] = value
            
            self._save_config()
            return True
        except Exception as e:
            print(f"Error updating config: {e}")
            return False
    
    def get_config(self):
        """Get current email configuration"""
        return self.config.copy()
    
    def validate_config(self):
        """Validate email configuration"""
        required_fields = ['smtp_server', 'smtp_port', 'email_address', 'email_password']
        missing_fields = []
        
        for field in required_fields:
            if not self.config.get(field):
                missing_fields.append(field)
        
        if missing_fields:
            return False, f"Missing required fields: {', '.join(missing_fields)}"
        
        # Validate port number
        try:
            port = int(self.config['smtp_port'])
            if port < 1 or port > 65535:
                return False, "Invalid port number"
        except (ValueError, TypeError):
            return False, "Port must be a valid number"
        
        return True, "Configuration is valid"
    
    def test_connection(self):
        """Test SMTP connection with current configuration"""
        try:
            is_valid, message = self.validate_config()
            if not is_valid:
                return False, message
            
            # Create SMTP connection
            if self.config.get('use_ssl', False):
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(
                    self.config['smtp_server'], 
                    int(self.config['smtp_port']), 
                    timeout=self.config.get('timeout', 30),
                    context=context
                )
            else:
                server = smtplib.SMTP(
                    self.config['smtp_server'], 
                    int(self.config['smtp_port']), 
                    timeout=self.config.get('timeout', 30)
                )
                
                if self.config.get('use_tls', True):
                    server.starttls()
            
            # Login
            server.login(self.config['email_address'], self.config['email_password'])
            server.quit()
            
            return True, "Connection successful"
            
        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed. Check email and password."
        except smtplib.SMTPConnectError:
            return False, "Connection failed. Check SMTP server and port."
        except smtplib.SMTPException as e:
            return False, f"SMTP error: {str(e)}"
        except Exception as e:
            return False, f"Connection test failed: {str(e)}"
    
    def send_email(self, to=None, subject=None, body=None, cc=None, bcc=None, 
                   attachments=None, html_body=None, reply_to=None):
        """
        Send email with dynamic parameters
        
        Args:
            to (str or list): Recipient email addresses
            subject (str): Email subject
            body (str): Plain text body
            cc (str or list): CC recipients
            bcc (str or list): BCC recipients
            attachments (list): List of file paths to attach
            html_body (str): HTML body content
            reply_to (str): Reply-to email address
        
        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            # Validate configuration
            is_valid, message = self.validate_config()
            if not is_valid:
                return False, message
            
            # Use defaults if not provided
            to = to or self.config.get('default_to', '')
            subject = subject or self.config.get('default_subject', '')
            body = body or self.config.get('default_body', '')
            cc = cc or self.config.get('default_cc', '')
            bcc = bcc or self.config.get('default_bcc', '')
            
            # Validate required fields
            if not to:
                return False, "Recipient email address is required"
            if not subject:
                return False, "Email subject is required"
            
            # Convert string recipients to list
            if isinstance(to, str):
                to = [email.strip() for email in to.split(',') if email.strip()]
            if isinstance(cc, str) and cc:
                cc = [email.strip() for email in cc.split(',') if email.strip()]
            if isinstance(bcc, str) and bcc:
                bcc = [email.strip() for email in bcc.split(',') if email.strip()]
            
            # Create message
            msg = MIMEMultipart('alternative')
            msg['From'] = self.config['email_address']
            msg['To'] = ', '.join(to)
            msg['Subject'] = subject
            
            if cc:
                msg['Cc'] = ', '.join(cc)
            if bcc:
                msg['Bcc'] = ', '.join(bcc)
            if reply_to:
                msg['Reply-To'] = reply_to
            
            # Add body content
            if html_body:
                # Create HTML part
                html_part = MIMEText(html_body, 'html')
                msg.attach(html_part)
                
                # Add plain text version if provided
                if body:
                    text_part = MIMEText(body, 'plain')
                    msg.attach(text_part)
            else:
                # Plain text only
                text_part = MIMEText(body, 'plain')
                msg.attach(text_part)
            
            # Add attachments
            if attachments:
                for attachment_path in attachments:
                    if os.path.exists(attachment_path):
                        self._add_attachment(msg, attachment_path)
                    else:
                        return False, f"Attachment file not found: {attachment_path}"
            
            # Send email
            return self._send_message(msg, to + (cc or []) + (bcc or []))
            
        except Exception as e:
            error_msg = f"Failed to send email: {str(e)}"
            print(f"Email error: {error_msg}")
            print(f"Traceback: {traceback.format_exc()}")
            return False, error_msg
    
    def _add_attachment(self, msg, file_path):
        """Add file attachment to email message"""
        try:
            with open(file_path, 'rb') as attachment:
                # Determine file type
                if file_path.lower().endswith(('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx')):
                    # Use MIMEApplication for binary files
                    part = MIMEApplication(attachment.read())
                else:
                    # Use MIMEBase for other files
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                
                # Set filename
                filename = os.path.basename(file_path)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {filename}'
                )
                msg.attach(part)
                
        except Exception as e:
            raise Exception(f"Failed to add attachment {file_path}: {str(e)}")
    
    def _send_message(self, msg, recipients):
        """Send the email message"""
        max_retries = self.config.get('max_retries', 3)
        retry_delay = self.config.get('retry_delay', 5)
        
        for attempt in range(max_retries):
            try:
                # Create SMTP connection
                if self.config.get('use_ssl', False):
                    context = ssl.create_default_context()
                    server = smtplib.SMTP_SSL(
                        self.config['smtp_server'], 
                        int(self.config['smtp_port']), 
                        timeout=self.config.get('timeout', 30),
                        context=context
                    )
                else:
                    server = smtplib.SMTP(
                        self.config['smtp_server'], 
                        int(self.config['smtp_port']), 
                        timeout=self.config.get('timeout', 30)
                    )
                    
                    if self.config.get('use_tls', True):
                        server.starttls()
                
                # Login and send
                server.login(self.config['email_address'], self.config['email_password'])
                server.send_message(msg, to_addrs=recipients)
                server.quit()
                
                return True, f"Email sent successfully to {len(recipients)} recipient(s)"
                
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"Attempt {attempt + 1} failed: {str(e)}. Retrying in {retry_delay} seconds...")
                    import time
                    time.sleep(retry_delay)
                else:
                    raise e
    
    def send_template_email(self, template_name, **kwargs):
        """Send email using predefined template"""
        try:
            # Load template
            template_file = f"email_templates/{template_name}.json"
            if not os.path.exists(template_file):
                return False, f"Template '{template_name}' not found"
            
            with open(template_file, 'r', encoding='utf-8') as f:
                template = json.load(f)
            
            # Replace placeholders in template
            subject = template.get('subject', '')
            body = template.get('body', '')
            html_body = template.get('html_body', '')
            
            # Replace placeholders with provided values
            for key, value in kwargs.items():
                placeholder = f"{{{key}}}"
                subject = subject.replace(placeholder, str(value))
                body = body.replace(placeholder, str(value))
                if html_body:
                    html_body = html_body.replace(placeholder, str(value))
            
            # Send email
            return self.send_email(
                to=template.get('to'),
                subject=subject,
                body=body,
                html_body=html_body,
                cc=template.get('cc'),
                bcc=template.get('bcc'),
                attachments=template.get('attachments')
            )
            
        except Exception as e:
            return False, f"Failed to send template email: {str(e)}"
    
    def get_email_status(self):
        """Get current email configuration status"""
        is_valid, message = self.validate_config()
        return {
            'config_valid': is_valid,
            'config_message': message,
            'smtp_server': self.config.get('smtp_server', ''),
            'smtp_port': self.config.get('smtp_port', ''),
            'email_address': self.config.get('email_address', ''),
            'has_password': bool(self.config.get('email_password')),
            'use_tls': self.config.get('use_tls', True),
            'use_ssl': self.config.get('use_ssl', False)
        }


# Utility functions for easy integration
def create_email_sender(config_file="email_config.json"):
    """Create EmailSender instance"""
    return EmailSender(config_file)

def send_quick_email(to, subject, body, config_file="email_config.json", **kwargs):
    """Quick email sending function"""
    sender = EmailSender(config_file)
    return sender.send_email(to=to, subject=subject, body=body, **kwargs)

def test_email_config(config_file="email_config.json"):
    """Test email configuration"""
    sender = EmailSender(config_file)
    return sender.test_connection()
