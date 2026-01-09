import win32com.client as win32
import os


def send_outlook_email(recipient, subject, body, attachment_paths=None):
    """
    Send an email via Outlook using win32com.client.
    
    Args:
        recipient (str): Email recipient (can be comma-separated for multiple recipients)
        subject (str): Email subject line
        body (str): Email body (HTML or plain text)
        attachment_paths (list or str, optional): Single file path (str) or list of file paths to attach
    
    Returns:
        bool: True if email was sent successfully, False otherwise
    """
    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set recipient (handle comma-separated emails)
        if isinstance(recipient, str):
            mail.To = recipient
        else:
            mail.To = ';'.join(recipient) if isinstance(recipient, list) else str(recipient)
        
        mail.Subject = subject
        mail.HTMLBody = body if body else ""
        
        # Handle attachments - can be single path (str) or list of paths
        if attachment_paths:
            if isinstance(attachment_paths, str):
                attachment_paths = [attachment_paths]
            
            for attachment_path in attachment_paths:
                if os.path.exists(attachment_path):
                    mail.Attachments.Add(attachment_path)
                else:
                    print(f"Attachment file not found: {attachment_path}")
                    return False
        
        # Try to send email directly, but fallback to Display if security blocks it
        try:
            mail.Send()
            return True
        except Exception as send_error:
            # Check if this is the "Operation aborted" error (Outlook security blocking)
            error_args = getattr(send_error, 'args', [])
            error_code = error_args[0] if error_args else None
            
            # Error code -2147467260 is E_ABORT (Operation aborted)
            if error_code == -2147467260 or "aborted" in str(send_error).lower():
                # Outlook security is blocking Send(), use Display() instead
                try:
                    mail.Display(False)  # False = non-modal
                    print("Email opened in Outlook. Please review and send manually (Outlook security blocked automatic send).")
                    return True
                except Exception as display_error:
                    print(f"Failed to display email: {display_error}")
                    raise send_error  # Re-raise original error if Display also fails
            else:
                # Different error, re-raise it
                raise
        
    except Exception as e:
        print(f"An error occurred while sending email: {e}")
        return False

