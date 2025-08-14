import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os
from pathlib import Path
import sys

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config import Config

def send_email_with_attachment(to_address, subject, body, attachment_path=None, cc_address=None, thread_info=None):
    """
    Send email with optional attachment, CC support, and threading.
    
    Args:
        to_address (str): Primary recipient email address
        subject (str): Email subject
        body (str): Email body text
        attachment_path (str, optional): Path to file to attach (None for no attachment)
        cc_address (str, optional): CC recipient email address
        thread_info (dict, optional): Threading information for replies
        
    Raises:
        ValueError: If required configuration is missing
        smtplib.SMTPException: If email sending fails
        FileNotFoundError: If attachment file doesn't exist
    """
    try:
        # Validate configuration
        if not Config.EMAIL_USER or not Config.EMAIL_PASS:
            raise ValueError("Email credentials not configured in environment variables")
        
        if not to_address:
            raise ValueError("Recipient email address is required")
        
        print(f"📧 Preparing to send email...")
        print(f"   From: {Config.EMAIL_USER}")
        print(f"   To: {to_address}")
        if cc_address:
            print(f"   CC: {cc_address}")
        print(f"   Subject: {subject}")
        
        # Create message
        msg = EmailMessage()
        msg['From'] = Config.EMAIL_USER
        msg['To'] = to_address
        msg['Subject'] = subject
        
        # Add CC if provided
        if cc_address:
            msg['Cc'] = cc_address
        
        # Handle threading - critical for single thread functionality
        if thread_info:
            print(f"🔗 Adding threading headers...")
            
            # If replying to an existing message
            if thread_info.get('message_id'):
                msg['In-Reply-To'] = thread_info['message_id']
                print(f"   In-Reply-To: {thread_info['message_id']}")
            
            # Build References header for proper threading
            references = []
            if thread_info.get('references'):
                # Add existing references
                references.extend(thread_info['references'].split())
            if thread_info.get('message_id'):
                # Add the message we're replying to
                references.append(thread_info['message_id'])
            
            if references:
                msg['References'] = ' '.join(references)
                print(f"   References: {len(references)} message(s)")
        
        msg.set_content(body)

        # Attach file if provided
        if attachment_path:
            attachment_file = Path(attachment_path)
            if not attachment_file.exists():
                raise FileNotFoundError(f"Attachment file not found: {attachment_path}")
            
            try:
                with open(attachment_file, 'rb') as f:
                    file_data = f.read()
                    
                # Determine MIME type based on file extension
                file_extension = attachment_file.suffix.lower()
                if file_extension in ['.xlsx', '.xls']:
                    maintype = 'application'
                    subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                else:
                    maintype = 'application'
                    subtype = 'octet-stream'
                
                msg.add_attachment(
                    file_data, 
                    maintype=maintype, 
                    subtype=subtype, 
                    filename=attachment_file.name
                )
                print(f"📎 Attachment added: {attachment_file.name} ({len(file_data)} bytes)")
                
            except Exception as e:
                raise Exception(f"Failed to attach file {attachment_path}: {e}")

        # Connect and send
        print(f"🔗 Connecting to {Config.SMTP_SERVER}:{Config.SMTP_PORT}")
        
        try:
            with smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT) as server:
                server.starttls()  # Upgrade to secure connection
                server.login(Config.EMAIL_USER, Config.EMAIL_PASS)
                
                # Send the message
                server.send_message(msg)

            if thread_info:
                print("✅ Email sent successfully (in thread).")
            else:
                print("✅ Email sent successfully (new thread).")
            
        except smtplib.SMTPAuthenticationError:
            raise Exception("Email authentication failed. Check your email credentials.")
        except smtplib.SMTPRecipientsRefused:
            raise Exception(f"Recipients refused: {to_address}")
        except smtplib.SMTPException as e:
            raise Exception(f"SMTP error occurred: {e}")
            
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
        raise  # Re-raise to trigger error handling upstream

def send_thread_reply(to_address, body, attachment_path=None, cc_address=None, thread_info=None):
    """
    Convenience function to send a reply in the MAG bot conversation thread.
    Always uses the standard bot subject line for consistency.
    
    Args:
        to_address (str): Primary recipient email address
        body (str): Email body text
        attachment_path (str, optional): Path to file to attach
        cc_address (str, optional): CC recipient email address
        thread_info (dict, optional): Threading information for replies
    """
    subject = "Re: Interact with the MAG bot to configure your file"
    
    return send_email_with_attachment(
        to_address=to_address,
        subject=subject,
        body=body,
        attachment_path=attachment_path,
        cc_address=cc_address,
        thread_info=thread_info
    )

def send_conversation_starter(to_address, cc_address=None, thread_info=None):
    """
    Send the initial conversation starter when user sends "Start conversation".
    
    Args:
        to_address (str): Primary recipient email address
        cc_address (str, optional): CC recipient email address
        thread_info (dict, optional): Threading information from the original message
    """
    subject = "Re: Start conversation"
    body = """Hello,

I'm here to help you process your Salesforce pipeline files.

WHAT I CAN DO:

Adjust Columns: Type "Adjust Columns" and I'll send you the current template to modify. When you're done, reply with "Here" and attach your updated template.

Process Files: Attach your pipeline Excel file and I'll process it for you.

All our conversation will stay in this email thread.

Best regards,
MAG Pipeline Bot


Ready to get started? Just reply to this email with what you'd like to do!

Best regards,
MAG Pipeline Bot 

---
This is an automated response from your pipeline automation system.
"""
    
    return send_email_with_attachment(
        to_address=to_address,
        subject=subject,
        body=body,
        attachment_path=None,
        cc_address=cc_address,
        thread_info=thread_info
    )

def send_test_email():
    """Send a test email to verify configuration."""
    try:
        test_subject = "🧪 Pipeline System Test Email"
        test_body = """This is a test email from your Salesforce Pipeline Automation System.

If you received this email, your email configuration is working correctly!

System Information:
- SMTP Server: {smtp_server}:{smtp_port}
- From Address: {from_address}
- Timestamp: {timestamp}

Best regards,
Your Pipeline Automation System
""".format(
            smtp_server=Config.SMTP_SERVER,
            smtp_port=Config.SMTP_PORT, 
            from_address=Config.EMAIL_USER,
            timestamp=str(Path(__file__).stat().st_mtime)
        )
        
        send_email_with_attachment(
            to_address=Config.YOUR_EMAIL,
            subject=test_subject,
            body=test_body,
            attachment_path=None
        )
        
        print("✅ Test email sent successfully!")
        return True
        
    except Exception as e:
        print(f"❌ Test email failed: {e}")
        return False

if __name__ == "__main__":
    try:
        Config.validate_config()
        print("✅ Configuration validated")
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        sys.exit(1)
    
    # Send test email
    print("🧪 Sending test email...")
    if send_test_email():
        print("✅ Email system is working correctly!")
    else:
        print("❌ Email system test failed!")
        sys.exit(1)