import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
import smtplib
from email.message import EmailMessage
import os
from pathlib import Path
import sys

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config import Config

def send_email_with_attachment(to_address, subject, body, attachment_path=None, cc_address=None):
    """
    Send email with optional attachment and CC support.
    
    Args:
        to_address (str): Primary recipient email address
        subject (str): Email subject
        body (str): Email body text
        attachment_path (str, optional): Path to file to attach (None for no attachment)
        cc_address (str, optional): CC recipient email address
        
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

            print("✅ Email sent successfully.")
            
        except smtplib.SMTPAuthenticationError:
            raise Exception("Email authentication failed. Check your email credentials.")
        except smtplib.SMTPRecipientsRefused:
            raise Exception(f"Recipients refused: {to_address}")
        except smtplib.SMTPException as e:
            raise Exception(f"SMTP error occurred: {e}")
            
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
        raise  # Re-raise to trigger error handling upstream

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
def send_email_with_attachment(to_address, subject, body, attachment_path=None, cc_address=None):
    """
    Send email with optional attachment and CC support.
    
    Args:
        to_address (str): Primary recipient email address
        subject (str): Email subject
        body (str): Email body text
        attachment_path (str, optional): Path to file to attach (None for no attachment)
        cc_address (str, optional): CC recipient email address
    """
    try:
        msg = EmailMessage()
        msg['From'] = EMAIL_USER
        msg['To'] = to_address
        msg['Subject'] = subject
        
        # Add CC if provided
        if cc_address:
            msg['Cc'] = cc_address
            print(f"📧 Sending to: {to_address} (CC: {cc_address})")
        else:
            print(f"📧 Sending to: {to_address}")
        
        msg.set_content(body)

        # Attach file if provided
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(attachment_path)
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
            print(f"📎 Attachment added: {file_name}")
        elif attachment_path:
            print(f"⚠️ Attachment file not found: {attachment_path}")

        # Connect and send
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Upgrades to secure connection
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)

        print("✅ Email sent successfully.")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
        raise  # Re-raise to trigger error handling

if __name__ == "__main__":
    # Test function
    print("This is the send_email module. Import and use send_email_with_attachment() function.")