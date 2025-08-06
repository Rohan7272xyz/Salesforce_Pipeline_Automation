import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

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