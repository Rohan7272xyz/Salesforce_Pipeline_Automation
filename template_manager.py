import imaplib
import email
import os
import shutil
import logging
from datetime import datetime
from dotenv import load_dotenv
from send_email import send_email_with_attachment

# Load environment variables
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_PORT = int(os.getenv("IMAP_PORT", 993))

# Setup directories
BACKUP_DIR = "template_backups"
TEMP_DIR = "temp_templates"
os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

# Template path
TEMPLATE_PATH = r"C:\Users\rohan\Personal Projects\Email_Excel_Python_Alg\C5SDEC_Pipeline_Overview_v3_070325.xlsx"

# Contact info
YOUR_EMAIL = "Rohan.Anand@mag.us"
JOE_EMAIL = "Joseph.Findley@mag.us"

def backup_current_template():
    """Create a backup of the current template before replacing it."""
    if os.path.exists(TEMPLATE_PATH):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"template_backup_{timestamp}.xlsx"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        shutil.copy2(TEMPLATE_PATH, backup_path)
        print(f"📋 Template backed up to: {backup_path}")
        return backup_path
    return None

def replace_template(new_template_path):
    """Replace the current template with the new one from Joe."""
    try:
        # Create backup first
        backup_path = backup_current_template()
        
        # Replace the template
        shutil.copy2(new_template_path, TEMPLATE_PATH)
        print(f"✅ Template updated successfully!")
        print(f"🔄 Old template backed up to: {backup_path}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to replace template: {e}")
        return False

def send_template_to_joe():
    """Send the current template to Joe for modification."""
    try:
        template_subject = "📋 Excel Template for Column Adjustment"
        template_body = """Hi Joe,

You requested to adjust columns in the pipeline system. I'm sending you the current Excel template that the automation system uses.

INSTRUCTIONS:
1. Make the same column changes to this template file that you made to your raw data report
2. Save the file (keep the same format)
3. Reply to this email with the updated template attached
4. Use "Here" as the subject line (exactly like that - just the word "Here")

The system will automatically update to use your new template format.

If you have any questions about this process, please let me know.

Best regards,
Rohan's Pipeline System

---
This is an automated response to your "Adjust Columns" request.
"""
        
        if not os.path.exists(TEMPLATE_PATH):
            raise FileNotFoundError(f"Template file not found at {TEMPLATE_PATH}")
        
        send_email_with_attachment(
            to_address=JOE_EMAIL,
            cc_address=YOUR_EMAIL,
            subject=template_subject,
            body=template_body,
            attachment_path=TEMPLATE_PATH
        )
        
        print(f"📧 Template sent to Joe ({JOE_EMAIL})")
        logging.info(f"Template sent to Joe for column adjustment")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send template to Joe: {e}")
        logging.error(f"Failed to send template to Joe: {e}")
        return False

def send_template_confirmation(sender, success=True):
    """Send confirmation email about template update."""
    try:
        if success:
            subject = "✅ Template Updated Successfully"
            body = """Hi Joe,

Great news! The pipeline system template has been successfully updated with your changes.

WHAT HAPPENED:
- Your new template has been installed
- The old template was backed up for safety
- The system is now ready to process pipeline reports with the new column structure

The automated pipeline processing will now work correctly with your updated data format.

Best regards,
Rohan's Pipeline System

---
This is an automated confirmation of your template update.
"""
        else:
            subject = "❌ Template Update Failed"
            body = """Hi Joe,

There was an issue updating the pipeline system template.

WHAT TO DO:
1. Make sure your email subject line is exactly "Here" (just that word)
2. Make sure you attached an Excel (.xlsx) file
3. Try sending the template again

If you continue having issues, please contact Rohan directly at Rohan.Anand@mag.us

Best regards,
Rohan's Pipeline System

---
This is an automated error notification.
"""
        
        send_email_with_attachment(
            to_address=sender,
            cc_address=YOUR_EMAIL,
            subject=subject,
            body=body,
            attachment_path=None
        )
        
        print(f"📧 Template confirmation sent to {sender}")
        logging.info(f"Template update confirmation sent - Success: {success}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send template confirmation: {e}")
        logging.error(f"Failed to send template confirmation: {e}")
        return False

def check_template_management_emails():
    """
    Check for template management emails (separate from normal pipeline processing).
    Returns: tuple (action_type, result_data) or None
    """
    try:
        print(f"🔧 Checking for template management emails...")
        
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        # Search for UNSEEN emails from Joe
        search_criteria = f'(UNSEEN FROM "{JOE_EMAIL}")'
        status, messages = mail.search(None, search_criteria)

        if status != "OK" or not messages[0]:
            mail.logout()
            return None

        email_ids = messages[0].split()
        
        # Check each email for template management subjects
        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            sender = email.utils.parseaddr(msg["From"])[1]
            subject = msg.get("Subject", "").strip()
            
            print(f"🔍 Found email from {sender} with subject: '{subject}'")

            # Check for "Adjust Columns" request
            if subject.lower() == "adjust columns":
                print("🔧 TEMPLATE ADJUSTMENT REQUEST DETECTED!")
                mail.logout()
                return ("ADJUST_COLUMNS", sender)
                
            # Check for "Here" (template update)
            elif subject.lower() == "here":
                print("📥 NEW TEMPLATE RECEIVED!")
                
                # Download the new template
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue

                    filename = part.get_filename()
                    if filename and filename.endswith(".xlsx"):
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        temp_template_path = os.path.join(TEMP_DIR, f"new_template_{timestamp}.xlsx")

                        with open(temp_template_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        print(f"📥 Downloaded new template to: {temp_template_path}")
                        
                        # Replace the current template
                        if replace_template(temp_template_path):
                            mail.logout()
                            return ("TEMPLATE_UPDATED", sender)
                        else:
                            mail.logout()
                            return ("TEMPLATE_UPDATE_FAILED", sender)

                # No Excel file found in "Here" email
                print("❌ No Excel template found in 'Here' email.")
                mail.logout()
                return ("TEMPLATE_UPDATE_FAILED", sender)

        mail.logout()
        return None

    except Exception as e:
        print(f"❌ Template management email check failed: {e}")
        return None

def process_template_management():
    """
    Main function to handle template management workflow.
    Returns True if a template management action was processed, False otherwise.
    """
    result = check_template_management_emails()
    
    if not result:
        return False
        
    action_type, sender = result
    
    if action_type == "ADJUST_COLUMNS":
        print(f"🔧 Processing column adjustment request from {sender}")
        if send_template_to_joe():
            logging.info(f"Template adjustment request processed for {sender}")
            return True
        else:
            logging.error(f"Failed to process template adjustment request from {sender}")
            return True  # Still processed, even if failed
    
    elif action_type == "TEMPLATE_UPDATED":
        print(f"✅ Template successfully updated by {sender}")
        send_template_confirmation(sender, success=True)
        logging.info(f"Template successfully updated by {sender}")
        return True
    
    elif action_type == "TEMPLATE_UPDATE_FAILED":
        print(f"❌ Template update failed from {sender}")
        send_template_confirmation(sender, success=False)
        logging.error(f"Template update failed from {sender}")
        return True
    
    return False

if __name__ == "__main__":
    # Test the template management system
    print("🔧 Testing Template Management System...")
    
    if process_template_management():
        print("✅ Template management action processed")
    else:
        print("📭 No template management emails found")