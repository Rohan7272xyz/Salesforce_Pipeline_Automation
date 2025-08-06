import imaplib
import email
import os
import shutil
from datetime import datetime
from dotenv import load_dotenv

# ✅ Load environment variables
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_PORT = int(os.getenv("IMAP_PORT", 993))  # fallback to 993 if not set

INPUT_DIR = "input"
BACKUP_DIR = "template_backups"
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

# Template path
TEMPLATE_PATH = r"C:\Users\rohan\Personal Projects\Email_Excel_Python_Alg\C5SDEC_Pipeline_Overview_v3_070325.xlsx"

# ✅ Only Joe's email address is authorized
AUTHORIZED_EMAILS = [
    "Joseph.Findley@mag.us"
]

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

def download_latest_attachment():
    """
    Download latest attachment and handle different email subjects:
    - Normal processing: Regular pipeline reports
    - "Adjust Columns": Send template to Joe for updates
    - "Here": Replace current template with Joe's updated version
    """
    try:
        # ✅ Debug print for validation
        print(f"Connecting to {IMAP_SERVER}:{IMAP_PORT} as {EMAIL_USER}")

        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        # ✅ Search for UNSEEN emails from authorized sender
        search_criteria = f'(UNSEEN FROM "{AUTHORIZED_EMAILS[0]}")'
        
        print(f"🔍 Searching with criteria: {search_criteria}")
        
        status, messages = mail.search(None, search_criteria)

        if status != "OK" or not messages[0]:
            print(f"📭 No new unread emails from {AUTHORIZED_EMAILS[0]}.")
            return None

        email_ids = messages[0].split()
        latest_email_id = email_ids[-1]

        status, msg_data = mail.fetch(latest_email_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        sender = email.utils.parseaddr(msg["From"])[1]
        subject = msg.get("Subject", "").strip()
        
        print(f"📧 Processing email from: {sender}")
        print(f"📋 Subject: {subject}")

        # Check for special subjects
        if subject.lower() == "adjust columns":
            print("🔧 TEMPLATE ADJUSTMENT REQUEST DETECTED!")
            mail.logout()
            return ("ADJUST_COLUMNS", sender, subject)
            
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
                    temp_template_path = os.path.join(INPUT_DIR, f"new_template_{timestamp}.xlsx")

                    with open(temp_template_path, "wb") as f:
                        f.write(part.get_payload(decode=True))

                    print(f"📥 Downloaded new template to: {temp_template_path}")
                    
                    # Replace the current template
                    if replace_template(temp_template_path):
                        mail.logout()
                        return ("TEMPLATE_UPDATED", sender, temp_template_path)
                    else:
                        mail.logout()
                        return ("TEMPLATE_UPDATE_FAILED", sender, temp_template_path)

            print("❌ No Excel template found in 'Here' email.")
            mail.logout()
            return ("TEMPLATE_UPDATE_FAILED", sender, "No attachment found")

        # Normal pipeline processing
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            filename = part.get_filename()
            if filename and filename.endswith(".xlsx"):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                saved_path = os.path.join(INPUT_DIR, f"pipeline_{timestamp}.xlsx")

                with open(saved_path, "wb") as f:
                    f.write(part.get_payload(decode=True))

                print(f"📥 Saved pipeline data from {sender} to: {saved_path}")
                mail.logout()
                return ("NORMAL_PROCESSING", saved_path, sender)

        print("📂 No Excel attachment found in the email.")
        mail.logout()
        return None

    except Exception as e:
        print(f"❌ IMAP connection failed: {e}")
        return None

if __name__ == "__main__":
    result = download_latest_attachment()
    if result:
        action_type = result[0]
        print(f"✅ Action detected: {action_type}")
        
        if action_type == "NORMAL_PROCESSING":
            file_path, sender = result[1], result[2]
            print(f"Ready for normal pipeline processing from {sender}")
        elif action_type == "ADJUST_COLUMNS":
            sender = result[1]
            print(f"Template adjustment request from {sender}")
        elif action_type == "TEMPLATE_UPDATED":
            sender = result[1]
            print(f"Template successfully updated from {sender}")
        elif action_type == "TEMPLATE_UPDATE_FAILED":
            sender = result[1]
            print(f"Template update failed from {sender}")
    else:
        print("❌ No valid emails found or processed")