import imaplib
import email
import shutil
from datetime import datetime
from pathlib import Path
import sys

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config import Config

def backup_current_template():
    """Create a backup of the current template before replacing it."""
    if Config.TEMPLATE_PATH.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"template_backup_{timestamp}.xlsx"
        backup_path = Config.BACKUP_DIR / backup_filename
        shutil.copy2(Config.TEMPLATE_PATH, backup_path)
        print(f"📋 Template backed up to: {backup_path}")
        return backup_path
    return None

def replace_template(new_template_path):
    """Replace the current template with the new one from Joe."""
    try:
        # Create backup first
        backup_path = backup_current_template()
        
        # Replace the template
        shutil.copy2(new_template_path, Config.TEMPLATE_PATH)
        print(f"✅ Template updated successfully!")
        print(f"🔄 Old template backed up to: {backup_path}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to replace template: {e}")
        return False

def find_latest_input_file():
    """Find the most recent pipeline input file to use for analysis."""
    try:
        input_files = list(Config.INPUT_DIR.glob("pipeline_*.xlsx"))
        if not input_files:
            print("⚠️ No input files found for template analysis")
            return None
        
        # Get the most recent file
        latest_file = max(input_files, key=lambda f: f.stat().st_mtime)
        print(f"📂 Found latest input file for analysis: {latest_file.name}")
        return str(latest_file)
        
    except Exception as e:
        print(f"⚠️ Error finding latest input file: {e}")
        return None

def download_latest_attachment():
    """
    Download latest attachment and handle different email subjects:
    - Normal processing: Regular pipeline reports
    - "Adjust Columns": Send template to Joe for updates
    - "Here": Replace current template with Joe's updated version
    """
    try:
        # Validate configuration first
        if not Config.EMAIL_PASS:
            print("❌ Email password not configured in environment variables")
            return None

        print(f"Connecting to {Config.IMAP_SERVER}:{Config.IMAP_PORT} as {Config.EMAIL_USER}")

        mail = imaplib.IMAP4_SSL(Config.IMAP_SERVER, Config.IMAP_PORT)
        mail.login(Config.EMAIL_USER, Config.EMAIL_PASS)
        mail.select("inbox")

        # Search for UNSEEN emails from authorized sender
        search_criteria = f'(UNSEEN FROM "{Config.AUTHORIZED_EMAILS[0]}")'
        
        print(f"🔍 Searching with criteria: {search_criteria}")
        
        status, messages = mail.search(None, search_criteria)

        if status != "OK" or not messages[0]:
            print(f"📭 No new unread emails from {Config.AUTHORIZED_EMAILS[0]}.")
            return None

        email_ids = messages[0].split()
        latest_email_id = email_ids[-1]

        status, msg_data = mail.fetch(latest_email_id, "(RFC822)")
        if status != "OK":
            print("❌ Failed to fetch email")
            mail.logout()
            return None
            
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        sender = email.utils.parseaddr(msg["From"])[1]
        subject = msg.get("Subject", "").strip()
        
        print(f"📧 Processing email from: {sender}")
        print(f"📋 Subject: {subject}")

        # Verify sender is authorized
        if sender not in Config.AUTHORIZED_EMAILS:
            print(f"⚠️ Unauthorized sender: {sender}")
            mail.logout()
            return None

        # Check for special subjects
        if subject.lower() == "adjust columns":
            print("🔧 TEMPLATE ADJUSTMENT REQUEST DETECTED!")
            mail.logout()
            return ("ADJUST_COLUMNS", sender, subject)
            
        elif subject.lower() == "here":
            print("📥 NEW TEMPLATE RECEIVED!")
            template_found = False
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue

                filename = part.get_filename()
                if filename and any(filename.lower().endswith(ext) for ext in ['.xlsx', '.xls']):
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    temp_template_path = Config.INPUT_DIR / f"new_template_{timestamp}.xlsx"

                    try:
                        with open(temp_template_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        print(f"📥 Downloaded new template to: {temp_template_path}")
                        
                        # Replace the current template
                        if replace_template(temp_template_path):
                            try:
                                # CRITICAL FIX: Find the latest input file and pass it to template analyzer
                                latest_input_file = find_latest_input_file()
                                
                                from template_analyzer import analyze_template_and_update_app
                                print(f"🔍 Running template analyzer with input file: {latest_input_file}")
                                
                                # Pass the latest input file to ensure correct column mapping
                                if analyze_template_and_update_app(latest_input_file):
                                    mail.logout()
                                    return ("TEMPLATE_UPDATED", sender, str(temp_template_path))
                                else:
                                    mail.logout()
                                    return ("TEMPLATE_UPDATE_FAILED", sender, "Failed to update app.py")
                            except Exception as e:
                                print(f"❌ Error during app update: {e}")
                                import traceback
                                print(traceback.format_exc())
                                mail.logout()
                                return ("TEMPLATE_UPDATE_FAILED", sender, str(e))
                        else:
                            mail.logout()
                            return ("TEMPLATE_UPDATE_FAILED", sender, "Failed to replace template")
                            
                    except Exception as e:
                        print(f"❌ Error saving template: {e}")
                        mail.logout()
                        return ("TEMPLATE_UPDATE_FAILED", sender, str(e))
                        
                    template_found = True
                    break

            if not template_found:
                print("❌ No Excel template found in 'Here' email.")
                mail.logout()
                return ("TEMPLATE_UPDATE_FAILED", sender, "No attachment found")

        # Normal pipeline processing
        else:
            attachment_found = False
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue

                filename = part.get_filename()
                if filename and any(filename.lower().endswith(ext) for ext in ['.xlsx', '.xls']):
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    saved_path = Config.INPUT_DIR / f"pipeline_{timestamp}.xlsx"

                    try:
                        with open(saved_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        print(f"📥 Saved pipeline data from {sender} to: {saved_path}")
                        mail.logout()
                        return ("NORMAL_PROCESSING", str(saved_path), sender)
                        
                    except Exception as e:
                        print(f"❌ Error saving attachment: {e}")
                        mail.logout()
                        return None
                        
                    attachment_found = True
                    break

            if not attachment_found:
                print("📂 No Excel attachment found in the email.")

        mail.logout()
        return None

    except imaplib.IMAP4.error as e:
        print(f"❌ IMAP error: {e}")
        return None
    except Exception as e:
        print(f"❌ Unexpected error in email processing: {e}")
        import traceback
        print(traceback.format_exc())
        return None

if __name__ == "__main__":
    try:
        Config.validate_config()
        print("✅ Configuration validated")
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        sys.exit(1)
        
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