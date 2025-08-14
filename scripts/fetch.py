import imaplib
import email
import shutil
import re
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

def extract_thread_info(msg):
    """Extract threading information from email headers."""
    thread_info = {
        'message_id': msg.get('Message-ID'),
        'in_reply_to': msg.get('In-Reply-To'),
        'references': msg.get('References'),
        'subject': msg.get('Subject', '').strip()
    }
    return thread_info

def is_thread_continuation(msg):
    """Check if this email is part of an existing thread."""
    thread_info = extract_thread_info(msg)
    
    # Check if it has In-Reply-To or References headers
    if thread_info['in_reply_to'] or thread_info['references']:
        return True, thread_info
    
    # Check if subject indicates it's a reply (Re: pattern)
    subject = thread_info['subject'].lower()
    if subject.startswith('re:') and 'interact with the mag bot' in subject:
        return True, thread_info
    
    return False, thread_info

def parse_email_body_for_commands(body_text):
    """
    FIXED: Parse email body for command keywords with proper user message extraction.
    Now correctly handles case-insensitive matching for "change format"
    """
    if not body_text:
        return None
    
    try:
        # Clean and normalize the body text
        body_lower = body_text.lower().strip()
        
        # Remove common email artifacts
        body_lower = re.sub(r'<[^>]+>', '', body_lower)  # Remove HTML tags
        body_lower = re.sub(r'\s+', ' ', body_lower)  # Normalize whitespace
        
        print(f"📝 Full body preview: {body_lower[:100]}...")
        
        # IMPROVED: Look for separator patterns and extract text BEFORE them
        separator_patterns = [
            '________________________________',
            'from: magpipelinemanager@gmail.com',
            'from:magpipelinemanager@gmail.com',
            'sent: monday',
            'sent: tuesday', 
            'sent: wednesday',
            'sent: thursday',
            'sent: friday',
            'sent: saturday',
            'sent: sunday',
            'hi there',
            'i see you want to adjust',
            'hello,',
            "i'm here to help",
            "i'm sending you"
        ]
        
        # Find the earliest separator
        earliest_separator_pos = len(body_lower)
        for pattern in separator_patterns:
            pos = body_lower.find(pattern)
            if pos != -1 and pos < earliest_separator_pos:
                earliest_separator_pos = pos
        
        # Extract user message (everything before the first separator)
        if earliest_separator_pos < len(body_lower):
            user_message = body_lower[:earliest_separator_pos].strip()
        else:
            user_message = body_lower.strip()
        
        # Clean up the user message
        user_message = re.sub(r'\s+', ' ', user_message)
        user_message = user_message.strip()
        
        print(f"📝 User message only: {user_message[:100]}...")
        
        # CHECK "HERE" FIRST - if user message starts with "here", that takes priority
        if (user_message.startswith('here') or 
            user_message == 'here' or
            user_message.startswith('here ') or
            user_message.startswith('here.')):
            print("📥 Command detected: HERE (template update)")
            return "HERE"
        
        # FIXED: Check for "change format" with lowercase comparison since user_message is lowercase
        if 'change format' in user_message or 'adjust column' in user_message:
            print("🔧 Command detected: Change Format")
            return "ADJUST_COLUMNS"
        
        # Check if it's just confidentiality notice or very short
        if (len(user_message) < 10 or
            'confidentiality notice' in user_message or
            'start conversation' in user_message or
            user_message == ''):
            print("ℹ️ Command detected: HELP/START (sending instructions)")
            return "HELP"
        
        # No specific command - regular processing
        print(f"ℹ️ No specific command detected in user message")
        return None
        
    except Exception as e:
        print(f"⚠️ Error parsing email body for commands: {e}")
        return None

def get_email_body(msg):
    """Extract the email body text."""
    body = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                charset = part.get_content_charset() or 'utf-8'
                try:
                    body += part.get_payload(decode=True).decode(charset)
                except:
                    body += part.get_payload(decode=True).decode('utf-8', errors='ignore')
    else:
        charset = msg.get_content_charset() or 'utf-8'
        try:
            body = msg.get_payload(decode=True).decode(charset)
        except:
            body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
    
    return body.strip()

def download_latest_attachment():
    """
    Download latest attachment and handle different email types:
    - "Start conversation": Initiate bot interaction thread
    - Thread continuation with body commands: "Change Format", "Here"
    - Thread continuation with file: Normal pipeline processing
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

        # Extract threading information
        is_thread, thread_info = is_thread_continuation(msg)
        
        # Check for "Start conversation" to initiate new thread
        if subject.lower() == "start conversation":
            print("🚀 NEW CONVERSATION START DETECTED!")
            mail.logout()
            return ("START_CONVERSATION", sender, subject, thread_info)
        
        # If it's a thread continuation, parse body for commands
        if is_thread:
            print("🔗 THREAD CONTINUATION DETECTED!")
            body_text = get_email_body(msg)
            print(f"📝 Email body preview: {body_text[:100]}...")
            
            # Parse body for commands using FIXED function
            command = parse_email_body_for_commands(body_text)
            
            if command == 'ADJUST_COLUMNS':
                print("🔧 Change Format command found in thread!")
                mail.logout()
                return ("THREAD_ADJUST_COLUMNS", sender, subject, thread_info)
            
            elif command == 'HERE':
                print("📥 HERE command found in thread - looking for template...")
                # Look for Excel attachment
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
                                    # Find the latest input file and pass it to template analyzer
                                    latest_input_file = find_latest_input_file()
                                    
                                    from template_analyzer import analyze_template_and_update_app
                                    print(f"🔍 Running template analyzer with input file: {latest_input_file}")
                                    
                                    if analyze_template_and_update_app(latest_input_file):
                                        mail.logout()
                                        return ("THREAD_TEMPLATE_UPDATED", sender, str(temp_template_path), thread_info)
                                    else:
                                        mail.logout()
                                        return ("THREAD_TEMPLATE_UPDATE_FAILED", sender, "Failed to update app.py", thread_info)
                                except Exception as e:
                                    print(f"❌ Error during app update: {e}")
                                    import traceback
                                    print(traceback.format_exc())
                                    mail.logout()
                                    return ("THREAD_TEMPLATE_UPDATE_FAILED", sender, str(e), thread_info)
                            else:
                                mail.logout()
                                return ("THREAD_TEMPLATE_UPDATE_FAILED", sender, "Failed to replace template", thread_info)
                                
                        except Exception as e:
                            print(f"❌ Error saving template: {e}")
                            mail.logout()
                            return ("THREAD_TEMPLATE_UPDATE_FAILED", sender, str(e), thread_info)
                            
                        template_found = True
                        break

                if not template_found:
                    print("❌ No Excel template found in 'Here' thread message.")
                    mail.logout()
                    return ("THREAD_TEMPLATE_UPDATE_FAILED", sender, "No attachment found", thread_info)
            
            else:
                # Check for file attachment (normal pipeline processing in thread)
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
                            return ("THREAD_NORMAL_PROCESSING", str(saved_path), sender, thread_info)
                            
                        except Exception as e:
                            print(f"❌ Error saving attachment: {e}")
                            mail.logout()
                            return None
                            
                        attachment_found = True
                        break

                if not attachment_found:
                    print("📂 No recognized command or attachment found in thread continuation.")
                    mail.logout()
                    return ("THREAD_UNCLEAR", sender, subject, thread_info)
        
        # Legacy support: Check for old-style subject-based commands (for backward compatibility)
        else:
            print("📧 Processing as legacy email (not in thread)")
            
            if subject.lower() == "change format":
                print("🔧 LEGACY: Column adjustment request")
                mail.logout()
                return ("ADJUST_COLUMNS", sender, subject, thread_info)
                
            elif subject.lower() == "here":
                print("📥 LEGACY: Template update")
                # Handle same as before for backward compatibility
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
                            
                            if replace_template(temp_template_path):
                                try:
                                    latest_input_file = find_latest_input_file()
                                    from template_analyzer import analyze_template_and_update_app
                                    print(f"🔍 Running template analyzer with input file: {latest_input_file}")
                                    
                                    if analyze_template_and_update_app(latest_input_file):
                                        mail.logout()
                                        return ("TEMPLATE_UPDATED", sender, str(temp_template_path), thread_info)
                                    else:
                                        mail.logout()
                                        return ("TEMPLATE_UPDATE_FAILED", sender, "Failed to update app.py", thread_info)
                                except Exception as e:
                                    print(f"❌ Error during app update: {e}")
                                    mail.logout()
                                    return ("TEMPLATE_UPDATE_FAILED", sender, str(e), thread_info)
                            else:
                                mail.logout()
                                return ("TEMPLATE_UPDATE_FAILED", sender, "Failed to replace template", thread_info)
                                
                        except Exception as e:
                            print(f"❌ Error saving template: {e}")
                            mail.logout()
                            return ("TEMPLATE_UPDATE_FAILED", sender, str(e), thread_info)
                            
                        template_found = True
                        break

                if not template_found:
                    print("❌ No Excel template found in 'Here' email.")
                    mail.logout()
                    return ("TEMPLATE_UPDATE_FAILED", sender, "No attachment found", thread_info)

            # Normal pipeline processing (legacy)
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
                            return ("NORMAL_PROCESSING", str(saved_path), sender, thread_info)
                            
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
        
        if action_type == "START_CONVERSATION":
            sender = result[1]
            print(f"Ready to start conversation with {sender}")
        elif action_type.startswith("THREAD_"):
            sender = result[1]
            print(f"Thread action from {sender}: {action_type}")
        elif action_type == "NORMAL_PROCESSING":
            file_path, sender = result[1], result[2]
            print(f"Ready for normal pipeline processing from {sender}")
        else:
            print(f"Other action: {action_type}")
    else:
        print("❌ No valid emails found or processed")