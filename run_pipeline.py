#!/usr/bin/env python3
"""
Production Pipeline with Robust Single Thread Embedding
Implements the complete mitigation plan for reliable email threading
"""

import logging
import time
import sys
from pathlib import Path

# Add project root to Python path for imports
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "scripts"))

from config import Config
from scripts.fetch import download_latest_attachment
from scripts.process import generate_gantt_chart
from robust_email_sender import RobustEmailSender

# Configuration
DEBUG_MODE = True  # Set to False for production
THREADING_ENABLED = True  # Set to False to disable threading entirely

print("✅ Running PRODUCTION pipeline with Single Thread Embedding (v8)")
print(f"🐛 Debug mode: {'ON' if DEBUG_MODE else 'OFF'}")
print(f"🔗 Threading: {'ENABLED' if THREADING_ENABLED else 'DISABLED'}")

# Setup logging
Config.ensure_directories()
logging.basicConfig(
    filename=Config.LOGS_DIR / "pipeline_log.txt",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# Initialize robust email sender
email_sender = RobustEmailSender(debug=DEBUG_MODE)

def send_help_instructions(sender, thread_info=None):
    """Send help/instructions to user when they start a conversation."""
    try:
        help_body = """Hello,

I'm here to help you process your Salesforce pipeline files.

WHAT I CAN DO:

1. Assist in Sheet Formatting Changes: If you tell me to "Change Format", I'll send you the current excel template to modify. When you're done, reply with "Here" and attach the updated template.

2. Process Files: Attach your pipeline Excel file and I'll organize/process it for you.

Ready to get started? Just reply to this email with what you'd like to do! All our conversations will stay in this email thread.

Best regards,
MAG Pipeline Bot

---
Type "Help" anytime to see these instructions again
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        success = email_sender.send_email(
            to_address=sender,
            subject="Interact with the MAG bot to configure your file",
            body=help_body,
            cc_address=Config.YOUR_EMAIL,
            thread_info=use_threading
        )
        
        if success:
            print(f"📧 Help instructions sent to {sender}")
            logging.info(f"Help instructions sent to {sender}")
        
        return success
        
    except Exception as e:
        print(f"❌ Failed to send help instructions: {e}")
        logging.error(f"Failed to send help instructions: {e}")
        return False

def send_template_to_user(sender, thread_info=None):
    """Send the current template to user for modification."""
    try:
        template_body = """Hello,

I'm sending you the current Excel template that the system uses.

INSTRUCTIONS:
1. Download and open the attached template file
2. Make your column/format changes to this template
3. Save the file as .xlsx format
4. IMPORTANT: Reply to this email thread with your updated template attached and "Here" typed in the message body
5. IMPORTANT: Type "Here" in the message body 

Best regards,
MAG Pipeline Bot 

---
Automated response to your "Adjust Columns" request
"""
        
        if not Config.TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"Template file not found at {Config.TEMPLATE_PATH}")
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        success = email_sender.send_email(
            to_address=sender,
            subject="Template for Column Adjustment",
            body=template_body,
            attachment_path=str(Config.TEMPLATE_PATH),
            cc_address=Config.YOUR_EMAIL,
            thread_info=use_threading
        )
        
        if success:
            print(f"📧 Template sent to {sender}")
            logging.info(f"Template sent to {sender} for column adjustment")
        
        return success
        
    except Exception as e:
        print(f"❌ Failed to send template to user: {e}")
        logging.error(f"Failed to send template to user: {e}")
        return False

def send_template_confirmation(sender, success=True, thread_info=None, error_details=None):
    """Send confirmation email about template update."""
    try:
        if success:
            body = """Hello,

Your template has been successfully updated in my server.

The system is now ready to process pipeline files with your new formatting change.

You can now attach and send your pipeline files for processing.

Best regards,
MAG Pipeline Bot

---
Automated confirmation of successful template update
"""
        else:
            body = f"""Oops! There was an issue updating the template. ❌

ERROR DETAILS
{error_details or 'Unknown error occurred'}

WHAT TO DO:
1. Make sure you typed "Here" at the beginning of your message
2. Make sure you attached an Excel (.xlsx) file
3. Try replying to this thread again with:
   - The word "Here" in the message body
   - Your template file attached

 TIPS:
• Don't include extra text before "Here"
• Make sure the file is .xlsx format
• Check that the file isn't corrupted

If you continue having issues, please contact Rohan directly at Rohan.Anand@mag.us

Best regards,
MAG Pipeline Bot 

---
Automated error notification - template update failed
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        subject = "Template Update Confirmation" if success else "Template Update Failed"
        
        email_success = email_sender.send_email(
            to_address=sender,
            subject=subject,
            body=body,
            cc_address=Config.YOUR_EMAIL,
            thread_info=use_threading
        )
        
        if email_success:
            print(f"📧 Template confirmation sent to {sender}")
            logging.info(f"Template update confirmation sent - Success: {success}")
        
        return email_success
        
    except Exception as e:
        print(f"❌ Failed to send template confirmation: {e}")
        logging.error(f"Failed to send template confirmation: {e}")
        return False

def send_successful_processing_email(sender, final_file, thread_info=None):
    """Send successful processing notification with file."""
    try:
        # Determine the greeting based on sender
        if "joseph.findley" in sender.lower() or "joe" in sender.lower():
            greeting_name = "Joe"
        else:
            # Extract first name from email or use a generic greeting
            sender_parts = sender.split('@')[0].split('.')
            if len(sender_parts) > 0:
                greeting_name = sender_parts[0].title()
            else:
                greeting_name = "there"
        
        success_body = f"""Hi {greeting_name}!

I have successfully processed your Salesforce pipeline and made all the necessary adjustments.

WHAT I DID:
- Applied all formatting rules
- Sorted data by capture manager
- Generated the Gantt chart view
- Cleaned up the data presentation

The formatted file is attached.

Best regards,
MAG Pipeline Bot

---
Pipeline processed at {time.strftime('%Y-%m-%d %H:%M:%S')}
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        success = email_sender.send_email(
            to_address=sender,
            subject="Your Processed Salesforce Pipeline",
            body=success_body,
            attachment_path=final_file,
            cc_address=Config.YOUR_EMAIL,
            thread_info=use_threading
        )
        
        if success:
            logging.info(f"Pipeline processed and sent successfully to {sender}")
            print("✅ Pipeline processing completed successfully.")
        
        return success
        
    except Exception as e:
        print(f"❌ Failed to send success email: {e}")
        logging.error(f"Failed to send success email: {e}")
        return False

def send_error_alert_email(error_message, sender_email):
    """Send error alert email to admin."""
    try:
        alert_subject = "🚨 URGENT: Salesforce Pipeline Automation Error"
        alert_body = f"""PIPELINE ERROR ALERT

TIME: {time.strftime('%Y-%m-%d %H:%M:%S')}
SENDER: {sender_email}
ERROR: {error_message}

The automated Salesforce pipeline system has encountered an error and needs your attention.

The user has been automatically notified that there's an issue and that you're working on a fix.

Please check the system logs and resolve the issue as soon as possible.

---
This is an automated error alert from your pipeline system.
"""
        
        success = email_sender.send_email(
            to_address=Config.YOUR_EMAIL,
            subject=alert_subject,
            body=alert_body,
            thread_info=None  # Always send errors as new threads
        )
        
        if success:
            print(f"📧 Error alert sent to admin ({Config.YOUR_EMAIL})")
            logging.info(f"Error alert email sent to {Config.YOUR_EMAIL}")
        
        return success
        
    except Exception as e:
        print(f"❌ Failed to send error alert email: {e}")
        logging.error(f"Error alert email failed: {e}")
        return False

def send_error_email_to_user(error_message, sender_email, thread_info=None):
    """Send error notification email to user."""
    try:
        error_body = f"""Hi there,

I wanted to let you know that there was a technical issue with processing your request.

ISSUE DETAILS:
• Time: {time.strftime('%Y-%m-%d %H:%M:%S')}
• Error: {error_message}

NEXT STEPS:
• Rohan has been automatically notified
• He's working on identifying and fixing the issue
• The system will be back online as quickly as possible

In the meantime, if you need immediate pipeline processing, please contact Rohan directly at Rohan.Anand@mag.us

Best regards,
MAG Pipeline Bot 

---
Automated error notification
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        success = email_sender.send_email(
            to_address=sender_email,
            subject="Pipeline Processing Error",
            body=error_body,
            cc_address=Config.YOUR_EMAIL,
            thread_info=use_threading
        )
        
        if success:
            print(f"📧 Error notification sent to user ({sender_email})")
            logging.info(f"Error notification email sent to user: {sender_email}")
        
        return success
        
    except Exception as e:
        print(f"❌ Failed to send error email to user: {e}")
        logging.error(f"Error email to user failed: {e}")
        return False

def handle_error(error_message, sender_email=None, thread_info=None):
    """Handle errors by sending alert email to admin and notification email to user."""
    print(f"🚨 PIPELINE ERROR: {error_message}")
    
    # Log the error
    logging.error(f"Pipeline error: {error_message} (sender: {sender_email})")
    
    # Send alert email to admin
    send_error_alert_email(error_message, sender_email or "Unknown")
    
    # Send error email to user (with threading if available)
    if sender_email:
        send_error_email_to_user(error_message, sender_email, thread_info)

def main():
    """Main pipeline loop with robust single thread support."""
    try:
        # Validate configuration before starting
        Config.validate_config()
        print("✅ Configuration validated successfully")
        
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        logging.error(f"Configuration error: {e}")
        return
    
    print(f"🔄 Starting pipeline monitoring (checking every {Config.CHECK_INTERVAL_SECONDS} seconds)")
    print("📧 Single Thread Embedding: ACTIVE")
    
    while True:
        if DEBUG_MODE:
            print("🔍 Checking for new emails...")
        
        try:
            result = download_latest_attachment()
            if result:
                action_type = result[0]
                print(f"🎯 Action detected: {action_type}")
                
                if action_type == "START_CONVERSATION":
                    # User wants to start a new conversation
                    sender = result[1]
                    subject = result[2]
                    thread_info = result[3]
                    print(f"🚀 Starting new conversation with {sender}")
                    
                    if send_help_instructions(sender, None):  # Don't thread the initial welcome
                        logging.info(f"New conversation started with {sender}")
                    else:
                        handle_error("Failed to send conversation starter", sender, None)
                
                elif action_type in ["ADJUST_COLUMNS", "THREAD_ADJUST_COLUMNS"]:
                    # User wants to adjust columns
                    sender = result[1]
                    subject = result[2]
                    thread_info = result[3]
                    print(f"🔧 Column adjustment request from {sender}")
                    
                    # Use threading for THREAD_ADJUST_COLUMNS
                    use_threading = thread_info if action_type == "THREAD_ADJUST_COLUMNS" else None
                    
                    if send_template_to_user(sender, use_threading):
                        logging.info(f"Template adjustment request processed for {sender}")
                    else:
                        handle_error("Failed to send template to user", sender, thread_info)
                
                elif action_type in ["TEMPLATE_UPDATED", "THREAD_TEMPLATE_UPDATED"]:
                    # User sent back the updated template - confirm success
                    sender = result[1]
                    template_path = result[2]
                    thread_info = result[3]
                    print(f"✅ Template successfully updated by {sender}")
                    
                    # Use threading for THREAD_TEMPLATE_UPDATED
                    use_threading = thread_info if action_type == "THREAD_TEMPLATE_UPDATED" else None
                    
                    send_template_confirmation(sender, success=True, thread_info=use_threading)
                    logging.info(f"Template successfully updated by {sender}")
                
                elif action_type in ["TEMPLATE_UPDATE_FAILED", "THREAD_TEMPLATE_UPDATE_FAILED"]:
                    # Template update failed - notify user
                    sender = result[1]
                    error_details = result[2]
                    thread_info = result[3]
                    print(f"❌ Template update failed from {sender}: {error_details}")
                    
                    # Use threading for THREAD_TEMPLATE_UPDATE_FAILED
                    use_threading = thread_info if action_type == "THREAD_TEMPLATE_UPDATE_FAILED" else None
                    
                    send_template_confirmation(sender, success=False, thread_info=use_threading, error_details=error_details)
                    logging.error(f"Template update failed from {sender}: {error_details}")
                
                elif action_type in ["NORMAL_PROCESSING", "THREAD_NORMAL_PROCESSING"]:
                    # Regular pipeline processing
                    file_path = result[1]
                    sender = result[2]
                    thread_info = result[3]
                    print(f"📊 Processing pipeline file from: {sender}")
                    
                    try:
                        final_file = generate_gantt_chart(file_path)
                        print(f"📄 Generated file: {final_file}")
                        
                        # Use threading for THREAD_NORMAL_PROCESSING
                        use_threading = thread_info if action_type == "THREAD_NORMAL_PROCESSING" else None
                        
                        if send_successful_processing_email(sender, final_file, use_threading):
                            logging.info(f"Pipeline processed and sent successfully to {sender}")
                        else:
                            handle_error("Failed to send processed file", sender, thread_info)
                        
                    except Exception as e:
                        error_msg = f"Processing failed for {sender}: {str(e)}"
                        handle_error(error_msg, sender, thread_info)
                
                elif action_type == "THREAD_UNCLEAR":
                    # Thread message was unclear - send help
                    sender = result[1]
                    subject = result[2] 
                    thread_info = result[3]
                    print(f"❓ Unclear thread message from {sender}")
                    
                    if send_help_instructions(sender, thread_info):
                        logging.info(f"Unclear message help sent to {sender}")
                    else:
                        handle_error("Failed to send unclear message help", sender, thread_info)
                
                elif action_type == "HELP":
                    # Legacy help request
                    sender = result[1]
                    subject = result[2]
                    thread_info = result[3]
                    print(f"ℹ️ Help request from {sender}")
                    
                    if send_help_instructions(sender, None):  # Don't thread legacy help
                        logging.info(f"Help instructions sent to {sender}")
                    else:
                        handle_error("Failed to send help instructions", sender, None)
                        
        except Exception as e:
            # Handle system-level errors (IMAP connection, etc.)
            error_msg = f"System error: {str(e)}"
            handle_error(error_msg)
            if DEBUG_MODE:
                import traceback
                print(traceback.format_exc())

        time.sleep(Config.CHECK_INTERVAL_SECONDS)

if __name__ == "__main__":
    main()