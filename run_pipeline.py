import logging
import time
import sys
from pathlib import Path

# Add project root to Python path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "scripts"))

from config import Config
from scripts.fetch import download_latest_attachment
from scripts.process import generate_gantt_chart
from scripts.send_email import send_email_with_attachment

print("✅ Running the resilient pipeline script (v6 - Fixed Configuration)")

# Setup logging
Config.ensure_directories()
logging.basicConfig(
    filename=Config.LOGS_DIR / "pipeline_log.txt",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

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
        
        if not Config.TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"Template file not found at {Config.TEMPLATE_PATH}")
        
        send_email_with_attachment(
            to_address=Config.JOE_EMAIL,
            cc_address=Config.YOUR_EMAIL,
            subject=template_subject,
            body=template_body,
            attachment_path=str(Config.TEMPLATE_PATH)
        )
        
        print(f"📧 Template sent to Joe ({Config.JOE_EMAIL})")
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
            cc_address=Config.YOUR_EMAIL,
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

def send_error_alert_email(error_message, sender_email):
    """Send error alert email to you."""
    try:
        alert_subject = "🚨 URGENT: Salesforce Pipeline Automation Error"
        alert_body = f"""PIPELINE ERROR ALERT

TIME: {time.strftime('%Y-%m-%d %H:%M:%S')}
SENDER: {sender_email}
ERROR: {error_message}

The automated Salesforce pipeline system has encountered an error and needs your attention.

Joe has been automatically notified that there's an issue and that you're working on a fix.

Please check the system logs and resolve the issue as soon as possible.

---
This is an automated error alert from your pipeline system.
"""
        
        send_email_with_attachment(
            to_address=Config.YOUR_EMAIL,
            subject=alert_subject,
            body=alert_body,
            attachment_path=None
        )
        
        print(f"📧 Error alert sent to you ({Config.YOUR_EMAIL})")
        logging.info(f"Error alert email sent to {Config.YOUR_EMAIL}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send error alert email: {e}")
        logging.error(f"Error alert email failed: {e}")
        return False

def send_error_email_to_joe(error_message, sender_email):
    """Send error notification email to Joe."""
    try:
        error_subject = "⚠️ Salesforce Pipeline Automation - Technical Issue Detected"
        error_body = f"""Hi Joe,

I wanted to let you know that there was a technical issue with the automated Salesforce pipeline processing system.

ISSUE DETAILS:
- Time: {time.strftime('%Y-%m-%d %H:%M:%S')}
- Original sender: {sender_email}
- Error: {error_message}

NEXT STEPS:
I (Rohan) have been automatically notified and am already working on identifying and fixing the issue. I will have this resolved as quickly as possible and will follow up with you once the system is back online.

In the meantime, if you need immediate pipeline processing, please feel free to send me the file directly and I can process it manually.

Best regards,
Rohan's Automated Pipeline System

---
This is an automated error notification. If you have questions, please contact Rohan.Anand@mag.us directly.
"""
        
        send_email_with_attachment(
            to_address=Config.JOE_EMAIL,
            cc_address=Config.YOUR_EMAIL,
            subject=error_subject,
            body=error_body,
            attachment_path=None
        )
        
        print(f"📧 Error notification sent to Joe ({Config.JOE_EMAIL})")
        logging.info(f"Error notification email sent to Joe: {Config.JOE_EMAIL}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send error email to Joe: {e}")
        logging.error(f"Error email to Joe failed: {e}")
        return False

def handle_error(error_message, sender_email=None):
    """Handle errors by sending alert email to you and notification email to Joe."""
    print(f"🚨 PIPELINE ERROR: {error_message}")
    
    # Log the error
    logging.error(f"Pipeline error: {error_message} (sender: {sender_email})")
    
    # Send alert email to you
    send_error_alert_email(error_message, sender_email or "Unknown")
    
    # Send error email to Joe
    if sender_email:
        send_error_email_to_joe(error_message, sender_email)
    else:
        send_error_email_to_joe(error_message, "Unknown")

def main():
    """Main pipeline loop."""
    try:
        # Validate configuration before starting
        Config.validate_config()
        print("✅ Configuration validated successfully")
        
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        logging.error(f"Configuration error: {e}")
        return
    
    print(f"🔄 Starting pipeline monitoring (checking every {Config.CHECK_INTERVAL_SECONDS} seconds)")
    
    while True:
        print("🔍 Checking for new emails...")
        
        try:
            result = download_latest_attachment()
            if result:
                action_type = result[0]
                
                if action_type == "ADJUST_COLUMNS":
                    # Joe wants to adjust columns - send him the template
                    sender = result[1]
                    print(f"🔧 Column adjustment request from {sender}")
                    
                    if send_template_to_joe():
                        logging.info(f"Template adjustment request processed for {sender}")
                    else:
                        handle_error("Failed to send template to Joe", sender)
                
                elif action_type == "TEMPLATE_UPDATED":
                    # Joe sent back the updated template - confirm success
                    sender = result[1]
                    print(f"✅ Template successfully updated by {sender}")
                    
                    send_template_confirmation(sender, success=True)
                    logging.info(f"Template successfully updated by {sender}")
                
                elif action_type == "TEMPLATE_UPDATE_FAILED":
                    # Template update failed - notify Joe
                    sender = result[1]
                    error_details = result[2] if len(result) > 2 else "Unknown error"
                    print(f"❌ Template update failed from {sender}: {error_details}")
                    
                    send_template_confirmation(sender, success=False)
                    logging.error(f"Template update failed from {sender}: {error_details}")
                
                elif action_type == "NORMAL_PROCESSING":
                    # Regular pipeline processing
                    file_path, sender = result[1], result[2]
                    print(f"📊 Processing pipeline file from: {sender}")
                    
                    try:
                        final_file = generate_gantt_chart(file_path)
                        print(f"📄 Generated file: {final_file}")
                        
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
                        
                        print(f"📧 Sending processed pipeline to {sender} with CC to {Config.YOUR_EMAIL}")
                        
                        send_email_with_attachment(
                            to_address=sender,
                            cc_address=Config.YOUR_EMAIL,
                            subject="See the updated DEC & C5S Pipeline",
                            body=(
                                f"Hi {greeting_name},\n\n"
                                "I have successfully processed the Salesforce pipeline for you and made the necessary adjustments."
                                " Should you have any questions or identify any problems, please do not hesitate to reach out to me directly at Rohan.Anand@mag.us.\n\n"
                                "Best regards,\n"
                                "Rohan\n\n\n\n"
                            ),
                            attachment_path=final_file
                        )
                        logging.info(f"Pipeline processed and sent successfully to {sender}")
                        print("✅ Pipeline processing completed successfully.")
                        
                    except Exception as e:
                        error_msg = f"Processing failed for {sender}: {str(e)}"
                        handle_error(error_msg, sender)
                        
        except Exception as e:
            # Handle system-level errors (IMAP connection, etc.)
            error_msg = f"System error: {str(e)}"
            handle_error(error_msg)

        time.sleep(Config.CHECK_INTERVAL_SECONDS)

if __name__ == "__main__":
    main()