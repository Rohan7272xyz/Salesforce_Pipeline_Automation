#!/usr/bin/env python3
"""
Production Pipeline with Multi-User Authorization and Enhanced Error Handling
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

print("✅ Running PRODUCTION pipeline with Multi-User Support (v9)")
print(f"🐛 Debug mode: {'ON' if DEBUG_MODE else 'OFF'}")
print(f"🔗 Threading: {'ENABLED' if THREADING_ENABLED else 'DISABLED'}")
print(f"👥 Authorized users: {len(Config.AUTHORIZED_EMAILS)}")
for email in Config.AUTHORIZED_EMAILS:
    print(f"   ✅ {email}")

# Setup logging
Config.ensure_directories()
logging.basicConfig(
    filename=Config.LOGS_DIR / "pipeline_log.txt",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# Initialize robust email sender
email_sender = RobustEmailSender(debug=DEBUG_MODE)

def get_user_greeting(sender_email):
    """Get personalized greeting based on sender email."""
    sender_lower = sender_email.lower()
    
    if "joseph.findley" in sender_lower or "joe" in sender_lower:
        return "Joe"
    elif "rohan" in sender_lower:
        return "Rohan"
    elif "person1" in sender_lower:
        return "Person1"  # Customize as needed
    elif "person2" in sender_lower:
        return "Person2"  # Customize as needed
    else:
        # Extract first name from email or use generic greeting
        name_part = sender_email.split('@')[0].split('.')[0]
        return name_part.title() if name_part else "there"

def send_help_instructions(sender, thread_info=None):
    """Send help/instructions to user when they start a conversation."""
    try:
        greeting_name = get_user_greeting(sender)
        
        help_body = f"""MAG Pipeline Automation Assistant

Hello {greeting_name}! 👋

Welcome to your intelligent Salesforce pipeline processing system. I'm here to streamline your workflow and ensure your data is always perfectly formatted.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚀 WHAT I CAN DO FOR YOU

📊 Pipeline Processing
   Simply attach your Salesforce pipeline Excel file to any email, and I'll:
   ✅ Clean and organize your data
   ✅ Apply consistent formatting  
   ✅ Generate professional Gantt chart layouts
   ✅ Return a polished, presentation-ready file

🔧 Template Management
   Need to adjust column structures? Just type "Change Format" and I'll:
   📤 Send you the current template for modification
   🔄 Guide you through the update process
   ✅ Automatically integrate your changes

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 GETTING STARTED

For Pipeline Processing: Attach your Excel file to this email thread
For Template Changes: Reply with "Change Format"  
Need Help: Type "Help" anytime

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔗 All conversations stay in this email thread for easy tracking
👥 Available to all authorized MAG team members

Best regards,
MAG Pipeline Bot
Your Automated Pipeline Assistant

---
Type "Help" anytime to see these instructions again
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        # CC all error recipients for transparency (with safety check)
        cc_list = Config.get_error_cc_list()
        # Remove sender from CC list to avoid duplicates
        cc_list = [email for email in cc_list if email.lower() != sender.lower()]
        # Use first available CC or fallback to primary admin
        cc_address = cc_list[0] if cc_list else Config.YOUR_EMAIL
        
        success = email_sender.send_email(
            to_address=sender,
            subject="Interact with the MAG bot to configure your file",
            body=help_body,
            cc_address=cc_address,
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
        greeting_name = get_user_greeting(sender)
        
        template_body = f"""📋 Template Customization Request

Hello {greeting_name}!

I've attached the current Excel template that powers your pipeline automation system. You can now customize it to match your exact requirements.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔧 STEP-BY-STEP INSTRUCTIONS

Step 1: 📥 Download & Open
   Download the attached template file and open it in Excel

Step 2: ✏️ Customize  
   Make your desired column/format changes:
   • Add, remove, or reorder columns
   • Modify headers and formatting
   • Adjust any layout preferences

Step 3: 💾 Save Properly
   Save the file in .xlsx format (Excel format)

Step 4: 📤 Return Updated Template
   Reply to this email thread with:
   ✅ Your updated template file attached
   ✅ The word "Here" in the message body

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚠️ IMPORTANT NOTES

🔹 Reply to this thread (don't start a new email)  
🔹 Type "Here" in the message body (this triggers the update)  
🔹 Attach your modified template as .xlsx file  
🔹 Keep original structure where possible for best results

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📧 Once you reply with "Here" + attachment, I'll automatically:
   • Update the system template
   • Confirm the changes  
   • Ready the system for processing with your new format

Best regards,
MAG Pipeline Bot
Template Management System

---
Automated response to your "Change Format" request
"""
        
        if not Config.TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"Template file not found at {Config.TEMPLATE_PATH}")
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        # CC admins for transparency
        cc_list = Config.get_error_cc_list()
        cc_list = [email for email in cc_list if email.lower() != sender.lower()]
        cc_address = cc_list[0] if cc_list else None
        
        success = email_sender.send_email(
            to_address=sender,
            subject="Template for Column Adjustment",
            body=template_body,
            attachment_path=str(Config.TEMPLATE_PATH),
            cc_address=cc_address,
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
        greeting_name = get_user_greeting(sender)
        
        if success:
            body = f"""✅ Template Update Successful!

Hello {greeting_name}!

Great news! Your template customization has been successfully processed and integrated into the system.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎉 WHAT JUST HAPPENED

✅ Template Updated - Your new format is now active  
✅ System Reconfigured - All processing rules updated automatically  
✅ Backup Created - Previous template safely archived  
✅ Ready for Processing - System is live with your changes

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚀 YOU'RE ALL SET!

Your pipeline automation system is now configured with your custom template and ready to process files using your new formatting structure.

Next Steps:
📎 Attach your Salesforce pipeline Excel files to this email thread  
⚡ I'll process them automatically with your new format  
📊 Receive your polished, formatted reports back via email

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔄 NEED CHANGES?

If you need to make additional template adjustments:
- Type "Change Format" to start the process again
- I'll send you the current (updated) template to modify

Best regards,
MAG Pipeline Bot
Template Management System

🎯 Your automation system is now optimized and ready for action!

---
Automated confirmation of successful template update
"""
        else:
            body = f"""There was an issue updating the template. ❌

ERROR DETAILS
{error_details or 'Unknown error occurred'}

Hello {greeting_name}!

I encountered an issue while trying to update your template. Don't worry - this is easily fixable!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔍 QUICK TROUBLESHOOTING

Please check the following and try again:

✅ Message Format
   • Type "Here" at the very beginning of your message
   • Don't include any other text before "Here"
   • The word "Here" triggers the update process

📎 File Requirements
   • Attach your Excel file (.xlsx format)
   • Ensure the file isn't corrupted or password-protected
   • File should be saved properly in Excel format

📧 Email Thread
   • Reply to this email thread (don't create new email)
   • Keep the conversation in the same thread

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🆘 STILL HAVING ISSUES?

If you continue experiencing problems after following these steps:

📧 Contact: {', '.join(Config.get_error_cc_list())}
📞 For: Direct technical support

Best regards,
MAG Pipeline Bot
Template Management System

🔧 I'm here to help - let's get your template updated successfully!

---
Automated error notification - template update failed
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        subject = "Template Update Confirmation" if success else "Template Update Failed"
        
        # CC all error recipients
        cc_list = Config.get_error_cc_list()
        cc_list = [email for email in cc_list if email.lower() != sender.lower()]
        cc_address = cc_list[0] if cc_list else None
        
        email_success = email_sender.send_email(
            to_address=sender,
            subject=subject,
            body=body,
            cc_address=cc_address,
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
        greeting_name = get_user_greeting(sender)
        
        success_body = f"""✅ Pipeline Processing Complete!

Hello {greeting_name}!

Your Salesforce pipeline has been successfully processed and is ready for use. I've transformed your raw data into a polished, presentation-ready format.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 PROCESSING SUMMARY

✅ Data Optimization
   • Cleaned & validated all data entries
   • Sorted by Capture Manager for logical organization
   • Applied formatting rules for consistency and readability
   • Removed duplicates and corrected inconsistencies

📊 Visual Enhancements
   • Generated Gantt Chart view for timeline visualization
   • Professional layout with proper spacing and alignment
   • Consistent styling across all columns and sections
   • Export-ready format for presentations and reports

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📎 YOUR PROCESSED FILE

📁 Attached: Your formatted pipeline report  
🎨 Format: Professional Excel with Gantt chart layout  
📊 Status: Ready for presentations, analysis, and stakeholder review

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚀 WHAT'S NEXT?

Ready to Use:
   • Open the attached file to review your formatted pipeline
   • Share with stakeholders for decision-making
   • Use for project planning and resource allocation

Need More Processing:
   • Send additional pipeline files anytime
   • I'll process them with the same high-quality standards
   • All processing happens automatically in this email thread

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 QUALITY ASSURANCE

✅ Data Integrity - All original data preserved and enhanced  
✅ Professional Standards - Corporate-ready formatting applied  
✅ Gantt Visualization - Clear timeline and milestone tracking  
✅ Stakeholder Ready - Polished presentation format

Best regards,
MAG Pipeline Bot
Your Automated Pipeline Specialist

🎊 Another successful pipeline transformation complete! Ready for your next file.

---
Pipeline processed at {time.strftime('%Y-%m-%d %H:%M:%S')}
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        # CC admins for transparency
        cc_list = Config.get_error_cc_list()
        cc_list = [email for email in cc_list if email.lower() != sender.lower()]
        cc_address = cc_list[0] if cc_list else None
        
        success = email_sender.send_email(
            to_address=sender,
            subject="Your Processed Salesforce Pipeline",
            body=success_body,
            attachment_path=final_file,
            cc_address=cc_address,
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
    """Send error alert email to ALL error recipients."""
    try:
        alert_subject = "🚨 URGENT: Salesforce Pipeline Automation Error"
        alert_body = f"""🚨 URGENT: Pipeline System Alert

System Administrators,

The MAG Salesforce Pipeline Automation system has encountered a critical error requiring immediate attention.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚠️ INCIDENT SUMMARY

🕐 Time: {time.strftime('%Y-%m-%d %H:%M:%S')}  
👤 Affected User: {sender_email}  
📍 Error Location: Pipeline Processing Module  
⚡ Status: User Automatically Notified

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔍 ERROR DETAILS

Error Message:
{error_message}

Impact Level: Service Interruption  
User Experience: Processing request failed  
System Status: Requires manual intervention

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ AUTOMATIC ACTIONS TAKEN

🔹 User Notification - Professional error message sent to user  
🔹 Error Logging - Full details captured in system logs  
🔹 Thread Preservation - Email conversation maintained  
🔹 System Stability - Core service remains operational for other users

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🛠️ REQUIRED ACTIONS

Immediate (< 30 minutes):
1. Check System Logs - Review pipeline_log.txt for detailed error trace
2. Assess Impact - Determine if this affects other users
3. Initial Diagnosis - Identify probable cause

Next Steps (< 2 hours):
1. Implement Fix - Apply necessary corrections
2. Test Resolution - Verify system functionality
3. User Follow-up - Notify user when system is restored

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📊 SYSTEM MONITORING

📁 Log Files: logs/pipeline_log.txt  
🔧 Configuration: Review recent template updates  
📧 Email Queue: Check for pending notifications  
🖥️ System Health: Monitor for additional errors

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📞 ESCALATION PATH

If resolution requires extended time:
- Update user with realistic timeline
- Consider manual processing as temporary workaround
- Document lessons learned for system improvement

This is an automated alert from your MAG Pipeline Automation System  
🤖 System Monitor | 📧 Auto-Generated Alert

⏰ Response Time Target: < 2 hours | 🎯 Resolution Priority: HIGH

---
Sent to all error recipients: {', '.join(Config.get_error_cc_list())}
"""
        
        # Send to ALL error recipients
        success_count = 0
        for admin_email in Config.get_error_cc_list():
            try:
                success = email_sender.send_email(
                    to_address=admin_email,
                    subject=alert_subject,
                    body=alert_body,
                    thread_info=None  # Always send errors as new threads
                )
                
                if success:
                    success_count += 1
                    print(f"📧 Error alert sent to admin ({admin_email})")
                    logging.info(f"Error alert email sent to {admin_email}")
                    
            except Exception as e:
                print(f"❌ Failed to send error alert to {admin_email}: {e}")
                logging.error(f"Error alert email failed for {admin_email}: {e}")
        
        return success_count > 0
        
    except Exception as e:
        print(f"❌ Failed to send error alert emails: {e}")
        logging.error(f"Error alert email system failed: {e}")
        return False

def send_error_email_to_user(error_message, sender_email, thread_info=None):
    """Send error notification email to user."""
    try:
        greeting_name = get_user_greeting(sender_email)
        
        error_body = f"""⚠️ Temporary Processing Issue

Hello {greeting_name}!

I wanted to personally let you know that I encountered a technical issue while processing your request. I sincerely apologize for any inconvenience this may cause.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔍 ISSUE SUMMARY

📅 Time: {time.strftime('%Y-%m-%d %H:%M:%S')}  
⚙️ Issue: Technical processing error  
📊 Your Data: Safely preserved and will be processed once resolved

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🛠️ RESOLUTION IN PROGRESS

✅ Immediate Action Taken
   • System administrators have been automatically notified
   • Error details have been logged for rapid diagnosis
   • System monitoring is active to prevent further issues

🔄 What's Happening Now
   • Technical team is investigating the root cause
   • Fix is being developed and tested
   • System will be restored as quickly as possible

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⏰ EXPECTED TIMELINE

🎯 Target Resolution: Within 2 hours  
📧 Update Frequency: You'll be notified when system is restored  
🔄 Auto-Retry: Your request will be processed automatically once fixed

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🆘 NEED IMMEDIATE ASSISTANCE?

If you have urgent pipeline processing needs that cannot wait:

📧 Direct Contact: {', '.join(Config.get_error_cc_list())}
📞 For: Immediate manual processing or technical support  
⚡ OR Call: (+1)202-961-1540

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 WHAT YOU CAN DO

🔹 Keep this email thread - I'll notify you when the system is restored  
🔹 No need to resend - Your original request is queued for processing  
🔹 Continue normal workflow - This is an isolated technical issue

Thank you for your patience while we resolve this quickly.

Best regards,
MAG Pipeline Bot
Pipeline Automation System

🔧 Committed to reliable service - we'll have this fixed shortly!

---
Automated error notification
"""
        
        # Use threading if enabled and thread_info provided
        use_threading = thread_info if THREADING_ENABLED else None
        
        # CC admins on error notifications
        cc_list = Config.get_error_cc_list()
        cc_list = [email for email in cc_list if email.lower() != sender_email.lower()]
        cc_address = cc_list[0] if cc_list else None
        
        success = email_sender.send_email(
            to_address=sender_email,
            subject="Pipeline Processing Error",
            body=error_body,
            cc_address=cc_address,
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
    """Handle errors by sending alert email to admins and notification email to user."""
    print(f"🚨 PIPELINE ERROR: {error_message}")
    
    # Log the error
    logging.error(f"Pipeline error: {error_message} (sender: {sender_email})")
    
    # Send alert email to ALL admins
    send_error_alert_email(error_message, sender_email or "Unknown")
    
    # Send error email to user (with threading if available)
    if sender_email:
        send_error_email_to_user(error_message, sender_email, thread_info)

def main():
    """Main pipeline loop with multi-user support."""
    try:
        # Validate configuration before starting
        Config.validate_config()
        print("✅ Configuration validated successfully")
        
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        logging.error(f"Configuration error: {e}")
        return
    
    print(f"🔄 Starting pipeline monitoring (checking every {Config.CHECK_INTERVAL_SECONDS} seconds)")
    print("👥 Multi-User Support: ACTIVE")
    print(f"📧 Error recipients: {', '.join(Config.get_error_cc_list())}")
    
    while True:
        # Only show detailed checking message occasionally in debug mode
        if DEBUG_MODE:
            check_count = getattr(main, 'check_count', 0) + 1
            main.check_count = check_count
            if check_count % 12 == 1:  # Every minute (12 * 5 seconds)
                print("🔍 Monitoring for emails... (showing every 60 seconds)")
        
        try:
            print("🔍 DEBUG: About to call download_latest_attachment()")
            result = download_latest_attachment()
            print(f"🔍 DEBUG: download_latest_attachment() returned: {result}")
            print(f"🔍 DEBUG: Result type: {type(result)}")
            
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
                    # User wants to Change Format
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