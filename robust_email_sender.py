#!/usr/bin/env python3
"""
Robust Email Sender with Progressive Threading Fallbacks
Implements the mitigation plan for reliable email threading
"""

import smtplib
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
import os
from pathlib import Path
import sys

# Add project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config import Config
from email_header_cleaner import EmailHeaderCleaner

class RobustEmailSender:
    """
    Email sender with progressive fallback strategies for threading.
    """
    
    def __init__(self, debug=False):
        self.debug = debug
        self.cleaner = EmailHeaderCleaner()
        
    def log(self, message, level="INFO"):
        """Log messages if debug mode is enabled."""
        if self.debug:
            print(f"[{level}] {message}")
    
    def send_email(self, to_address, subject, body, attachment_path=None, cc_address=None, 
                   thread_info=None, force_new_thread=False):
        """
        Send email with robust threading and fallback strategies.
        
        Args:
            to_address (str): Primary recipient
            subject (str): Email subject
            body (str): Email body
            attachment_path (str): Path to attachment file
            cc_address (str): CC recipient
            thread_info (dict): Threading information with keys:
                - message_id: Message-ID to reply to
                - references: References header chain
                - subject: Original subject
            force_new_thread (bool): Force sending as new thread (no threading headers)
            
        Returns:
            bool: True if sent successfully, False otherwise
        """
        try:
            self.log(f"Preparing email to {to_address}")
            self.log(f"Subject: {subject}")
            self.log(f"Threading enabled: {thread_info is not None and not force_new_thread}")
            
            # Validate configuration
            if not Config.EMAIL_USER or not Config.EMAIL_PASS:
                raise ValueError("Email credentials not configured")
            
            # Create message
            msg = EmailMessage()
            msg['From'] = Config.EMAIL_USER
            msg['To'] = to_address
            msg['Date'] = formatdate(localtime=True)
            msg['Message-ID'] = make_msgid()
            
            # Add CC if provided
            if cc_address:
                msg['Cc'] = cc_address
                self.log(f"CC: {cc_address}")
            
            # Handle threading with progressive fallbacks
            if thread_info and not force_new_thread:
                success = self._apply_threading_with_fallbacks(msg, subject, thread_info)
                if not success:
                    self.log("All threading strategies failed, sending as new thread", "WARNING")
            else:
                # New thread - just clean the subject
                msg['Subject'] = self.cleaner.clean_subject(subject)
                self.log(f"New thread subject: {msg['Subject']}")
            
            # Set body
            msg.set_content(body)
            
            # Add attachment if provided
            if attachment_path:
                self._add_attachment(msg, attachment_path)
            
            # Validate all headers before sending
            headers_dict = {key: value for key, value in msg.items()}
            is_valid, error_msg = self.cleaner.validate_headers(headers_dict)
            
            if not is_valid:
                self.log(f"Header validation failed: {error_msg}", "ERROR")
                return False
            
            # Send the email
            return self._send_smtp(msg)
            
        except Exception as e:
            self.log(f"Email sending failed: {e}", "ERROR")
            return False
    
    def _apply_threading_with_fallbacks(self, msg, subject, thread_info):
        """
        Apply threading headers with multiple fallback strategies.
        
        Returns:
            bool: True if threading was successfully applied
        """
        strategies = [
            self._strategy_full_threading,
            self._strategy_simple_threading,
            self._strategy_subject_only_threading
        ]
        
        for i, strategy in enumerate(strategies, 1):
            self.log(f"Trying threading strategy {i}/{len(strategies)}")
            try:
                if strategy(msg, subject, thread_info):
                    self.log(f"Threading strategy {i} succeeded")
                    return True
            except Exception as e:
                self.log(f"Threading strategy {i} failed: {e}", "WARNING")
                continue
        
        return False
    
    def _strategy_full_threading(self, msg, subject, thread_info):
        """
        Strategy 1: Full threading with In-Reply-To and References headers.
        """
        original_message_id = thread_info.get('message_id')
        original_references = thread_info.get('references')
        
        if not original_message_id:
            return False
        
        # Build clean threading headers
        threading_headers = self.cleaner.build_threading_headers(
            original_message_id, 
            original_references
        )
        
        if not threading_headers:
            return False
        
        # Apply headers
        for header_name, header_value in threading_headers.items():
            msg[header_name] = header_value
            self.log(f"Added {header_name}: {header_value}")
        
        # Set threaded subject
        original_subject = thread_info.get('subject', subject)
        clean_subject = self.cleaner.clean_subject(original_subject)
        
        if not clean_subject.startswith('Re:'):
            clean_subject = f"Re: {clean_subject}"
        
        msg['Subject'] = clean_subject
        self.log(f"Full threading subject: {clean_subject}")
        
        return True
    
    def _strategy_simple_threading(self, msg, subject, thread_info):
        """
        Strategy 2: Simple threading with just In-Reply-To header.
        """
        original_message_id = thread_info.get('message_id')
        
        if not original_message_id:
            return False
        
        clean_msg_id = self.cleaner.extract_message_id(original_message_id)
        
        if not clean_msg_id:
            return False
        
        # Only add In-Reply-To (skip References to avoid complexity)
        msg['In-Reply-To'] = clean_msg_id
        self.log(f"Simple threading - In-Reply-To: {clean_msg_id}")
        
        # Use consistent subject for threading
        msg['Subject'] = "Re: Interact with the MAG bot to configure your file"
        self.log(f"Simple threading subject: {msg['Subject']}")
        
        return True
    
    def _strategy_subject_only_threading(self, msg, subject, thread_info):
        """
        Strategy 3: Subject-only threading (Gmail will group by subject pattern).
        """
        # Use consistent subject that Gmail will group
        msg['Subject'] = "Re: Interact with the MAG bot to configure your file"
        self.log(f"Subject-only threading: {msg['Subject']}")
        
        return True
    
    def _add_attachment(self, msg, attachment_path):
        """Add file attachment to email message."""
        attachment_file = Path(attachment_path)
        
        if not attachment_file.exists():
            raise FileNotFoundError(f"Attachment not found: {attachment_path}")
        
        with open(attachment_file, 'rb') as f:
            file_data = f.read()
        
        # Determine MIME type
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
        
        self.log(f"Attachment added: {attachment_file.name} ({len(file_data)} bytes)")
    
    def _send_smtp(self, msg):
        """Send email via SMTP with error handling."""
        try:
            self.log(f"Connecting to {Config.SMTP_SERVER}:{Config.SMTP_PORT}")
            
            with smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT) as server:
                server.starttls()
                server.login(Config.EMAIL_USER, Config.EMAIL_PASS)
                server.send_message(msg)
            
            self.log("Email sent successfully!", "SUCCESS")
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            self.log(f"SMTP authentication failed: {e}", "ERROR")
            return False
        except smtplib.SMTPRecipientsRefused as e:
            self.log(f"Recipients refused: {e}", "ERROR")
            return False
        except Exception as e:
            self.log(f"SMTP error: {e}", "ERROR")
            return False

def send_email_with_attachment(to_address, subject, body, attachment_path=None, 
                               cc_address=None, thread_info=None, debug=False):
    """
    Convenience function that matches the original interface.
    """
    sender = RobustEmailSender(debug=debug)
    return sender.send_email(
        to_address=to_address,
        subject=subject,
        body=body,
        attachment_path=attachment_path,
        cc_address=cc_address,
        thread_info=thread_info
    )

def test_robust_sender():
    """Test the robust email sender with various scenarios."""
    print("🧪 TESTING ROBUST EMAIL SENDER")
    print("=" * 50)
    
    try:
        Config.validate_config()
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        return
    
    sender = RobustEmailSender(debug=True)
    
    # Test 1: New thread
    print("\n📧 Test 1: New thread")
    sender.send_email(
        to_address=Config.YOUR_EMAIL,
        subject="Robust Sender Test - New Thread",
        body="This is a test of the new robust email sender.",
        thread_info=None
    )
    
    # Test 2: Threading with clean headers
    print("\n📧 Test 2: Threading with clean headers")
    clean_thread_info = {
        'message_id': '<test123@example.com>',
        'references': None,
        'subject': 'Original Subject'
    }
    sender.send_email(
        to_address=Config.YOUR_EMAIL,
        subject="Robust Sender Test - Clean Threading",
        body="This is a test of threading with clean headers.",
        thread_info=clean_thread_info
    )
    
    # Test 3: Threading with problematic headers
    print("\n📧 Test 3: Threading with problematic headers")
    problematic_thread_info = {
        'message_id': '<SA1P110MB18627B2BAE15D8F41A61A072892AA@SA1P110\nMB18627.NAMP110.PROD.OUTLOOK.COM>',
        'references': '<msg1@example.com>\n<msg2@bad\rformatted.com>',
        'subject': 'Re: [External] Interact with the MAG bot to configure your file'
    }
    sender.send_email(
        to_address=Config.YOUR_EMAIL,
        subject="Robust Sender Test - Problematic Threading",
        body="This is a test of threading with problematic headers that should fallback gracefully.",
        thread_info=problematic_thread_info
    )
    
    print("\n✅ Robust sender tests completed!")

if __name__ == "__main__":
    test_robust_sender()