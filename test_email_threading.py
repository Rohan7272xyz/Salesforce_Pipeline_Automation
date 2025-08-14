#!/usr/bin/env python3
"""
Isolated Email Threading Test
Tests ONLY the email sending and threading functionality
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

def clean_header_value(value):
    """
    Aggressively clean email header values to remove problematic characters.
    """
    if not value:
        return ""
    
    # Convert to string and remove problematic characters
    cleaned = str(value)
    
    # Remove newlines, carriage returns, tabs
    cleaned = cleaned.replace('\n', '').replace('\r', '').replace('\t', '')
    
    # Remove any non-ASCII characters that might cause issues
    cleaned = ''.join(char for char in cleaned if ord(char) < 128)
    
    # Remove extra whitespace
    cleaned = ' '.join(cleaned.split())
    
    # Limit length to prevent issues
    cleaned = cleaned[:500]
    
    return cleaned

def extract_message_id(raw_message_id):
    """
    Extract clean Message-ID from potentially malformed header.
    """
    if not raw_message_id:
        return None
    
    # Clean the raw input
    cleaned = clean_header_value(raw_message_id)
    
    # Try to extract Message-ID between < and >
    if '<' in cleaned and '>' in cleaned:
        start = cleaned.find('<')
        end = cleaned.find('>', start)
        if start != -1 and end != -1:
            return cleaned[start:end+1]
    
    # If no brackets, just return cleaned version
    return cleaned

def test_send_threaded_email(to_address, subject, body, reply_to_message_id=None, debug=True):
    """
    Test function to send a threaded email reply.
    
    Args:
        to_address: Recipient email
        subject: Email subject
        body: Email body
        reply_to_message_id: Message-ID to reply to (for threading)
        debug: Print detailed debug info
    """
    try:
        if debug:
            print(f"🧪 TESTING EMAIL SEND:")
            print(f"   To: {to_address}")
            print(f"   Subject: {subject}")
            print(f"   Reply-To Message-ID: {reply_to_message_id}")
        
        # Create message
        msg = EmailMessage()
        msg['From'] = Config.EMAIL_USER
        msg['To'] = to_address
        msg['Date'] = formatdate(localtime=True)
        msg['Message-ID'] = make_msgid()
        
        # Handle threading
        if reply_to_message_id:
            # Clean the Message-ID
            clean_msg_id = extract_message_id(reply_to_message_id)
            
            if clean_msg_id:
                if debug:
                    print(f"   Original Message-ID: {reply_to_message_id[:50]}...")
                    print(f"   Cleaned Message-ID: {clean_msg_id}")
                
                # Set threading headers
                msg['In-Reply-To'] = clean_msg_id
                msg['References'] = clean_msg_id
                msg['Subject'] = f"Re: {subject}" if not subject.startswith('Re:') else subject
                
                if debug:
                    print(f"   ✅ Threading headers added")
            else:
                if debug:
                    print(f"   ⚠️ Could not extract clean Message-ID, sending without threading")
                msg['Subject'] = subject
        else:
            msg['Subject'] = subject
            if debug:
                print(f"   📧 New conversation (no threading)")
        
        msg.set_content(body)
        
        # Test: Print all headers before sending
        if debug:
            print(f"\n📋 FINAL EMAIL HEADERS:")
            for key, value in msg.items():
                # Check for problematic characters
                has_newlines = '\n' in str(value) or '\r' in str(value)
                status = "❌ HAS NEWLINES" if has_newlines else "✅ CLEAN"
                print(f"   {key}: {str(value)[:100]}... {status}")
        
        # Connect and send
        print(f"\n🔗 Connecting to {Config.SMTP_SERVER}:{Config.SMTP_PORT}")
        
        with smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT) as server:
            server.starttls()
            server.login(Config.EMAIL_USER, Config.EMAIL_PASS)
            server.send_message(msg)
        
        print("✅ Email sent successfully!")
        return True
        
    except Exception as e:
        print(f"❌ Email send failed: {e}")
        return False

def test_header_cleaning():
    """Test header cleaning with various problematic inputs."""
    print("🧪 TESTING HEADER CLEANING:")
    
    test_cases = [
        "<SA1P110MB18627B2BAE15D8F41A61A072892AA@SA1P110MB18627.NAMP110.PROD.OUTLOOK.COM>",
        "<SA1P110MB18627B2BAE15D8F41A61A072892AA@SA1P110\nMB18627.NAMP110.PROD.OUTLOOK.COM>",
        "  <SA1P110MB18627B2BAE15D8F41A61A072892AA@SA1P110\r\nMB18627.NAMP110.PROD.OUTLOOK.COM>  ",
        "malformed-message-id-without-brackets",
        "",
        None
    ]
    
    for i, test_input in enumerate(test_cases, 1):
        print(f"\n   Test {i}: {repr(test_input)}")
        cleaned = extract_message_id(test_input)
        print(f"   Result: {repr(cleaned)}")
        
        # Check for problematic characters
        if cleaned:
            has_issues = '\n' in cleaned or '\r' in cleaned or '\t' in cleaned
            print(f"   Status: {'❌ HAS ISSUES' if has_issues else '✅ CLEAN'}")

if __name__ == "__main__":
    try:
        Config.validate_config()
        print("✅ Configuration validated")
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        sys.exit(1)
    
    print("=" * 60)
    print("EMAIL THREADING ISOLATION TEST")
    print("=" * 60)
    
    # Test 1: Header cleaning
    test_header_cleaning()
    
    print("\n" + "=" * 60)
    
    # Test 2: Send simple email without threading
    print("🧪 TEST 2: New email (no threading)")
    test_send_threaded_email(
        to_address=Config.YOUR_EMAIL,
        subject="Test Email Threading - New Conversation",
        body="This is a test email to verify basic sending works.",
        reply_to_message_id=None,
        debug=True
    )
    
    print("\n" + "=" * 60)
    
    # Test 3: Send email with threading (using a clean Message-ID)
    print("🧪 TEST 3: Threaded reply (clean Message-ID)")
    clean_message_id = "<test123@example.com>"
    test_send_threaded_email(
        to_address=Config.YOUR_EMAIL,
        subject="Test Email Threading - Reply",
        body="This is a test threaded reply with clean headers.",
        reply_to_message_id=clean_message_id,
        debug=True
    )
    
    print("\n" + "=" * 60)
    
    # Test 4: Send email with threading (using problematic Message-ID)
    print("🧪 TEST 4: Threaded reply (problematic Message-ID)")
    problematic_message_id = "<SA1P110MB18627B2BAE15D8F41A61A072892AA@SA1P110\nMB18627.NAMP110.PROD.OUTLOOK.COM>"
    test_send_threaded_email(
        to_address=Config.YOUR_EMAIL,
        subject="Test Email Threading - Problematic Headers",
        body="This is a test threaded reply with cleaned problematic headers.",
        reply_to_message_id=problematic_message_id,
        debug=True
    )
    
    print("\n" + "=" * 60)
    print("✅ Email threading tests completed!")
    print("Check your email to verify which tests worked.")