#!/usr/bin/env python3
"""
Robust Email Header Cleaning Module
Handles all the edge cases for cleaning email headers to prevent SMTP errors
"""

import re
import unicodedata

class EmailHeaderCleaner:
    """
    Comprehensive email header cleaning with multiple fallback strategies.
    """
    
    @staticmethod
    def clean_subject(subject):
        """
        Clean email subject line for safe use in headers.
        
        Args:
            subject (str): Raw subject line
            
        Returns:
            str: Cleaned subject line safe for email headers
        """
        if not subject:
            return "Interact with the MAG bot to configure your file"
        
        # Convert to string and handle Unicode
        subject = str(subject)
        
        # Normalize Unicode characters
        subject = unicodedata.normalize('NFKD', subject)
        
        # Remove control characters (including newlines, carriage returns, tabs)
        subject = ''.join(char for char in subject if unicodedata.category(char)[0] != 'C')
        
        # Collapse multiple whitespace into single spaces
        subject = re.sub(r'\s+', ' ', subject)
        
        # Trim whitespace
        subject = subject.strip()
        
        # Limit length (email headers should be reasonable length)
        if len(subject) > 200:
            subject = subject[:197] + "..."
        
        # Handle Re: prefix properly
        if subject and not subject.startswith('Re:'):
            # Only add Re: if this appears to be a reply
            if any(keyword in subject.lower() for keyword in ['re:', 'reply', 'response']):
                if not subject.startswith('Re:'):
                    subject = f"Re: {subject}"
        
        return subject or "Interact with the MAG bot to configure your file"
    
    @staticmethod
    def extract_message_id(raw_message_id):
        """
        Extract and clean Message-ID from email headers.
        
        Args:
            raw_message_id (str): Raw Message-ID header value
            
        Returns:
            str or None: Cleaned Message-ID or None if invalid
        """
        if not raw_message_id:
            return None
        
        # Convert to string
        raw_message_id = str(raw_message_id)
        
        # Remove all control characters
        cleaned = ''.join(char for char in raw_message_id if unicodedata.category(char)[0] != 'C')
        
        # Remove extra whitespace
        cleaned = re.sub(r'\s+', '', cleaned)
        
        # Message-ID should be in format <localpart@domain>
        # Extract content between < and >
        match = re.search(r'<([^<>]+)>', cleaned)
        if match:
            message_id_content = match.group(1)
            
            # Validate that it looks like a proper Message-ID
            if '@' in message_id_content and '.' in message_id_content:
                # Reconstruct with clean brackets
                return f"<{message_id_content}>"
        
        # If no valid Message-ID found, return None
        return None
    
    @staticmethod
    def clean_references(raw_references):
        """
        Clean References header value.
        
        Args:
            raw_references (str): Raw References header value
            
        Returns:
            str or None: Cleaned References header or None if invalid
        """
        if not raw_references:
            return None
        
        # Convert to string and remove control characters
        raw_references = str(raw_references)
        cleaned = ''.join(char for char in raw_references if unicodedata.category(char)[0] != 'C')
        
        # References is a space-separated list of Message-IDs
        # Extract all valid Message-IDs
        message_ids = re.findall(r'<[^<>]+@[^<>]+>', cleaned)
        
        if message_ids:
            # Return space-separated list of clean Message-IDs
            return ' '.join(message_ids)
        
        return None
    
    @staticmethod
    def build_threading_headers(original_message_id, original_references=None):
        """
        Build clean threading headers for email replies.
        
        Args:
            original_message_id (str): Message-ID we're replying to
            original_references (str): Existing References chain
            
        Returns:
            dict: Clean threading headers or empty dict if invalid
        """
        threading_headers = {}
        
        # Clean the Message-ID we're replying to
        clean_msg_id = EmailHeaderCleaner.extract_message_id(original_message_id)
        
        if clean_msg_id:
            # Set In-Reply-To
            threading_headers['In-Reply-To'] = clean_msg_id
            
            # Build References chain
            references = []
            
            # Add existing references if valid
            clean_refs = EmailHeaderCleaner.clean_references(original_references)
            if clean_refs:
                references.extend(clean_refs.split())
            
            # Add the message we're replying to (if not already in references)
            if clean_msg_id not in references:
                references.append(clean_msg_id)
            
            # Limit references to reasonable number (RFC recommends max ~20)
            if len(references) > 20:
                references = references[-20:]
            
            threading_headers['References'] = ' '.join(references)
        
        return threading_headers
    
    @staticmethod
    def validate_headers(headers_dict):
        """
        Validate that all headers are safe for SMTP transmission.
        
        Args:
            headers_dict (dict): Dictionary of header name -> value
            
        Returns:
            tuple: (is_valid, error_message)
        """
        for header_name, header_value in headers_dict.items():
            if not header_value:
                continue
            
            header_str = str(header_value)
            
            # Check for control characters that cause SMTP errors
            for char in header_str:
                if unicodedata.category(char)[0] == 'C' and char not in ['\t']:
                    return False, f"Header '{header_name}' contains control character: {repr(char)}"
            
            # Check for excessively long headers
            if len(header_str) > 1000:
                return False, f"Header '{header_name}' is too long: {len(header_str)} characters"
        
        return True, "Headers are valid"

def test_header_cleaner():
    """Test the header cleaning functionality."""
    print("🧪 TESTING EMAIL HEADER CLEANER")
    print("=" * 50)
    
    # Test cases with various problematic inputs
    test_cases = [
        {
            'name': 'Clean Message-ID',
            'message_id': '<clean123@example.com>',
            'references': None
        },
        {
            'name': 'Outlook Message-ID with newlines',
            'message_id': '<SA1P110MB18627B2BAE15D8F41A61A072892AA@SA1P110\nMB18627.NAMP110.PROD.OUTLOOK.COM>',
            'references': None
        },
        {
            'name': 'Complex References chain',
            'message_id': '<msg3@example.com>',
            'references': '<msg1@example.com> <msg2@example.com>'
        },
        {
            'name': 'Problematic References with newlines',
            'message_id': '<msg2@example.com>',
            'references': '<msg1@example.com>\n<msg2@bad\rformatted.com>'
        },
        {
            'name': 'Invalid Message-ID (no brackets)',
            'message_id': 'invalid-message-id',
            'references': None
        }
    ]
    
    cleaner = EmailHeaderCleaner()
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n🔍 Test {i}: {test_case['name']}")
        print(f"   Input Message-ID: {repr(test_case['message_id'])}")
        print(f"   Input References: {repr(test_case['references'])}")
        
        # Test Message-ID cleaning
        clean_msg_id = cleaner.extract_message_id(test_case['message_id'])
        print(f"   Cleaned Message-ID: {repr(clean_msg_id)}")
        
        # Test References cleaning
        clean_refs = cleaner.clean_references(test_case['references'])
        print(f"   Cleaned References: {repr(clean_refs)}")
        
        # Test building threading headers
        threading_headers = cleaner.build_threading_headers(
            test_case['message_id'], 
            test_case['references']
        )
        print(f"   Threading Headers: {threading_headers}")
        
        # Validate the headers
        is_valid, error_msg = cleaner.validate_headers(threading_headers)
        print(f"   Validation: {'✅ VALID' if is_valid else '❌ INVALID'} - {error_msg}")
    
    # Test subject cleaning
    print(f"\n🔍 Subject Line Cleaning Tests:")
    subject_tests = [
        "Start Conversation",
        "Re: [External] Interact with the MAG bot to configure your file",
        "Subject with\nnewlines and\ttabs",
        "Very long subject line that exceeds normal length limits and should be truncated properly to prevent issues with email clients and servers that have header length restrictions",
        "Subject with émojis and special characters: 🤖📧✅",
        ""
    ]
    
    for subject in subject_tests:
        cleaned = cleaner.clean_subject(subject)
        print(f"   '{subject}' → '{cleaned}'")

if __name__ == "__main__":
    test_header_cleaner()