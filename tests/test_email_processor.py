"""
Test suite for email_processor module.

This module contains comprehensive unit tests for the EmailProcessor class,
including tests for email item processing, metadata extraction, and error handling.
"""

import unittest
from unittest.mock import Mock, MagicMock, patch
from datetime import datetime
import pandas as pd
import sys
import os

# Add src directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from outlook2ai.core.email_processor import EmailProcessor


class TestEmailProcessor(unittest.TestCase):
    """Test cases for EmailProcessor class."""
    
    def setUp(self):
        """Set up test fixtures before each test method."""
        self.processor = EmailProcessor()
        
        # Create mock email item
        self.mock_email = Mock()
        self.mock_email.Subject = "Test Subject"
        self.mock_email.Body = "Test body content"
        self.mock_email.HTMLBody = "<html><body>Test HTML content</body></html>"
        self.mock_email.ReceivedTime = datetime(2024, 1, 15, 10, 30, 0)
        self.mock_email.SentOn = datetime(2024, 1, 15, 10, 25, 0)
        self.mock_email.Size = 1024
        self.mock_email.UnRead = False
        self.mock_email.Importance = 2  # Normal importance
        self.mock_email.Sensitivity = 0  # Normal sensitivity
        self.mock_email.SenderName = "John Doe"
        self.mock_email.SenderEmailAddress = "john.doe@example.com"
        self.mock_email.ConversationID = "conversation123"
        self.mock_email.ConversationTopic = "Test Conversation"
        self.mock_email.EntryID = "entry123"
        self.mock_email.Categories = "Category1; Category2"
        
        # Mock Recipients collection
        mock_recipient1 = Mock()
        mock_recipient1.Name = "Jane Smith"
        mock_recipient1.Address = "jane.smith@example.com"
        mock_recipient1.Type = 1  # TO
        
        mock_recipient2 = Mock()
        mock_recipient2.Name = "Bob Johnson"
        mock_recipient2.Address = "bob.johnson@example.com"
        mock_recipient2.Type = 2  # CC
        
        self.mock_email.Recipients = [mock_recipient1, mock_recipient2]
        
        # Mock Attachments collection
        mock_attachment = Mock()
        mock_attachment.FileName = "document.pdf"
        mock_attachment.Size = 2048
        mock_attachment.Type = 1  # File attachment
        
        self.mock_email.Attachments = [mock_attachment]
    
    def test_process_email_item_success(self):
        """Test successful processing of an email item."""
        result = self.processor.process_email_item(self.mock_email)
        
        # Verify basic fields
        self.assertEqual(result['subject'], "Test Subject")
        self.assertEqual(result['body'], "Test body content")
        self.assertEqual(result['html_body'], "<html><body>Test HTML content</body></html>")
        self.assertEqual(result['sender_name'], "John Doe")
        self.assertEqual(result['sender_email'], "john.doe@example.com")
        self.assertFalse(result['unread'])
        self.assertEqual(result['size'], 1024)
        self.assertEqual(result['importance'], 2)
        self.assertEqual(result['sensitivity'], 0)
        
        # Verify datetime fields
        self.assertEqual(result['received_time'], datetime(2024, 1, 15, 10, 30, 0))
        self.assertEqual(result['sent_time'], datetime(2024, 1, 15, 10, 25, 0))
        
        # Verify recipient processing
        self.assertIn('jane.smith@example.com', result['to_recipients'])
        self.assertIn('bob.johnson@example.com', result['cc_recipients'])
        
        # Verify attachment processing
        self.assertEqual(result['attachment_count'], 1)
        self.assertIn('document.pdf', result['attachment_names'])
        
    def test_process_email_item_minimal_fields(self):
        """Test processing email with minimal required fields."""
        minimal_email = Mock()
        minimal_email.Subject = "Minimal Subject"
        minimal_email.Body = "Minimal body"
        minimal_email.HTMLBody = ""
        minimal_email.ReceivedTime = datetime.now()
        minimal_email.SentOn = datetime.now()
        minimal_email.Size = 100
        minimal_email.UnRead = True
        minimal_email.Importance = 1
        minimal_email.Sensitivity = 0
        minimal_email.SenderName = ""
        minimal_email.SenderEmailAddress = ""
        minimal_email.ConversationID = ""
        minimal_email.ConversationTopic = ""
        minimal_email.EntryID = ""
        minimal_email.Categories = ""
        minimal_email.Recipients = []
        minimal_email.Attachments = []
        
        result = self.processor.process_email_item(minimal_email)
        
        self.assertEqual(result['subject'], "Minimal Subject")
        self.assertEqual(result['body'], "Minimal body")
        self.assertEqual(result['attachment_count'], 0)
        self.assertEqual(len(result['to_recipients']), 0)
        self.assertEqual(len(result['cc_recipients']), 0)
        self.assertEqual(len(result['bcc_recipients']), 0)
    
    def test_process_recipients_all_types(self):
        """Test processing recipients of all types (TO, CC, BCC)."""
        recipients = []
        
        # TO recipient
        to_recipient = Mock()
        to_recipient.Name = "To Person"
        to_recipient.Address = "to@example.com"
        to_recipient.Type = 1
        recipients.append(to_recipient)
        
        # CC recipient
        cc_recipient = Mock()
        cc_recipient.Name = "CC Person"
        cc_recipient.Address = "cc@example.com"
        cc_recipient.Type = 2
        recipients.append(cc_recipient)
        
        # BCC recipient
        bcc_recipient = Mock()
        bcc_recipient.Name = "BCC Person"
        bcc_recipient.Address = "bcc@example.com"
        bcc_recipient.Type = 3
        recipients.append(bcc_recipient)
        
        to_list, cc_list, bcc_list = self.processor._process_recipients(recipients)
        
        self.assertIn("to@example.com", to_list)
        self.assertIn("cc@example.com", cc_list)
        self.assertIn("bcc@example.com", bcc_list)
    
    def test_process_recipients_missing_address(self):
        """Test processing recipients with missing email addresses."""
        recipients = []
        
        # Recipient with missing address
        recipient = Mock()
        recipient.Name = "No Email Person"
        recipient.Address = ""
        recipient.Type = 1
        recipients.append(recipient)
        
        to_list, cc_list, bcc_list = self.processor._process_recipients(recipients)
        
        # Should skip recipients without email addresses
        self.assertEqual(len(to_list), 0)
    
    def test_process_attachments_multiple_types(self):
        """Test processing different types of attachments."""
        attachments = []
        
        # File attachment
        file_attachment = Mock()
        file_attachment.FileName = "document.docx"
        file_attachment.Size = 1024
        file_attachment.Type = 1
        attachments.append(file_attachment)
        
        # Embedded message
        embedded_msg = Mock()
        embedded_msg.FileName = "FW: Message"
        embedded_msg.Size = 2048
        embedded_msg.Type = 5
        attachments.append(embedded_msg)
        
        count, names, total_size = self.processor._process_attachments(attachments)
        
        self.assertEqual(count, 2)
        self.assertIn("document.docx", names)
        self.assertIn("FW: Message", names)
        self.assertEqual(total_size, 3072)
    
    def test_safe_get_attribute_success(self):
        """Test safe attribute getting with valid attribute."""
        obj = Mock()
        obj.test_attr = "test_value"
        
        result = self.processor._safe_get_attribute(obj, 'test_attr', 'default')
        self.assertEqual(result, "test_value")
    
    def test_safe_get_attribute_missing(self):
        """Test safe attribute getting with missing attribute."""
        obj = Mock()
        
        result = self.processor._safe_get_attribute(obj, 'missing_attr', 'default')
        self.assertEqual(result, "default")
    
    def test_safe_get_attribute_exception(self):
        """Test safe attribute getting when exception occurs."""
        obj = Mock()
        obj.test_attr = Mock(side_effect=Exception("Test error"))
        
        result = self.processor._safe_get_attribute(obj, 'test_attr', 'default')
        self.assertEqual(result, "default")
    
    def test_process_email_item_with_exception(self):
        """Test email processing when an exception occurs."""
        # Create a mock that raises an exception for Subject
        error_email = Mock()
        error_email.Subject = Mock(side_effect=Exception("COM Error"))
        
        result = self.processor.process_email_item(error_email)
        
        # Should return None when critical error occurs
        self.assertIsNone(result)
    
    def test_process_email_item_com_error(self):
        """Test handling of COM errors during email processing."""
        with patch('outlook2ai.core.email_processor.logger') as mock_logger:
            # Mock email that raises COM error on property access
            com_error_email = Mock()
            com_error_email.Subject = Mock(side_effect=Exception("COM object error"))
            
            result = self.processor.process_email_item(com_error_email)
            
            # Should log error and return None
            self.assertIsNone(result)
            mock_logger.error.assert_called()
    
    def test_process_empty_recipients_collection(self):
        """Test processing when recipients collection is empty."""
        to_list, cc_list, bcc_list = self.processor._process_recipients([])
        
        self.assertEqual(len(to_list), 0)
        self.assertEqual(len(cc_list), 0)
        self.assertEqual(len(bcc_list), 0)
    
    def test_process_empty_attachments_collection(self):
        """Test processing when attachments collection is empty."""
        count, names, total_size = self.processor._process_attachments([])
        
        self.assertEqual(count, 0)
        self.assertEqual(len(names), 0)
        self.assertEqual(total_size, 0)
    
    def test_recipient_type_mapping(self):
        """Test that recipient types are correctly mapped."""
        recipients = []
        
        # Test various recipient types
        for recipient_type in range(1, 4):  # 1=TO, 2=CC, 3=BCC
            recipient = Mock()
            recipient.Name = f"Person {recipient_type}"
            recipient.Address = f"person{recipient_type}@example.com"
            recipient.Type = recipient_type
            recipients.append(recipient)
        
        to_list, cc_list, bcc_list = self.processor._process_recipients(recipients)
        
        # Verify correct type mapping
        self.assertEqual(len(to_list), 1)
        self.assertEqual(len(cc_list), 1)
        self.assertEqual(len(bcc_list), 1)
        
        self.assertIn("person1@example.com", to_list)
        self.assertIn("person2@example.com", cc_list)
        self.assertIn("person3@example.com", bcc_list)
    
    @patch('outlook2ai.core.email_processor.logger')
    def test_logging_on_error(self, mock_logger):
        """Test that errors are properly logged."""
        # Create email that will cause an error
        error_email = Mock()
        error_email.Subject = Mock(side_effect=Exception("Test exception"))
        
        result = self.processor.process_email_item(error_email)
        
        # Verify error was logged
        mock_logger.error.assert_called()
        self.assertIsNone(result)
    
    def test_text_processor_integration(self):
        """Test integration with text processor if available."""
        with patch('outlook2ai.core.email_processor.TextProcessor') as mock_text_processor:
            # Mock the text processor
            mock_processor_instance = Mock()
            mock_processor_instance.clean_html.return_value = "Cleaned HTML"
            mock_processor_instance.extract_text_statistics.return_value = {
                'word_count': 10,
                'char_count': 50
            }
            mock_text_processor.return_value = mock_processor_instance
            
            # Create new processor instance to trigger text processor initialization
            processor = EmailProcessor()
            
            # Process email
            result = processor.process_email_item(self.mock_email)
            
            # Verify text processor was used
            self.assertIsNotNone(result)
    
    def test_malformed_datetime_handling(self):
        """Test handling of malformed datetime objects."""
        malformed_email = Mock()
        malformed_email.Subject = "Test"
        malformed_email.Body = "Test body"
        malformed_email.HTMLBody = ""
        malformed_email.ReceivedTime = "Not a datetime"  # Invalid datetime
        malformed_email.SentOn = None  # Null datetime
        malformed_email.Size = 100
        malformed_email.UnRead = False
        malformed_email.Importance = 1
        malformed_email.Sensitivity = 0
        malformed_email.SenderName = "Test Sender"
        malformed_email.SenderEmailAddress = "test@example.com"
        malformed_email.ConversationID = ""
        malformed_email.ConversationTopic = ""
        malformed_email.EntryID = ""
        malformed_email.Categories = ""
        malformed_email.Recipients = []
        malformed_email.Attachments = []
        
        result = self.processor.process_email_item(malformed_email)
        
        # Should still process successfully with default datetime values
        self.assertIsNotNone(result)
        self.assertEqual(result['subject'], "Test")


if __name__ == '__main__':
    # Configure test runner
    unittest.main(verbosity=2)
