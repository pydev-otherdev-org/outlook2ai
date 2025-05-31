"""
Email Processing Module

This module handles the processing of individual email items from Outlook,
extracting metadata and content for DataFrame creation.
"""

import logging
from datetime import datetime, timezone
from typing import Dict, List, Any, Optional
import win32com.client
from outlook2ai.processors.text_processor import TextProcessor

class EmailProcessor:
    """Processes individual email items and extracts relevant data."""
    
    def __init__(self):
        """Initialize email processor."""
        self.logger = logging.getLogger(__name__)
        self.text_processor = TextProcessor()
        
    def process_email_item(self, mail_item: Any, folder_name: str) -> Dict[str, Any]:
        """
        Process a single email item and extract all relevant data.
        
        Args:
            mail_item: Outlook mail item object
            folder_name: Name of the folder containing the email
            
        Returns:
            Dict[str, Any]: Dictionary containing all email data
        """
        try:
            email_data = {}
            
            # Basic email information
            email_data['folder_name'] = folder_name
            email_data['subject'] = self._safe_get_property(mail_item, 'Subject', '')
            email_data['sender_email'] = self._extract_sender_email(mail_item)
            email_data['sender_name'] = self._safe_get_property(mail_item, 'SenderName', '')
            
            # Date/time information
            email_data['received_time'] = self._convert_outlook_time(
                self._safe_get_property(mail_item, 'ReceivedTime')
            )
            email_data['sent_time'] = self._convert_outlook_time(
                self._safe_get_property(mail_item, 'SentOn')
            )
            
            # Body content
            email_data['body_html'] = self._safe_get_property(mail_item, 'HTMLBody', '')
            email_data['body_text'] = self._safe_get_property(mail_item, 'Body', '')
            
            # Process body content
            processed_body = self.text_processor.process_email_body(
                email_data['body_html'],
                email_data['body_text']
            )
            email_data.update(processed_body)
            
            # Email metadata
            email_data['importance'] = self._safe_get_property(mail_item, 'Importance', 1)
            email_data['size'] = self._safe_get_property(mail_item, 'Size', 0)
            email_data['unread'] = self._safe_get_property(mail_item, 'UnRead', False)
            email_data['message_class'] = self._safe_get_property(mail_item, 'MessageClass', '')
            email_data['conversation_topic'] = self._safe_get_property(mail_item, 'ConversationTopic', '')
            
            # Recipients
            email_data['to_recipients'] = self._extract_recipients(mail_item, 'To')
            email_data['cc_recipients'] = self._extract_recipients(mail_item, 'CC')
            email_data['bcc_recipients'] = self._extract_recipients(mail_item, 'BCC')
            
            # Attachments
            attachment_info = self._process_attachments(mail_item)
            email_data['has_attachments'] = attachment_info['has_attachments']
            email_data['attachment_count'] = attachment_info['attachment_count']
            email_data['attachment_names'] = attachment_info['attachment_names']
            email_data['attachment_sizes'] = attachment_info['attachment_sizes']
            
            # Categories
            email_data['categories'] = self._extract_categories(mail_item)
            
            # Additional computed fields for LLM analysis
            email_data['email_thread_id'] = self._safe_get_property(mail_item, 'ConversationID', '')
            email_data['message_id'] = self._safe_get_property(mail_item, 'EntryID', '')
            
            # Flags and properties
            email_data['is_replied'] = self._check_reply_status(mail_item)
            email_data['is_forwarded'] = self._check_forward_status(mail_item)
            email_data['priority'] = self._get_priority_text(email_data['importance'])
            
            # Time-based analysis
            email_data['day_of_week'] = email_data['received_time'].strftime('%A') if email_data['received_time'] else ''
            email_data['hour_of_day'] = email_data['received_time'].hour if email_data['received_time'] else 0
            
            return email_data
            
        except Exception as e:
            self.logger.error(f"Error processing email item: {e}")
            return self._create_error_record(folder_name, str(e))
    
    def _safe_get_property(self, mail_item: Any, property_name: str, default: Any = None) -> Any:
        """
        Safely get a property from a mail item.
        
        Args:
            mail_item: Outlook mail item object
            property_name: Name of the property to get
            default: Default value if property access fails
            
        Returns:
            Any: Property value or default
        """
        try:
            return getattr(mail_item, property_name, default)
        except Exception as e:
            self.logger.debug(f"Failed to get property {property_name}: {e}")
            return default
    
    def _extract_sender_email(self, mail_item: Any) -> str:
        """
        Extract sender email address from mail item.
        
        Args:
            mail_item: Outlook mail item object
            
        Returns:
            str: Sender email address
        """
        try:
            # Try different methods to get sender email
            sender_email = self._safe_get_property(mail_item, 'SenderEmailAddress', '')
            
            if not sender_email or sender_email.startswith('/'):
                # Try to get from Sender object
                sender = self._safe_get_property(mail_item, 'Sender')
                if sender:
                    sender_email = self._safe_get_property(sender, 'Address', '')
            
            if not sender_email or sender_email.startswith('/'):
                # Try to extract from Reply Recipients
                reply_recipients = self._safe_get_property(mail_item, 'ReplyRecipients')
                if reply_recipients and reply_recipients.Count > 0:
                    sender_email = self._safe_get_property(reply_recipients.Item(1), 'Address', '')
            
            return sender_email if sender_email and not sender_email.startswith('/') else ''
            
        except Exception as e:
            self.logger.debug(f"Failed to extract sender email: {e}")
            return ''
    
    def _convert_outlook_time(self, outlook_time: Any) -> Optional[datetime]:
        """
        Convert Outlook time to Python datetime.
        
        Args:
            outlook_time: Outlook datetime object
            
        Returns:
            Optional[datetime]: Converted datetime or None
        """
        try:
            if outlook_time is None:
                return None
            
            # Outlook times are typically in local timezone
            if hasattr(outlook_time, 'strftime'):
                return outlook_time.replace(tzinfo=timezone.utc)
            
            return datetime.fromisoformat(str(outlook_time)).replace(tzinfo=timezone.utc)
            
        except Exception as e:
            self.logger.debug(f"Failed to convert time: {e}")
            return None
    
    def _extract_recipients(self, mail_item: Any, recipient_type: str) -> str:
        """
        Extract recipients of specified type.
        
        Args:
            mail_item: Outlook mail item object
            recipient_type: Type of recipients ('To', 'CC', 'BCC')
            
        Returns:
            str: Semicolon-separated list of recipient emails
        """
        try:
            recipients = []
            recipient_collection = self._safe_get_property(mail_item, f'{recipient_type}Recipients')
            
            if recipient_collection:
                for i in range(1, recipient_collection.Count + 1):
                    recipient = recipient_collection.Item(i)
                    email = self._safe_get_property(recipient, 'Address', '')
                    name = self._safe_get_property(recipient, 'Name', '')
                    
                    if email:
                        if name and name != email:
                            recipients.append(f"{name} <{email}>")
                        else:
                            recipients.append(email)
            
            return '; '.join(recipients)
            
        except Exception as e:
            self.logger.debug(f"Failed to extract {recipient_type} recipients: {e}")
            return ''
    
    def _process_attachments(self, mail_item: Any) -> Dict[str, Any]:
        """
        Process email attachments.
        
        Args:
            mail_item: Outlook mail item object
            
        Returns:
            Dict[str, Any]: Attachment information
        """
        attachment_info = {
            'has_attachments': False,
            'attachment_count': 0,
            'attachment_names': '',
            'attachment_sizes': ''
        }
        
        try:
            attachments = self._safe_get_property(mail_item, 'Attachments')
            if attachments and attachments.Count > 0:
                attachment_info['has_attachments'] = True
                attachment_info['attachment_count'] = attachments.Count
                
                names = []
                sizes = []
                
                for i in range(1, attachments.Count + 1):
                    attachment = attachments.Item(i)
                    name = self._safe_get_property(attachment, 'FileName', f'Attachment_{i}')
                    size = self._safe_get_property(attachment, 'Size', 0)
                    
                    names.append(name)
                    sizes.append(str(size))
                
                attachment_info['attachment_names'] = '; '.join(names)
                attachment_info['attachment_sizes'] = '; '.join(sizes)
            
            return attachment_info
            
        except Exception as e:
            self.logger.debug(f"Failed to process attachments: {e}")
            return attachment_info
    
    def _extract_categories(self, mail_item: Any) -> str:
        """
        Extract email categories.
        
        Args:
            mail_item: Outlook mail item object
            
        Returns:
            str: Semicolon-separated list of categories
        """
        try:
            categories = self._safe_get_property(mail_item, 'Categories', '')
            return categories if categories else ''
        except Exception as e:
            self.logger.debug(f"Failed to extract categories: {e}")
            return ''
    
    def _check_reply_status(self, mail_item: Any) -> bool:
        """
        Check if email has been replied to.
        
        Args:
            mail_item: Outlook mail item object
            
        Returns:
            bool: True if email has been replied to
        """
        try:
            # Check various reply indicators
            reply_recipients = self._safe_get_property(mail_item, 'ReplyRecipients')
            if reply_recipients and reply_recipients.Count > 0:
                return True
            
            # Check subject for "RE:" prefix
            subject = self._safe_get_property(mail_item, 'Subject', '')
            return subject.upper().startswith('RE:')
            
        except Exception as e:
            self.logger.debug(f"Failed to check reply status: {e}")
            return False
    
    def _check_forward_status(self, mail_item: Any) -> bool:
        """
        Check if email has been forwarded.
        
        Args:
            mail_item: Outlook mail item object
            
        Returns:
            bool: True if email has been forwarded
        """
        try:
            # Check subject for "FW:" or "FWD:" prefix
            subject = self._safe_get_property(mail_item, 'Subject', '')
            subject_upper = subject.upper()
            return subject_upper.startswith('FW:') or subject_upper.startswith('FWD:')
            
        except Exception as e:
            self.logger.debug(f"Failed to check forward status: {e}")
            return False
    
    def _get_priority_text(self, importance: int) -> str:
        """
        Convert importance number to text.
        
        Args:
            importance: Outlook importance value
            
        Returns:
            str: Priority as text
        """
        priority_map = {
            0: 'Low',
            1: 'Normal',
            2: 'High'
        }
        return priority_map.get(importance, 'Normal')
    
    def _create_error_record(self, folder_name: str, error_message: str) -> Dict[str, Any]:
        """
        Create an error record for failed email processing.
        
        Args:
            folder_name: Name of the folder
            error_message: Error message
            
        Returns:
            Dict[str, Any]: Error record
        """
        return {
            'folder_name': folder_name,
            'subject': f'ERROR: {error_message}',
            'sender_email': '',
            'sender_name': '',
            'received_time': None,
            'sent_time': None,
            'body_text': '',
            'body_html': '',
            'importance': 1,
            'size': 0,
            'unread': False,
            'has_attachments': False,
            'attachment_count': 0,
            'categories': '',
            'message_class': 'ERROR',
            'conversation_topic': '',
            'to_recipients': '',
            'cc_recipients': '',
            'bcc_recipients': '',
            'error': True,
            'error_message': error_message
        }
