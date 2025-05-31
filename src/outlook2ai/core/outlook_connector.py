"""
MS Outlook Desktop Application Connector

This module provides direct integration with MS Outlook desktop application
using COM interface to extract emails from selected folders.
"""

import win32com.client
import pythoncom
from datetime import datetime, timezone
import logging
from typing import List, Dict, Optional, Any
import time

class OutlookConnector:
    """Connects to MS Outlook desktop application and extracts email data."""
    
    def __init__(self, timeout: int = 30):
        """
        Initialize Outlook connector.
        
        Args:
            timeout: Connection timeout in seconds
        """
        self.timeout = timeout
        self.outlook_app = None
        self.namespace = None
        self.logger = logging.getLogger(__name__)
        
    def connect(self) -> bool:
        """
        Connect to MS Outlook desktop application.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            self.logger.info("Connecting to MS Outlook desktop application...")
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Connect to Outlook application
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook_app.GetNamespace("MAPI")
            
            # Test connection by accessing default inbox
            inbox = self.namespace.GetDefaultFolder(6)  # olFolderInbox = 6
            self.logger.info(f"Connected successfully. Inbox contains {inbox.Items.Count} items")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {str(e)}")
            return False
    
    def disconnect(self):
        """Disconnect from Outlook and cleanup COM resources."""
        try:
            if self.namespace:
                self.namespace = None
            if self.outlook_app:
                self.outlook_app = None
            pythoncom.CoUninitialize()
            self.logger.info("Disconnected from Outlook")
        except Exception as e:
            self.logger.warning(f"Error during disconnect: {str(e)}")
    
    def get_folder_list(self) -> List[Dict[str, Any]]:
        """
        Get list of available folders in Outlook.
        
        Returns:
            List[Dict]: List of folder information
        """
        folders = []
        try:
            if not self.namespace:
                raise Exception("Not connected to Outlook")
            
            # Get all mail folders
            for store in self.namespace.Stores:
                root_folder = store.GetRootFolder()
                self._enumerate_folders(root_folder, folders, "")
                
        except Exception as e:
            self.logger.error(f"Error getting folder list: {str(e)}")
            
        return folders
    
    def _enumerate_folders(self, folder, folder_list: List[Dict], path: str):
        """Recursively enumerate all folders."""
        try:
            current_path = f"{path}/{folder.Name}" if path else folder.Name
            
            # Add folder if it contains mail items
            if folder.DefaultItemType == 0:  # olMailItem = 0
                folder_list.append({
                    'name': folder.Name,
                    'path': current_path,
                    'item_count': folder.Items.Count,
                    'folder_object': folder
                })
            
            # Recursively process subfolders
            for subfolder in folder.Folders:
                self._enumerate_folders(subfolder, folder_list, current_path)
                
        except Exception as e:
            self.logger.warning(f"Error accessing folder {folder.Name}: {str(e)}")
    
    def extract_emails_from_folder(self, folder_path: str, max_emails: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Extract emails from specified folder.
        
        Args:
            folder_path: Path to the folder (e.g., "Inbox/Subfolder")
            max_emails: Maximum number of emails to extract (None for all)
            
        Returns:
            List[Dict]: List of email data
        """
        emails = []
        try:
            if not self.namespace:
                raise Exception("Not connected to Outlook")
            
            # Find the folder
            folder = self._find_folder_by_path(folder_path)
            if not folder:
                raise Exception(f"Folder not found: {folder_path}")
            
            self.logger.info(f"Extracting emails from folder: {folder_path} ({folder.Items.Count} items)")
            
            # Get all items from folder
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # Sort by received time, descending
            
            count = 0
            for item in items:
                try:
                    # Only process mail items
                    if item.Class == 43:  # olMail = 43
                        email_data = self._extract_email_data(item, folder_path)
                        emails.append(email_data)
                        count += 1
                        
                        if max_emails and count >= max_emails:
                            break
                            
                except Exception as e:
                    self.logger.warning(f"Error processing email: {str(e)}")
                    continue
            
            self.logger.info(f"Successfully extracted {len(emails)} emails from {folder_path}")
            
        except Exception as e:
            self.logger.error(f"Error extracting emails from folder {folder_path}: {str(e)}")
            
        return emails
    
    def _find_folder_by_path(self, folder_path: str):
        """Find folder by path string."""
        try:
            path_parts = folder_path.split('/')
            
            # Start from default store
            store = self.namespace.DefaultStore
            current_folder = store.GetRootFolder()
            
            for part in path_parts:
                found = False
                for folder in current_folder.Folders:
                    if folder.Name.lower() == part.lower():
                        current_folder = folder
                        found = True
                        break
                
                if not found:
                    return None
            
            return current_folder
            
        except Exception as e:
            self.logger.error(f"Error finding folder {folder_path}: {str(e)}")
            return None
    
    def _extract_email_data(self, mail_item, folder_name: str) -> Dict[str, Any]:
        """Extract data from a single email item."""
        try:
            # Basic email properties
            email_data = {
                'folder_name': folder_name,
                'subject': getattr(mail_item, 'Subject', ''),
                'sender_email': self._get_sender_email(mail_item),
                'sender_name': getattr(mail_item, 'SenderName', ''),
                'received_time': self._convert_outlook_time(getattr(mail_item, 'ReceivedTime', None)),
                'sent_time': self._convert_outlook_time(getattr(mail_item, 'SentOn', None)),
                'body_text': getattr(mail_item, 'Body', ''),
                'body_html': getattr(mail_item, 'HTMLBody', ''),
                'importance': getattr(mail_item, 'Importance', 1),
                'size': getattr(mail_item, 'Size', 0),
                'unread': getattr(mail_item, 'UnRead', False),
                'has_attachments': getattr(mail_item, 'Attachments', None) and len(mail_item.Attachments) > 0,
                'attachment_count': len(getattr(mail_item, 'Attachments', [])),
                'categories': getattr(mail_item, 'Categories', ''),
                'message_class': getattr(mail_item, 'MessageClass', ''),
                'conversation_topic': getattr(mail_item, 'ConversationTopic', ''),
                'to_recipients': self._get_recipients(mail_item, 'To'),
                'cc_recipients': self._get_recipients(mail_item, 'CC'),
                'bcc_recipients': self._get_recipients(mail_item, 'BCC'),
            }
            
            # Additional LLM-useful fields
            email_data.update({
                'body_word_count': len(email_data['body_text'].split()) if email_data['body_text'] else 0,
                'subject_length': len(email_data['subject']) if email_data['subject'] else 0,
                'is_reply': 'RE:' in email_data['subject'].upper() if email_data['subject'] else False,
                'is_forward': 'FW:' in email_data['subject'].upper() if email_data['subject'] else False,
                'domain': email_data['sender_email'].split('@')[1] if '@' in email_data['sender_email'] else '',
                'hour_received': email_data['received_time'].hour if email_data['received_time'] else None,
                'day_of_week': email_data['received_time'].strftime('%A') if email_data['received_time'] else None,
            })
            
            return email_data
            
        except Exception as e:
            self.logger.error(f"Error extracting email data: {str(e)}")
            return {}
    
    def _get_sender_email(self, mail_item) -> str:
        """Extract sender email address."""
        try:
            if hasattr(mail_item, 'SenderEmailAddress'):
                return mail_item.SenderEmailAddress
            elif hasattr(mail_item, 'Sender') and mail_item.Sender:
                return mail_item.Sender.Address
            return ''
        except:
            return ''
    
    def _get_recipients(self, mail_item, recipient_type: str) -> str:
        """Extract recipients of specified type."""
        try:
            recipients = []
            if hasattr(mail_item, 'Recipients'):
                for recipient in mail_item.Recipients:
                    if recipient.Type == {'To': 1, 'CC': 2, 'BCC': 3}.get(recipient_type, 1):
                        recipients.append(recipient.Address)
            return '; '.join(recipients)
        except:
            return ''
    
    def _convert_outlook_time(self, outlook_time) -> Optional[datetime]:
        """Convert Outlook time to Python datetime."""
        try:
            if outlook_time:
                # Outlook times are in local time, convert to UTC
                return outlook_time.replace(tzinfo=timezone.utc)
            return None
        except:
            return None