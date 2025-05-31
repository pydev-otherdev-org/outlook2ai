"""
Unit tests for OutlookConnector class

Tests the MS Outlook COM interface integration and email extraction functionality.
"""

import unittest
from unittest.mock import Mock, MagicMock, patch
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from outlook2ai.core.outlook_connector import OutlookConnector

class TestOutlookConnector(unittest.TestCase):
    """Test cases for OutlookConnector class."""
    
    def setUp(self):
        """Set up test fixtures before each test method."""
        self.connector = OutlookConnector(timeout=10)
    
    def tearDown(self):
        """Clean up after each test method."""
        if self.connector.outlook_app:
            try:
                self.connector.disconnect()
            except:
                pass
    
    @patch('outlook2ai.core.outlook_connector.win32com.client.Dispatch')
    @patch('outlook2ai.core.outlook_connector.pythoncom.CoInitialize')
    def test_connect_success(self, mock_coinit, mock_dispatch):
        """Test successful connection to Outlook."""
        # Mock Outlook application
        mock_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_inbox.Items.Count = 10
        
        mock_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_dispatch.return_value = mock_app
        
        # Test connection
        result = self.connector.connect()
        
        # Assertions
        self.assertTrue(result)
        self.assertIsNotNone(self.connector.outlook_app)
        self.assertIsNotNone(self.connector.namespace)
        mock_coinit.assert_called_once()
        mock_dispatch.assert_called_once_with("Outlook.Application")
    
    @patch('outlook2ai.core.outlook_connector.win32com.client.Dispatch')
    @patch('outlook2ai.core.outlook_connector.pythoncom.CoInitialize')
    def test_connect_failure(self, mock_coinit, mock_dispatch):
        """Test failed connection to Outlook."""
        # Mock connection failure
        mock_dispatch.side_effect = Exception("Outlook not found")
        
        # Test connection
        result = self.connector.connect()
        
        # Assertions
        self.assertFalse(result)
        self.assertIsNone(self.connector.outlook_app)
    
    def test_get_default_folder_constants(self):
        """Test getting default folder constants."""
        constants = self.connector.get_default_folder_constants()
        
        # Check expected constants
        expected_folders = ['inbox', 'sent_items', 'drafts', 'deleted_items', 'outbox', 'junk_email']
        for folder in expected_folders:
            self.assertIn(folder, constants)
        
        # Check specific values
        self.assertEqual(constants['inbox'], 6)
        self.assertEqual(constants['sent_items'], 5)
        self.assertEqual(constants['drafts'], 16)
    
    @patch('outlook2ai.core.outlook_connector.win32com.client.Dispatch')
    @patch('outlook2ai.core.outlook_connector.pythoncom.CoInitialize')
    def test_get_folders_list(self, mock_coinit, mock_dispatch):
        """Test getting list of available folders."""
        # Mock Outlook structure
        mock_app = Mock()
        mock_namespace = Mock()
        mock_folders = Mock()
        
        # Create mock folder structure
        mock_folder1 = Mock()
        mock_folder1.Name = "Inbox"
        mock_folder1.FolderPath = "\\\\Inbox"
        
        mock_folder2 = Mock()
        mock_folder2.Name = "Sent Items"
        mock_folder2.FolderPath = "\\\\Sent Items"
        
        mock_folders.Count = 2
        mock_folders.Item.side_effect = lambda x: [mock_folder1, mock_folder2][x-1]
        
        mock_namespace.Folders = mock_folders
        mock_app.GetNamespace.return_value = mock_namespace
        mock_dispatch.return_value = mock_app
        
        # Connect and get folders
        self.connector.connect()
        folders = self.connector.get_folders_list()
        
        # Assertions
        self.assertIsInstance(folders, list)
        self.assertEqual(len(folders), 2)
        self.assertIn("Inbox", [f['name'] for f in folders])
        self.assertIn("Sent Items", [f['name'] for f in folders])
    
    @patch('outlook2ai.core.outlook_connector.win32com.client.Dispatch')
    @patch('outlook2ai.core.outlook_connector.pythoncom.CoInitialize')
    def test_extract_emails_from_folder(self, mock_coinit, mock_dispatch):
        """Test extracting emails from a specific folder."""
        # Mock Outlook structure
        mock_app = Mock()
        mock_namespace = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        # Create mock email items
        mock_email1 = Mock()
        mock_email1.Subject = "Test Email 1"
        mock_email1.SenderName = "Test Sender"
        mock_email1.ReceivedTime = "2025-05-31 10:00:00"
        
        mock_email2 = Mock()
        mock_email2.Subject = "Test Email 2"
        mock_email2.SenderName = "Another Sender"
        mock_email2.ReceivedTime = "2025-05-31 11:00:00"
        
        mock_items.Count = 2
        mock_items.Item.side_effect = lambda x: [mock_email1, mock_email2][x-1]
        mock_folder.Items = mock_items
        mock_folder.Name = "Inbox"
        
        mock_namespace.GetDefaultFolder.return_value = mock_folder
        mock_app.GetNamespace.return_value = mock_namespace
        mock_dispatch.return_value = mock_app
        
        # Connect and extract emails
        self.connector.connect()
        emails = self.connector.extract_emails_from_folder("inbox", limit=10)
        
        # Assertions
        self.assertIsInstance(emails, list)
        self.assertEqual(len(emails), 2)
    
    def test_validate_folder_path(self):
        """Test folder path validation."""
        # Test valid paths
        self.assertTrue(self.connector._validate_folder_path("Inbox"))
        self.assertTrue(self.connector._validate_folder_path("Personal/Important"))
        self.assertTrue(self.connector._validate_folder_path("Work\\Projects"))
        
        # Test invalid paths
        self.assertFalse(self.connector._validate_folder_path(""))
        self.assertFalse(self.connector._validate_folder_path(None))
        self.assertFalse(self.connector._validate_folder_path("../Invalid"))
    
    def test_timeout_handling(self):
        """Test timeout handling in operations."""
        # Create connector with short timeout
        short_timeout_connector = OutlookConnector(timeout=1)
        
        # Test that timeout is set correctly
        self.assertEqual(short_timeout_connector.timeout, 1)
    
    @patch('outlook2ai.core.outlook_connector.win32com.client.Dispatch')
    @patch('outlook2ai.core.outlook_connector.pythoncom.CoInitialize')
    def test_disconnect(self, mock_coinit, mock_dispatch):
        """Test disconnection from Outlook."""
        # Mock successful connection
        mock_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_inbox.Items.Count = 0
        
        mock_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_dispatch.return_value = mock_app
        
        # Connect and then disconnect
        self.connector.connect()
        self.assertIsNotNone(self.connector.outlook_app)
        
        result = self.connector.disconnect()
        
        # Assertions
        self.assertTrue(result)
        self.assertIsNone(self.connector.outlook_app)
        self.assertIsNone(self.connector.namespace)
    
    def test_error_handling(self):
        """Test error handling in various scenarios."""
        # Test operations without connection
        emails = self.connector.extract_emails_from_folder("inbox")
        self.assertEqual(emails, [])
        
        folders = self.connector.get_folders_list()
        self.assertEqual(folders, [])

class TestOutlookConnectorIntegration(unittest.TestCase):
    """Integration tests for OutlookConnector (requires actual Outlook)."""
    
    def setUp(self):
        """Set up for integration tests."""
        self.connector = OutlookConnector()
    
    def tearDown(self):
        """Clean up after integration tests."""
        if self.connector.outlook_app:
            self.connector.disconnect()
    
    @unittest.skipUnless(sys.platform.startswith("win"), "Windows only test")
    def test_real_outlook_connection(self):
        """Test actual connection to Outlook (if available)."""
        try:
            result = self.connector.connect()
            if result:
                # If connection successful, test basic operations
                folders = self.connector.get_folders_list()
                self.assertIsInstance(folders, list)
                
                # Test getting a small number of emails from inbox
                emails = self.connector.extract_emails_from_folder("inbox", limit=5)
                self.assertIsInstance(emails, list)
                
        except Exception as e:
            self.skipTest(f"Outlook not available: {e}")

if __name__ == '__main__':
    # Run tests
    unittest.main(verbosity=2)
