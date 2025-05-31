"""
Main Application Entry Point for Outlook2AI

This module provides the main interface for extracting emails from MS Outlook
and creating DataFrames for LLM analysis.
"""

import sys
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
import argparse
from datetime import datetime

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from outlook2ai.core.outlook_connector import OutlookConnector
from outlook2ai.core.dataframe_manager import DataFrameManager
from outlook2ai.utils.config_manager import ConfigManager
from outlook2ai.utils.logger import setup_logging

class Outlook2AI:
    """Main application class for Outlook email extraction and analysis."""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize Outlook2AI application.
        
        Args:
            config_path: Path to configuration file
        """
        # Setup logging
        setup_logging()
        self.logger = logging.getLogger(__name__)
        
        # Initialize components
        self.config = ConfigManager(config_path)
        self.outlook_connector = OutlookConnector()
        self.df_manager = DataFrameManager()
        
        self.logger.info("Outlook2AI initialized successfully")
    
    def connect_to_outlook(self) -> bool:
        """
        Connect to MS Outlook desktop application.
        
        Returns:
            bool: True if connection successful
        """
        try:
            self.logger.info("Attempting to connect to MS Outlook...")
            return self.outlook_connector.connect()
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {str(e)}")
            return False
    
    def list_folders(self) -> List[Dict[str, Any]]:
        """
        Get list of available Outlook folders.
        
        Returns:
            List[Dict]: Available folders with metadata
        """
        try:
            folders = self.outlook_connector.get_folder_list()
            self.logger.info(f"Found {len(folders)} available folders")
            return folders
        except Exception as e:
            self.logger.error(f"Error listing folders: {str(e)}")
            return []
    
    def extract_emails(self, folder_paths: List[str], max_emails_per_folder: Optional[int] = None) -> bool:
        """
        Extract emails from specified folders.
        
        Args:
            folder_paths: List of folder paths to extract from
            max_emails_per_folder: Maximum emails per folder (None for all)
            
        Returns:
            bool: True if extraction successful
        """
        try:
            all_emails = []
            
            for folder_path in folder_paths:
                self.logger.info(f"Extracting emails from folder: {folder_path}")
                
                emails = self.outlook_connector.extract_emails_from_folder(
                    folder_path, max_emails_per_folder
                )
                
                if emails:
                    all_emails.extend(emails)
                    self.logger.info(f"Extracted {len(emails)} emails from {folder_path}")
                else:
                    self.logger.warning(f"No emails extracted from {folder_path}")
            
            if all_emails:
                # Create DataFrame
                self.df = self.df_manager.create_dataframe(all_emails)
                self.logger.info(f"Created DataFrame with {len(self.df)} total emails")
                return True
            else:
                self.logger.error("No emails were extracted from any folder")
                return False
                
        except Exception as e:
            self.logger.error(f"Error during email extraction: {str(e)}")
            return False
    
    def get_dataframe(self):
        """Get the current email DataFrame."""
        return getattr(self, 'df', None)
    
    def get_summary_statistics(self) -> Dict[str, Any]:
        """Get summary statistics of extracted emails."""
        try:
            if hasattr(self, 'df') and self.df is not None:
                return self.df_manager.get_summary_stats()
            else:
                return {"error": "No data available. Extract emails first."}
        except Exception as e:
            self.logger.error(f"Error getting summary statistics: {str(e)}")
            return {"error": str(e)}
    
    def export_data(self, output_path: str, format_type: str = 'csv') -> bool:
        """
        Export email data for analysis.
        
        Args:
            output_path: Path to save exported data
            format_type: Export format ('csv', 'json', 'parquet')
            
        Returns:
            bool: True if export successful
        """
        try:
            if not hasattr(self, 'df') or self.df is None:
                self.logger.error("No data to export. Extract emails first.")
                return False
            
            return self.df_manager.export_for_llm_analysis(output_path, format_type)
            
        except Exception as e:
            self.logger.error(f"Error exporting data: {str(e)}")
            return False
    
    def prepare_for_llm(self, max_emails: int = 100) -> str:
        """
        Prepare data for LLM analysis.
        
        Args:
            max_emails: Maximum number of emails to include
            
        Returns:
            str: Formatted data for LLM analysis
        """
        try:
            if not hasattr(self, 'df') or self.df is None:
                return "No data available. Extract emails first."
            
            return self.df_manager.prepare_llm_prompt_data(max_emails)
            
        except Exception as e:
            self.logger.error(f"Error preparing LLM data: {str(e)}")
            return f"Error: {str(e)}"
    
    def disconnect(self):
        """Disconnect from Outlook and cleanup resources."""
        try:
            self.outlook_connector.disconnect()
            self.logger.info("Disconnected from Outlook")
        except Exception as e:
            self.logger.error(f"Error during disconnect: {str(e)}")

def main():
    """Main entry point for command-line usage."""
    parser = argparse.ArgumentParser(description="Extract emails from MS Outlook for LLM analysis")
    parser.add_argument("--folders", nargs="+", default=["Inbox"], 
                       help="Folder paths to extract emails from")
    parser.add_argument("--max-emails", type=int, default=None,
                       help="Maximum emails per folder (default: all)")
    parser.add_argument("--output", default="./data/emails.csv",
                       help="Output file path")
    parser.add_argument("--format", choices=['csv', 'json', 'parquet'], default='csv',
                       help="Output format")
    parser.add_argument("--list-folders", action='store_true',
                       help="List available folders and exit")
    parser.add_argument("--config", help="Path to configuration file")
    
    args = parser.parse_args()
    
    # Initialize application
    app = Outlook2AI(args.config)
    
    try:
        # Connect to Outlook
        if not app.connect_to_outlook():
            print("ERROR: Failed to connect to MS Outlook")
            return 1
        
        # List folders if requested
        if args.list_folders:
            folders = app.list_folders()
            print("\nAvailable Outlook Folders:")
            print("-" * 50)
            for folder in folders:
                print(f"Path: {folder['path']}")
                print(f"Items: {folder['item_count']}")
                print("-" * 30)
            return 0
        
        # Extract emails
        print(f"Extracting emails from folders: {args.folders}")
        if not app.extract_emails(args.folders, args.max_emails):
            print("ERROR: Failed to extract emails")
            return 1
        
        # Show summary
        stats = app.get_summary_statistics()
        print(f"\nExtraction Summary:")
        print(f"Total emails: {stats.get('total_emails', 0)}")
        print(f"Date range: {stats.get('date_range', {}).get('earliest')} to {stats.get('date_range', {}).get('latest')}")
        
        # Export data
        if app.export_data(args.output, args.format):
            print(f"Data exported successfully to: {args.output}")
        else:
            print("ERROR: Failed to export data")
            return 1
        
        return 0
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        return 1
    
    finally:
        app.disconnect()

if __name__ == "__main__":
    sys.exit(main())