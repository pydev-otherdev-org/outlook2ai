"""
DataFrame Manager for Email Data

Manages the creation, manipulation, and export of email data in DataFrame format
optimized for LLM analysis.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import List, Dict, Any, Optional
import logging
import json

class DataFrameManager:
    """Manages email data in DataFrame format for analysis and LLM processing."""
    
    def __init__(self):
        """Initialize DataFrame manager."""
        self.logger = logging.getLogger(__name__)
        self.df = None
        self.column_definitions = self._get_column_definitions()
    
    def _get_column_definitions(self) -> Dict[str, str]:
        """Define standard columns for email DataFrame."""
        return {
            # Core email fields
            'folder_name': 'str',
            'subject': 'str',
            'sender_email': 'str',
            'sender_name': 'str',
            'received_time': 'datetime64[ns]',
            'sent_time': 'datetime64[ns]',
            'body_text': 'str',
            'body_html': 'str',
            
            # Metadata fields
            'importance': 'int64',
            'size': 'int64',
            'unread': 'bool',
            'has_attachments': 'bool',
            'attachment_count': 'int64',
            'categories': 'str',
            'message_class': 'str',
            'conversation_topic': 'str',
            'to_recipients': 'str',
            'cc_recipients': 'str',
            'bcc_recipients': 'str',
            
            # LLM analysis fields
            'body_word_count': 'int64',
            'subject_length': 'int64',
            'is_reply': 'bool',
            'is_forward': 'bool',
            'domain': 'str',
            'hour_received': 'int64',
            'day_of_week': 'str',
            
            # Analysis fields (to be populated by LLM)
            'sentiment': 'str',
            'priority_score': 'float64',
            'topic_category': 'str',
            'requires_action': 'bool',
            'key_entities': 'str',
            'summary': 'str',
        }
    
    def create_dataframe(self, email_data: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Create DataFrame from email data.
        
        Args:
            email_data: List of email dictionaries
            
        Returns:
            pd.DataFrame: Processed email DataFrame
        """
        try:
            if not email_data:
                self.logger.warning("No email data provided")
                return pd.DataFrame()
            
            self.logger.info(f"Creating DataFrame from {len(email_data)} emails")
            
            # Create initial DataFrame
            self.df = pd.DataFrame(email_data)
            
            # Ensure all required columns exist
            for column, dtype in self.column_definitions.items():
                if column not in self.df.columns:
                    if dtype == 'str':
                        self.df[column] = ''
                    elif dtype == 'bool':
                        self.df[column] = False
                    elif 'int' in dtype:
                        self.df[column] = 0
                    elif 'float' in dtype:
                        self.df[column] = 0.0
                    elif 'datetime' in dtype:
                        self.df[column] = pd.NaT
            
            # Apply data type conversions
            self._apply_data_types()
            
            # Clean and process data
            self._clean_data()
            
            # Add computed fields
            self._add_computed_fields()
            
            self.logger.info(f"DataFrame created successfully with shape: {self.df.shape}")
            return self.df
            
        except Exception as e:
            self.logger.error(f"Error creating DataFrame: {str(e)}")
            return pd.DataFrame()
    
    def _apply_data_types(self):
        """Apply proper data types to DataFrame columns."""
        try:
            for column, dtype in self.column_definitions.items():
                if column in self.df.columns:
                    if dtype == 'datetime64[ns]':
                        self.df[column] = pd.to_datetime(self.df[column], errors='coerce')
                    elif dtype == 'bool':
                        self.df[column] = self.df[column].astype(bool)
                    elif 'int' in dtype:
                        self.df[column] = pd.to_numeric(self.df[column], errors='coerce').fillna(0).astype('int64')
                    elif 'float' in dtype:
                        self.df[column] = pd.to_numeric(self.df[column], errors='coerce').fillna(0.0)
                    else:  # string types
                        self.df[column] = self.df[column].astype(str).fillna('')
                        
        except Exception as e:
            self.logger.error(f"Error applying data types: {str(e)}")
    
    def _clean_data(self):
        """Clean and normalize email data."""
        try:
            # Clean text fields
            text_columns = ['subject', 'body_text', 'sender_name', 'sender_email']
            for col in text_columns:
                if col in self.df.columns:
                    self.df[col] = self.df[col].str.strip()
                    self.df[col] = self.df[col].replace('', np.nan)
            
            # Normalize email addresses
            if 'sender_email' in self.df.columns:
                self.df['sender_email'] = self.df['sender_email'].str.lower()
            
            # Clean HTML from body text if needed
            if 'body_text' in self.df.columns:
                self.df['body_text_clean'] = self.df['body_text'].apply(self._clean_text)
            
        except Exception as e:
            self.logger.error(f"Error cleaning data: {str(e)}")
    
    def _clean_text(self, text: str) -> str:
        """Clean text content for better analysis."""
        if pd.isna(text) or text == '':
            return ''
        
        try:
            # Remove excessive whitespace
            text = ' '.join(text.split())
            
            # Remove common email artifacts
            text = text.replace('\r\n', '\n').replace('\r', '\n')
            
            return text
            
        except Exception:
            return str(text)
    
    def _add_computed_fields(self):
        """Add computed fields for LLM analysis."""
        try:
            # Add email age in days
            if 'received_time' in self.df.columns:
                now = pd.Timestamp.now()
                self.df['age_days'] = (now - self.df['received_time']).dt.days
            
            # Add time-based categories
            if 'hour_received' in self.df.columns:
                self.df['time_category'] = self.df['hour_received'].apply(self._categorize_time)
            
            # Add size categories
            if 'size' in self.df.columns:
                self.df['size_category'] = pd.cut(
                    self.df['size'], 
                    bins=[0, 1000, 10000, 100000, float('inf')],
                    labels=['Small', 'Medium', 'Large', 'Very Large']
                )
            
            # Add recipient count
            for col in ['to_recipients', 'cc_recipients', 'bcc_recipients']:
                if col in self.df.columns:
                    count_col = col.replace('_recipients', '_count')
                    self.df[count_col] = self.df[col].apply(lambda x: len(x.split(';')) if x else 0)
            
        except Exception as e:
            self.logger.error(f"Error adding computed fields: {str(e)}")
    
    def _categorize_time(self, hour: int) -> str:
        """Categorize email by time of day."""
        if pd.isna(hour):
            return 'Unknown'
        elif 6 <= hour < 12:
            return 'Morning'
        elif 12 <= hour < 17:
            return 'Afternoon'
        elif 17 <= hour < 21:
            return 'Evening'
        else:
            return 'Night'
    
    def get_summary_stats(self) -> Dict[str, Any]:
        """Get summary statistics of the email DataFrame."""
        if self.df is None or self.df.empty:
            return {}
        
        try:
            stats = {
                'total_emails': len(self.df),
                'date_range': {
                    'earliest': self.df['received_time'].min(),
                    'latest': self.df['received_time'].max()
                },
                'folder_distribution': self.df['folder_name'].value_counts().to_dict(),
                'sender_distribution': self.df['sender_email'].value_counts().head(10).to_dict(),
                'domain_distribution': self.df['domain'].value_counts().head(10).to_dict(),
                'avg_body_length': self.df['body_word_count'].mean(),
                'unread_count': self.df['unread'].sum(),
                'with_attachments': self.df['has_attachments'].sum(),
                'size_stats': {
                    'mean': self.df['size'].mean(),
                    'median': self.df['size'].median(),
                    'max': self.df['size'].max()
                }
            }
            
            return stats
            
        except Exception as e:
            self.logger.error(f"Error generating summary stats: {str(e)}")
            return {}
    
    def export_for_llm_analysis(self, output_path: str, format_type: str = 'csv') -> bool:
        """
        Export DataFrame in format optimized for LLM analysis.
        
        Args:
            output_path: Path to save the exported data
            format_type: Export format ('csv', 'json', 'parquet')
            
        Returns:
            bool: True if export successful
        """
        try:
            if self.df is None or self.df.empty:
                self.logger.error("No data to export")
                return False
            
            # Select columns most relevant for LLM analysis
            llm_columns = [
                'folder_name', 'subject', 'sender_email', 'sender_name',
                'received_time', 'body_text_clean', 'importance',
                'has_attachments', 'body_word_count', 'is_reply', 'is_forward',
                'domain', 'time_category', 'age_days'
            ]
            
            # Filter to existing columns
            available_columns = [col for col in llm_columns if col in self.df.columns]
            export_df = self.df[available_columns].copy()
            
            # Export in requested format
            if format_type.lower() == 'csv':
                export_df.to_csv(output_path, index=False, encoding='utf-8')
            elif format_type.lower() == 'json':
                export_df.to_json(output_path, orient='records', date_format='iso', indent=2)
            elif format_type.lower() == 'parquet':
                export_df.to_parquet(output_path, index=False)
            else:
                raise ValueError(f"Unsupported format: {format_type}")
            
            self.logger.info(f"Data exported successfully to {output_path} in {format_type} format")
            return True
            
        except Exception as e:
            self.logger.error(f"Error exporting data: {str(e)}")
            return False
    
    def prepare_llm_prompt_data(self, max_emails: int = 100) -> str:
        """
        Prepare email data for LLM prompt analysis.
        
        Args:
            max_emails: Maximum number of emails to include
            
        Returns:
            str: Formatted string for LLM analysis
        """
        try:
            if self.df is None or self.df.empty:
                return "No email data available for analysis."
            
            # Sample data if too large
            sample_df = self.df.head(max_emails) if len(self.df) > max_emails else self.df.copy()
            
            # Create summary for LLM
            summary = []
            summary.append(f"EMAIL DATASET SUMMARY:")
            summary.append(f"Total emails: {len(sample_df)}")
            summary.append(f"Date range: {sample_df['received_time'].min()} to {sample_df['received_time'].max()}")
            summary.append(f"Unique senders: {sample_df['sender_email'].nunique()}")
            summary.append(f"Folders: {', '.join(sample_df['folder_name'].unique())}")
            summary.append("\nSAMPLE EMAIL DATA:")
            
            # Add sample emails with key fields
            for idx, row in sample_df.iterrows():
                email_summary = []
                email_summary.append(f"\nEmail {idx + 1}:")
                email_summary.append(f"  Folder: {row['folder_name']}")
                email_summary.append(f"  From: {row['sender_email']}")
                email_summary.append(f"  Subject: {row['subject'][:100]}...")
                email_summary.append(f"  Received: {row['received_time']}")
                email_summary.append(f"  Body (first 200 chars): {str(row.get('body_text_clean', ''))[:200]}...")
                
                summary.extend(email_summary)
            
            return '\n'.join(summary)
            
        except Exception as e:
            self.logger.error(f"Error preparing LLM prompt data: {str(e)}")
            return f"Error preparing data: {str(e)}"