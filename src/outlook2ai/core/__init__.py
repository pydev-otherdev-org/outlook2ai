"""
Core module for Outlook2AI.

This module contains the core components for MS Outlook integration and email processing.
"""

try:
    from .outlook_connector import OutlookConnector
    from .dataframe_manager import DataFrameManager
    from .email_processor import EmailProcessor
    
    __all__ = [
        'OutlookConnector',
        'DataFrameManager',
        'EmailProcessor'
    ]
except ImportError:
    # Handle import errors gracefully
    __all__ = []
