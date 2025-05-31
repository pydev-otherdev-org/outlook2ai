"""
Outlook2AI - MS Outlook Email Extraction and Analysis Tool

A comprehensive system for extracting emails from MS Outlook via COM interface,
processing them into pandas DataFrames, and preparing data for LLM analysis.

Author: em7admin
Version: 1.0.0
Created: 2024-01-15
"""

__version__ = "1.0.0"
__author__ = "em7admin"
__description__ = "MS Outlook Email Extraction and Analysis Tool"

# Import main classes for easier access
try:
    from .core.outlook_connector import OutlookConnector
    from .core.dataframe_manager import DataFrameManager
    from .core.email_processor import EmailProcessor
    from .processors.text_processor import TextProcessor
    
    __all__ = [
        'OutlookConnector',
        'DataFrameManager', 
        'EmailProcessor',
        'TextProcessor'
    ]
except ImportError:
    # Handle import errors gracefully for development/testing
    __all__ = []

# Package metadata
__package_info__ = {
    'name': 'outlook2ai',
    'version': __version__,
    'author': __author__,
    'description': __description__,
    'python_requires': '>=3.6',
    'platforms': ['Windows'],
    'dependencies': [
        'pandas>=1.2.0',
        'pywin32>=300',
        'pyyaml>=5.4.0',
        'beautifulsoup4>=4.9.0'
    ],
    'optional_dependencies': [
        'nltk>=3.6',
        'textstat>=0.7.0'
    ]
}
