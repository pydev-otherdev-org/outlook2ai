"""
Processors module for Outlook2AI.

This module contains text processing and content analysis components.
"""

try:
    from .text_processor import TextProcessor
    
    __all__ = ['TextProcessor']
except ImportError:
    # Handle import errors gracefully
    __all__ = []
