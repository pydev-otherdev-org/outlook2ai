"""
Text Processing Utilities for Email Content

This module provides text processing utilities for cleaning and preparing
email content for LLM analysis.
"""

import re
import html
try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None
from typing import Dict, List, Optional, Tuple, Any
import logging

class TextProcessor:
    """Handles text processing and cleaning for email content."""
    
    def __init__(self):
        """Initialize text processor."""
        self.logger = logging.getLogger(__name__)
        
    def clean_html_content(self, html_content: str) -> str:
        """
        Clean HTML content and extract plain text.
        
        Args:
            html_content: Raw HTML content from email
            
        Returns:
            str: Cleaned plain text content
        """
        if not html_content:
            return ""
        
        try:
            if BeautifulSoup is None:
                # Fallback: simple HTML tag removal
                text = re.sub(r'<[^>]+>', '', html_content)
                text = html.unescape(text)
                text = re.sub(r'\s+', ' ', text).strip()
                return text
                
            # Parse HTML content
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.decompose()
            
            # Get text content
            text = soup.get_text()
            
            # Clean up whitespace
            lines = (line.strip() for line in text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            text = ' '.join(chunk for chunk in chunks if chunk)
            
            return text
            
        except Exception as e:
            self.logger.error(f"Error cleaning HTML content: {e}")
            return html_content
    
    def clean_plain_text(self, text_content: str) -> str:
        """
        Clean plain text content.
        
        Args:
            text_content: Raw plain text content
            
        Returns:
            str: Cleaned text content
        """
        if not text_content:
            return ""
        
        try:
            # Decode HTML entities
            text = html.unescape(text_content)
            
            # Remove excessive whitespace
            text = re.sub(r'\s+', ' ', text)
            
            # Remove special characters but keep basic punctuation
            text = re.sub(r'[^\w\s\.\,\!\?\;\:\-\(\)\[\]\"\'@]', '', text)
            
            # Strip leading/trailing whitespace
            text = text.strip()
            
            return text
            
        except Exception as e:
            self.logger.error(f"Error cleaning plain text: {e}")
            return text_content
    
    def extract_email_addresses(self, text: str) -> List[str]:
        """
        Extract email addresses from text.
        
        Args:
            text: Text content to search
            
        Returns:
            List[str]: List of found email addresses
        """
        if not text:
            return []
        
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = re.findall(email_pattern, text)
        return list(set(emails))  # Remove duplicates
    
    def extract_phone_numbers(self, text: str) -> List[str]:
        """
        Extract phone numbers from text.
        
        Args:
            text: Text content to search
            
        Returns:
            List[str]: List of found phone numbers
        """
        if not text:
            return []
        
        # Common phone number patterns
        patterns = [
            r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b',  # XXX-XXX-XXXX or XXX.XXX.XXXX
            r'\b\(\d{3}\)\s?\d{3}[-.]?\d{4}\b',  # (XXX) XXX-XXXX
            r'\b\d{3}\s\d{3}\s\d{4}\b',  # XXX XXX XXXX
        ]
        
        phone_numbers = []
        for pattern in patterns:
            phones = re.findall(pattern, text)
            phone_numbers.extend(phones)
        
        return list(set(phone_numbers))  # Remove duplicates
    
    def extract_urls(self, text: str) -> List[str]:
        """
        Extract URLs from text.
        
        Args:
            text: Text content to search
            
        Returns:
            List[str]: List of found URLs
        """
        if not text:
            return []
        
        url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
        urls = re.findall(url_pattern, text)
        return list(set(urls))  # Remove duplicates
    
    def get_text_statistics(self, text: str) -> Dict[str, int]:
        """
        Get basic statistics about text content.
        
        Args:
            text: Text content to analyze
            
        Returns:
            Dict[str, int]: Dictionary containing text statistics
        """
        if not text:
            return {
                'character_count': 0,
                'word_count': 0,
                'sentence_count': 0,
                'paragraph_count': 0
            }
        
        # Character count
        char_count = len(text)
        
        # Word count
        words = text.split()
        word_count = len(words)
        
        # Sentence count (approximate)
        sentences = re.split(r'[.!?]+', text)
        sentence_count = len([s for s in sentences if s.strip()])
        
        # Paragraph count (approximate)
        paragraphs = text.split('\n\n')
        paragraph_count = len([p for p in paragraphs if p.strip()])
        
        return {
            'character_count': char_count,
            'word_count': word_count,
            'sentence_count': sentence_count,
            'paragraph_count': paragraph_count
        }
    
    def extract_keywords(self, text: str, min_length: int = 3) -> List[str]:
        """
        Extract potential keywords from text.
        
        Args:
            text: Text content to analyze
            min_length: Minimum length for keywords
            
        Returns:
            List[str]: List of potential keywords
        """
        if not text:
            return []
        
        # Convert to lowercase and split into words
        words = re.findall(r'\b\w+\b', text.lower())
        
        # Filter out common stop words
        stop_words = {
            'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for',
            'of', 'with', 'by', 'is', 'are', 'was', 'were', 'be', 'been', 'have',
            'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should',
            'may', 'might', 'must', 'can', 'this', 'that', 'these', 'those', 'i',
            'you', 'he', 'she', 'it', 'we', 'they', 'me', 'him', 'her', 'us',
            'them', 'my', 'your', 'his', 'her', 'its', 'our', 'their', 'am'
        }
        
        # Filter words
        keywords = [
            word for word in words 
            if len(word) >= min_length and word not in stop_words
        ]
        
        # Count frequency and return most common
        from collections import Counter
        word_freq = Counter(keywords)
        
        # Return words that appear more than once, sorted by frequency
        return [word for word, count in word_freq.most_common(50) if count > 1]
    
    def process_email_body(self, html_body: str, text_body: str) -> Dict[str, Any]:
        """
        Process email body content and extract useful information.
        
        Args:
            html_body: HTML version of email body
            text_body: Plain text version of email body
            
        Returns:
            Dict[str, Any]: Processed content and metadata
        """
        result = {}
        
        # Clean the content
        if html_body:
            result['cleaned_html'] = self.clean_html_content(html_body)
        else:
            result['cleaned_html'] = ""
        
        if text_body:
            result['cleaned_text'] = self.clean_plain_text(text_body)
        else:
            result['cleaned_text'] = ""
        
        # Use the better version for analysis
        analysis_text = result['cleaned_text'] if result['cleaned_text'] else result['cleaned_html']
        
        # Extract entities
        result['email_addresses'] = self.extract_email_addresses(analysis_text)
        result['phone_numbers'] = self.extract_phone_numbers(analysis_text)
        result['urls'] = self.extract_urls(analysis_text)
        
        # Get statistics
        result['statistics'] = self.get_text_statistics(analysis_text)
        
        # Extract keywords
        result['keywords'] = self.extract_keywords(analysis_text)
        
        # Create final processed text for LLM analysis
        result['llm_optimized_text'] = analysis_text[:10000]  # Limit for LLM context
        
        return result
