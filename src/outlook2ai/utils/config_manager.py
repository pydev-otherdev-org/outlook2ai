"""
Configuration Manager for Outlook2AI

Manages application configuration and settings.
"""

import yaml
import json
from pathlib import Path
from typing import Dict, Any, Optional
import logging

class ConfigManager:
    """Manages application configuration."""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration manager.
        
        Args:
            config_path: Path to configuration file
        """
        self.logger = logging.getLogger(__name__)
        self.config_path = config_path or self._get_default_config_path()
        self.config = self._load_config()
    
    def _get_default_config_path(self) -> str:
        """Get default configuration file path."""
        return str(Path(__file__).parent.parent.parent / "config" / "config.yaml")
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from file."""
        try:
            config_file = Path(self.config_path)
            
            if not config_file.exists():
                self.logger.warning(f"Config file not found: {self.config_path}")
                return self._get_default_config()
            
            with open(config_file, 'r', encoding='utf-8') as f:
                if config_file.suffix.lower() == '.yaml' or config_file.suffix.lower() == '.yml':
                    config = yaml.safe_load(f)
                else:
                    config = json.load(f)
            
            self.logger.info(f"Configuration loaded from: {self.config_path}")
            return config
            
        except Exception as e:
            self.logger.error(f"Error loading config: {str(e)}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration."""
        return {
            'outlook': {
                'timeout': 30,
                'default_folders': ['Inbox'],
                'max_emails_per_folder': None
            },
            'dataframe': {
                'export_format': 'csv',
                'include_html_body': False,
                'clean_text': True
            },
            'llm': {
                'max_emails_for_prompt': 100,
                'include_body_text': True,
                'max_body_length': 1000
            },
            'logging': {
                'level': 'INFO',
                'file': 'logs/outlook2ai.log'
            }
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value by key (supports dot notation)."""
        try:
            keys = key.split('.')
            value = self.config
            
            for k in keys:
                value = value[k]
            
            return value
            
        except (KeyError, TypeError):
            return default
    
    def save_config(self, config_path: Optional[str] = None) -> bool:
        """Save current configuration to file."""
        try:
            save_path = config_path or self.config_path
            config_file = Path(save_path)
            
            # Ensure directory exists
            config_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(config_file, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, default_flow_style=False, indent=2)
            
            self.logger.info(f"Configuration saved to: {save_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error saving config: {str(e)}")
            return False