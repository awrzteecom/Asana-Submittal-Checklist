"""
Configuration management utility for the DOCX to Asana CSV Generator.

This module provides functionality for loading, validating, and accessing
application configuration settings.
"""

import json
import os
from typing import Dict, Any, Optional

from .logger import get_logger

# Initialize logger
logger = get_logger(__name__)


class ConfigManager:
    """
    Configuration manager for the application.
    
    Handles loading configuration from JSON files and provides
    access to configuration settings.
    """
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize the configuration manager.
        
        Args:
            config_path: Optional path to a JSON configuration file
        """
        self.config: Dict[str, Any] = {}
        self.config_path = config_path or os.path.join(os.getcwd(), "config.json")
        
        # Load configuration
        self.load_config()
    
    def load_config(self) -> bool:
        """
        Load configuration from the specified JSON file.
        
        Returns:
            True if configuration was loaded successfully, False otherwise
        """
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    self.config = json.load(f)
                logger.info(f"Configuration loaded from {self.config_path}")
                return True
            else:
                logger.warning(f"Configuration file not found: {self.config_path}")
                return False
        except (json.JSONDecodeError, IOError) as e:
            logger.error(f"Error loading configuration: {e}")
            return False
    
    def save_config(self, config_path: Optional[str] = None) -> bool:
        """
        Save the current configuration to a JSON file.
        
        Args:
            config_path: Optional path to save the configuration to
                         (defaults to the current config_path)
        
        Returns:
            True if configuration was saved successfully, False otherwise
        """
        save_path = config_path or self.config_path
        try:
            with open(save_path, 'w') as f:
                json.dump(self.config, f, indent=4)
            logger.info(f"Configuration saved to {save_path}")
            return True
        except IOError as e:
            logger.error(f"Error saving configuration: {e}")
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        Get a configuration value by key.
        
        Args:
            key: The configuration key (can be a dot-separated path)
            default: The default value to return if the key is not found
        
        Returns:
            The configuration value or the default value
        """
        if "." in key:
            # Handle nested keys
            parts = key.split(".")
            value = self.config
            for part in parts:
                if isinstance(value, dict) and part in value:
                    value = value[part]
                else:
                    return default
            return value
        else:
            # Handle top-level keys
            return self.config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """
        Set a configuration value.
        
        Args:
            key: The configuration key (can be a dot-separated path)
            value: The value to set
        """
        if "." in key:
            # Handle nested keys
            parts = key.split(".")
            config = self.config
            for part in parts[:-1]:
                if part not in config:
                    config[part] = {}
                config = config[part]
            config[parts[-1]] = value
        else:
            # Handle top-level keys
            self.config[key] = value
    
    def validate(self, required_keys: list) -> bool:
        """
        Validate that the configuration contains all required keys.
        
        Args:
            required_keys: List of required configuration keys
        
        Returns:
            True if all required keys are present, False otherwise
        """
        for key in required_keys:
            if self.get(key) is None:
                logger.error(f"Missing required configuration key: {key}")
                return False
        return True


# Create a default configuration manager instance for import
default_config = ConfigManager()


def get_config(config_path: Optional[str] = None) -> ConfigManager:
    """
    Get a configuration manager instance.
    
    Args:
        config_path: Optional path to a configuration file
    
    Returns:
        A ConfigManager instance
    """
    return ConfigManager(config_path)
