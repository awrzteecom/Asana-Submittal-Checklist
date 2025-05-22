"""
Logger utility module for the DOCX to Asana CSV Generator.

This module provides logging functionality for the application, with configurable
log levels, formats, and output destinations.
"""

import logging
import os
import json
from typing import Optional, Dict, Any


class Logger:
    """
    Logger class for handling application logging.
    
    Provides methods for logging at different levels (debug, info, warning, error, critical)
    and configures logging based on application settings.
    """
    
    def __init__(self, name: str, config_path: Optional[str] = None):
        """
        Initialize the logger with the given name and optional configuration.
        
        Args:
            name: The name of the logger, typically the module name
            config_path: Optional path to a JSON configuration file
        """
        self.logger = logging.getLogger(name)
        self.config: Dict[str, Any] = {}
        
        # Set default configuration
        self._set_default_config()
        
        # Load configuration from file if provided
        if config_path and os.path.exists(config_path):
            self._load_config(config_path)
            
        # Configure the logger
        self._configure_logger()
    
    def _set_default_config(self) -> None:
        """Set default logging configuration."""
        self.config = {
            "level": "INFO",
            "file_path": "app.log",
            "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        }
    
    def _load_config(self, config_path: str) -> None:
        """
        Load logging configuration from a JSON file.
        
        Args:
            config_path: Path to the JSON configuration file
        """
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
                if "logging" in config:
                    self.config.update(config["logging"])
        except (json.JSONDecodeError, IOError) as e:
            print(f"Error loading logger configuration: {e}")
    
    def _configure_logger(self) -> None:
        """Configure the logger based on the current configuration."""
        # Set the logging level
        level = getattr(logging, self.config["level"], logging.INFO)
        self.logger.setLevel(level)
        
        # Create formatter
        formatter = logging.Formatter(self.config["format"])
        
        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        # Create file handler if file_path is specified
        if self.config["file_path"]:
            try:
                file_handler = logging.FileHandler(self.config["file_path"])
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
            except IOError as e:
                print(f"Error creating log file: {e}")
    
    def debug(self, message: str) -> None:
        """
        Log a debug message.
        
        Args:
            message: The message to log
        """
        self.logger.debug(message)
    
    def info(self, message: str) -> None:
        """
        Log an info message.
        
        Args:
            message: The message to log
        """
        self.logger.info(message)
    
    def warning(self, message: str) -> None:
        """
        Log a warning message.
        
        Args:
            message: The message to log
        """
        self.logger.warning(message)
    
    def error(self, message: str) -> None:
        """
        Log an error message.
        
        Args:
            message: The message to log
        """
        self.logger.error(message)
    
    def critical(self, message: str) -> None:
        """
        Log a critical message.
        
        Args:
            message: The message to log
        """
        self.logger.critical(message)


# Create a default logger instance for import
default_logger = Logger(__name__)


def get_logger(name: str, config_path: Optional[str] = None) -> Logger:
    """
    Get a configured logger instance.
    
    Args:
        name: The name of the logger
        config_path: Optional path to a configuration file
        
    Returns:
        A configured Logger instance
    """
    return Logger(name, config_path)
