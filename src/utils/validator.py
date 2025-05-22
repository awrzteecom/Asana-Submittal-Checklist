"""
Input validation utility for the DOCX to Asana CSV Generator.

This module provides functions for validating input data, such as file paths,
document structures, and configuration settings.
"""

import os
import re
from typing import List, Dict, Any, Optional, Tuple, Union

from .logger import get_logger

# Initialize logger
logger = get_logger(__name__)


def validate_file_path(file_path: str, extensions: Optional[List[str]] = None) -> bool:
    """
    Validate that a file path exists and has the correct extension.
    
    Args:
        file_path: The file path to validate
        extensions: Optional list of valid file extensions (e.g., ['.docx', '.doc'])
    
    Returns:
        True if the file path is valid, False otherwise
    """
    if not file_path:
        logger.error("File path is empty")
        return False
    
    if not os.path.exists(file_path):
        logger.error(f"File does not exist: {file_path}")
        return False
    
    if not os.path.isfile(file_path):
        logger.error(f"Path is not a file: {file_path}")
        return False
    
    if extensions:
        _, ext = os.path.splitext(file_path)
        if ext.lower() not in extensions:
            logger.error(f"Invalid file extension: {ext}. Expected one of: {extensions}")
            return False
    
    return True


def validate_directory_path(directory_path: str, create_if_missing: bool = False) -> bool:
    """
    Validate that a directory path exists.
    
    Args:
        directory_path: The directory path to validate
        create_if_missing: Whether to create the directory if it doesn't exist
    
    Returns:
        True if the directory path is valid, False otherwise
    """
    if not directory_path:
        logger.error("Directory path is empty")
        return False
    
    if os.path.exists(directory_path):
        if not os.path.isdir(directory_path):
            logger.error(f"Path is not a directory: {directory_path}")
            return False
    elif create_if_missing:
        try:
            os.makedirs(directory_path, exist_ok=True)
            logger.info(f"Created directory: {directory_path}")
        except OSError as e:
            logger.error(f"Error creating directory: {e}")
            return False
    else:
        logger.error(f"Directory does not exist: {directory_path}")
        return False
    
    return True


def validate_docx_structure(document_data: Dict[str, Any]) -> Tuple[bool, List[str]]:
    """
    Validate the structure of a parsed DOCX document.
    
    Args:
        document_data: Dictionary containing the parsed document data
    
    Returns:
        A tuple containing:
        - A boolean indicating whether the document structure is valid
        - A list of validation error messages
    """
    errors = []
    
    # Check for required sections
    if "filename" not in document_data or not document_data["filename"]:
        errors.append("Document filename is missing")
    
    if "products_section" not in document_data or not document_data["products_section"]:
        errors.append("Products section not found in document")
    
    # Check for product types
    if "product_types" not in document_data or not document_data["product_types"]:
        errors.append("No product types found in document")
    
    # Log validation results
    if errors:
        for error in errors:
            logger.error(error)
        return False, errors
    
    return True, []


def validate_csv_data(csv_data: List[Dict[str, str]], required_columns: List[str]) -> Tuple[bool, List[str]]:
    """
    Validate CSV data before writing to file.
    
    Args:
        csv_data: List of dictionaries representing CSV rows
        required_columns: List of required column names
    
    Returns:
        A tuple containing:
        - A boolean indicating whether the CSV data is valid
        - A list of validation error messages
    """
    errors = []
    
    if not csv_data:
        errors.append("CSV data is empty")
        return False, errors
    
    # Check for required columns in each row
    for i, row in enumerate(csv_data):
        for column in required_columns:
            if column not in row:
                errors.append(f"Row {i+1} is missing required column: {column}")
    
    # Log validation results
    if errors:
        for error in errors:
            logger.error(error)
        return False, errors
    
    return True, []


def sanitize_string(text: str) -> str:
    """
    Sanitize a string for CSV output.
    
    Args:
        text: The string to sanitize
    
    Returns:
        The sanitized string
    """
    if not text:
        return ""
    
    # Replace newlines with spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Remove any characters that might cause issues in CSV
    text = text.replace('"', '""')  # Escape double quotes
    
    return text.strip()


def validate_config(config: Dict[str, Any], required_keys: List[str]) -> Tuple[bool, List[str]]:
    """
    Validate configuration settings.
    
    Args:
        config: Dictionary containing configuration settings
        required_keys: List of required configuration keys
    
    Returns:
        A tuple containing:
        - A boolean indicating whether the configuration is valid
        - A list of validation error messages
    """
    errors = []
    
    for key in required_keys:
        if key not in config:
            errors.append(f"Missing required configuration key: {key}")
    
    # Log validation results
    if errors:
        for error in errors:
            logger.error(error)
        return False, errors
    
    return True, []
