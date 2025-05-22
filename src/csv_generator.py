"""
CSV generator for the DOCX to Asana CSV Generator.

This module handles the creation of Asana-compatible CSV files from
parsed document data, formatting the data according to Asana's import specifications.
"""

import os
import csv
import pandas as pd
from typing import Dict, List, Any, Optional

from .utils.logger import get_logger
from .utils.config import get_config
from .utils.validator import validate_csv_data, sanitize_string

# Initialize logger and config
logger = get_logger(__name__)
config = get_config()


class CSVGenerator:
    """
    Generator for Asana-compatible CSV files.
    
    Creates CSV files from parsed document data, following Asana's
    import format requirements.
    """
    
    def __init__(self):
        """Initialize the CSV generator with configuration settings."""
        # Load CSV configuration
        self.csv_columns = config.get("output.csv_columns", [
            "Task Name",
            "Section/Column",
            "Assignee",
            "Due Date",
            "Priority",
            "Notes",
            "Parent Task",
            "Project"
        ])
        
        self.csv_encoding = config.get("output.csv_encoding", "utf-8")
        self.default_section = config.get("asana.default_section", "CA Submittal Check-list")
        self.default_project = config.get("asana.default_project", "")
    
    def generate_csv(self, document_data: Dict[str, Any], output_path: str) -> bool:
        """
        Generate an Asana-compatible CSV file from parsed document data.
        
        Args:
            document_data: Dictionary containing the parsed document data
            output_path: Path to save the CSV file
        
        Returns:
            True if the CSV file was generated successfully, False otherwise
        """
        try:
            # Extract document filename (root task)
            filename = document_data.get("filename", "")
            if not filename:
                logger.error("Document filename is missing")
                return False
            
            # Create CSV data
            csv_data = self._create_csv_data(document_data)
            
            # Validate CSV data
            valid, errors = validate_csv_data(csv_data, self.csv_columns)
            if not valid:
                logger.error(f"CSV data validation failed: {errors}")
                return False
            
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            
            # Write CSV file using pandas for proper handling of special characters
            df = pd.DataFrame(csv_data)
            df.to_csv(output_path, index=False, encoding=self.csv_encoding)
            
            logger.info(f"CSV file generated: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating CSV file: {e}")
            return False
    
    def _create_csv_data(self, document_data: Dict[str, Any]) -> List[Dict[str, str]]:
        """
        Create CSV data from parsed document data.
        
        Args:
            document_data: Dictionary containing the parsed document data
        
        Returns:
            A list of dictionaries representing CSV rows
        """
        csv_data = []
        
        # Extract document filename (root task)
        filename = document_data.get("filename", "")
        
        # Create root task
        root_task = {
            "Task Name": sanitize_string(filename),
            "Section/Column": self.default_section,
            "Assignee": "",
            "Due Date": "",
            "Priority": "",
            "Notes": "Root task for document",
            "Parent Task": "",
            "Project": self.default_project
        }
        csv_data.append(root_task)
        
        # Process product types
        product_types = document_data.get("product_types", [])
        for product_type in product_types:
            product_name = product_type.get("name", "")
            if not product_name:
                continue
            
            # Create product type task
            product_task = {
                "Task Name": sanitize_string(product_name),
                "Section/Column": self.default_section,
                "Assignee": "",
                "Due Date": "",
                "Priority": "",
                "Notes": "",
                "Parent Task": sanitize_string(filename),
                "Project": self.default_project
            }
            csv_data.append(product_task)
            
            # Process manufacturers
            manufacturers = product_type.get("manufacturers", [])
            for manufacturer in manufacturers:
                manufacturer_name = manufacturer.get("name", "")
                if not manufacturer_name:
                    continue
                
                # Combine descriptions
                descriptions = manufacturer.get("descriptions", [])
                description_text = "\n".join(descriptions) if descriptions else ""
                
                # Create notes with manufacturer info and descriptions
                notes = f"{manufacturer_name}\n\n{description_text}" if description_text else manufacturer_name
                
                # Create manufacturer task
                manufacturer_task = {
                    "Task Name": sanitize_string(manufacturer_name),
                    "Section/Column": self.default_section,
                    "Assignee": "",
                    "Due Date": "",
                    "Priority": "",
                    "Notes": sanitize_string(notes),
                    "Parent Task": sanitize_string(product_name),
                    "Project": self.default_project
                }
                csv_data.append(manufacturer_task)
        
        return csv_data
    
    def preview_csv(self, document_data: Dict[str, Any]) -> str:
        """
        Generate a preview of the CSV data.
        
        Args:
            document_data: Dictionary containing the parsed document data
        
        Returns:
            A string containing a preview of the CSV data
        """
        try:
            # Create CSV data
            csv_data = self._create_csv_data(document_data)
            
            # Create DataFrame
            df = pd.DataFrame(csv_data)
            
            # Return string representation
            return df.to_string(index=False)
            
        except Exception as e:
            logger.error(f"Error generating CSV preview: {e}")
            return f"Error: {e}"


def generate_csv(document_data: Dict[str, Any], output_path: str) -> bool:
    """
    Generate an Asana-compatible CSV file from parsed document data.
    
    Args:
        document_data: Dictionary containing the parsed document data
        output_path: Path to save the CSV file
    
    Returns:
        True if the CSV file was generated successfully, False otherwise
    """
    generator = CSVGenerator()
    return generator.generate_csv(document_data, output_path)


def generate_csvs(documents_data: List[Dict[str, Any]], output_dir: str) -> Dict[str, bool]:
    """
    Generate multiple Asana-compatible CSV files from parsed document data.
    
    Args:
        documents_data: List of dictionaries containing parsed document data
        output_dir: Directory to save the CSV files
    
    Returns:
        A dictionary mapping filenames to success/failure status
    """
    generator = CSVGenerator()
    results = {}
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    for document_data in documents_data:
        filename = document_data.get("filename", "")
        if not filename:
            continue
        
        output_path = os.path.join(output_dir, f"{filename}.csv")
        success = generator.generate_csv(document_data, output_path)
        results[filename] = success
    
    return results
