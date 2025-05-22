"""
Tests for the CSV generator module.

This module contains tests for the CSV generation functionality,
focusing on Asana-compatible CSV file creation from parsed document data.
"""

import os
import csv
import pytest
import pandas as pd
from typing import Dict, List, Any

import sys
import os

# Add the parent directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.csv_generator import CSVGenerator, generate_csv
from src.utils.config import get_config

# Test constants
TEST_OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")


def setup_module():
    """Set up the test module."""
    # Create output directory if it doesn't exist
    os.makedirs(TEST_OUTPUT_DIR, exist_ok=True)


def create_test_document_data(filename="Test Document"):
    """
    Create test document data for CSV generation.
    
    Args:
        filename: Name of the document
    
    Returns:
        Dictionary containing test document data
    """
    return {
        "filename": filename,
        "products_section": {
            "index": 1,
            "text": "Products"
        },
        "product_types": [
            {
                "name": "Product Type 1",
                "manufacturers": [
                    {
                        "name": "Manufacturer: Company A",
                        "descriptions": [
                            "Description 1",
                            "Description 2"
                        ]
                    }
                ]
            },
            {
                "name": "Product Type 2",
                "manufacturers": [
                    {
                        "name": "Manufacturers: Company B, Company C",
                        "descriptions": [
                            "Description 3"
                        ]
                    },
                    {
                        "name": "Manufacturer: Company D",
                        "descriptions": []
                    }
                ]
            }
        ]
    }


def test_csv_generator_initialization():
    """Test that the CSVGenerator initializes correctly."""
    generator = CSVGenerator()
    
    # Check that the generator loaded configuration
    assert generator.csv_columns is not None
    assert len(generator.csv_columns) > 0
    assert generator.csv_encoding == "utf-8"
    assert generator.default_section == "CA Submittal Check-list"


def test_create_csv_data():
    """Test creating CSV data from document data."""
    # Create test document data
    document_data = create_test_document_data()
    
    # Create CSV data
    generator = CSVGenerator()
    csv_data = generator._create_csv_data(document_data)
    
    # Check that CSV data was created correctly
    assert len(csv_data) == 5  # 1 root task + 2 product types + 2 manufacturers
    
    # Check root task
    assert csv_data[0]["Task Name"] == "Test Document"
    assert csv_data[0]["Section/Column"] == "CA Submittal Check-list"
    assert csv_data[0]["Parent Task"] == ""
    
    # Check product types
    assert csv_data[1]["Task Name"] == "Product Type 1"
    assert csv_data[1]["Parent Task"] == "Test Document"
    
    assert csv_data[2]["Task Name"] == "Manufacturer: Company A"
    assert csv_data[2]["Parent Task"] == "Product Type 1"
    assert "Description 1" in csv_data[2]["Notes"]
    assert "Description 2" in csv_data[2]["Notes"]
    
    assert csv_data[3]["Task Name"] == "Product Type 2"
    assert csv_data[3]["Parent Task"] == "Test Document"
    
    assert csv_data[4]["Task Name"] == "Manufacturers: Company B, Company C"
    assert csv_data[4]["Parent Task"] == "Product Type 2"
    assert "Description 3" in csv_data[4]["Notes"]


def test_generate_csv_file():
    """Test generating a CSV file from document data."""
    # Create test document data
    document_data = create_test_document_data()
    
    # Generate CSV file
    output_path = os.path.join(TEST_OUTPUT_DIR, "test_output.csv")
    result = generate_csv(document_data, output_path)
    
    # Check that the CSV file was generated successfully
    assert result is True
    assert os.path.exists(output_path)
    
    # Read the CSV file and check its contents
    df = pd.read_csv(output_path)
    
    # Check that the CSV file has the correct number of rows
    assert len(df) == 5  # 1 root task + 2 product types + 2 manufacturers
    
    # Check that the CSV file has the correct columns
    expected_columns = [
        "Task Name",
        "Section/Column",
        "Assignee",
        "Due Date",
        "Priority",
        "Notes",
        "Parent Task",
        "Project"
    ]
    for column in expected_columns:
        assert column in df.columns


def test_generate_csv_with_special_characters():
    """Test generating a CSV file with special characters."""
    # Create test document data with special characters
    document_data = create_test_document_data("Test & Document")
    document_data["product_types"][0]["name"] = "Product Type & 1"
    document_data["product_types"][0]["manufacturers"][0]["name"] = "Manufacturer: Company A & B"
    document_data["product_types"][0]["manufacturers"][0]["descriptions"] = ["Description with & and \"quotes\""]
    
    # Generate CSV file
    output_path = os.path.join(TEST_OUTPUT_DIR, "test_special_chars.csv")
    result = generate_csv(document_data, output_path)
    
    # Check that the CSV file was generated successfully
    assert result is True
    assert os.path.exists(output_path)
    
    # Read the CSV file and check its contents
    df = pd.read_csv(output_path)
    
    # Check that special characters were handled correctly
    assert df["Task Name"][0] == "Test & Document"
    assert df["Task Name"][1] == "Product Type & 1"
    assert df["Task Name"][2] == "Manufacturer: Company A & B"
    assert "Description with & and" in df["Notes"][2]
    assert "quotes" in df["Notes"][2]


def test_generate_csv_with_empty_document_data():
    """Test generating a CSV file with empty document data."""
    # Create empty document data
    document_data = {
        "filename": "Empty Document",
        "products_section": None,
        "product_types": []
    }
    
    # Generate CSV file
    output_path = os.path.join(TEST_OUTPUT_DIR, "test_empty.csv")
    result = generate_csv(document_data, output_path)
    
    # Check that the CSV file was generated successfully
    assert result is True
    assert os.path.exists(output_path)
    
    # Read the CSV file and check its contents
    df = pd.read_csv(output_path)
    
    # Check that the CSV file has only the root task
    assert len(df) == 1
    assert df["Task Name"][0] == "Empty Document"


def test_preview_csv():
    """Test generating a preview of the CSV data."""
    # Create test document data
    document_data = create_test_document_data()
    
    # Generate CSV preview
    generator = CSVGenerator()
    preview = generator.preview_csv(document_data)
    
    # Check that the preview was generated
    assert preview is not None
    assert isinstance(preview, str)
    assert "Test Document" in preview
    assert "Product Type 1" in preview
    assert "Product Type 2" in preview


def test_generate_csv_with_missing_filename():
    """Test generating a CSV file with missing filename."""
    # Create document data with missing filename
    document_data = create_test_document_data("")
    
    # Generate CSV file
    output_path = os.path.join(TEST_OUTPUT_DIR, "test_missing_filename.csv")
    result = generate_csv(document_data, output_path)
    
    # Check that the CSV file generation failed
    assert result is False


if __name__ == "__main__":
    pytest.main(["-v", __file__])
