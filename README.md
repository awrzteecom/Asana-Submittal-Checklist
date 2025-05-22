# Asana-Submittal-Checklist

A Python application that generates a checklist that can be imported into Asana for parsing a Master Format specification's Part 2 product list into tasks and subtasks.

## Overview

The DOCX to Asana CSV Generator processes DOCX files containing product specifications and converts them into Asana-compatible CSV files. The application extracts document structures (headings and content) and organizes them into task hierarchies for project management in Asana.

## Features

- **Document Processing**: Parses DOCX files and extracts structured data from heading levels
- **Tracked Changes Handling**: Automatically accepts all tracked changes and comments before processing
- **Hierarchical Task Creation**: Converts document structure into a hierarchical task list
- **Asana CSV Format**: Generates CSV files compatible with Asana's import format
- **User-Friendly GUI**: Provides a graphical interface for selecting files and folders
- **Batch Processing**: Supports processing multiple documents at once
- **Error Handling**: Robust error handling with detailed logging

## Installation

### Prerequisites

- Python 3.8 or higher
- Required Python packages (installed automatically by the installer):
  - python-docx
  - pandas
  - openpyxl
  - pytest (for running tests)

### Windows Installation

1. Clone or download this repository
2. Run one of the installer scripts:

   **Option 1: Batch File Installer (Recommended)**
   ```
   .\installer\install.bat
   ```

   **Option 2: PowerShell Installer**
   ```
   .\installer\install.ps1
   ```
   
The installer will:
- Check for Python installation
- Install required dependencies
- Create a desktop shortcut
- Set up the application in the user's home directory (~\DOCX-to-Asana)

### Manual Installation

1. Clone or download this repository
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Graphical User Interface

1. Launch the application using the desktop shortcut or by running:
   ```
   python src/docx_to_asana.py --gui
   ```
2. Select the input folder containing DOCX files
3. Select the output folder for CSV files
4. Click "Process Files" to start the conversion

### Command Line Interface

Process a single file:
```
python src/docx_to_asana.py --input path/to/document.docx --output path/to/output/folder
```

Process all files in a directory:
```
python src/docx_to_asana.py --input path/to/documents/folder --output path/to/output/folder
```

Launch the GUI:
```
python src/docx_to_asana.py --gui
```

Use a custom configuration file:
```
python src/docx_to_asana.py --config path/to/config.json --input path/to/document.docx --output path/to/output/folder
```

## Document Structure Requirements

The application processes DOCX files with the following structure:

1. **Root Task**: Each DOCX filename becomes a root task
2. **Products Section**: The application looks for the 2nd instance of "Heading 1" style with text containing "Products"
3. **Product Types**: All "Heading 2" elements under the Products section become sub-tasks
4. **Manufacturers**: "Heading 3" elements with text containing manufacturer-related terms
5. **Descriptions**: All "Heading 4" content under manufacturer sections is combined

### Flexible Heading Detection

The application implements flexible heading detection to handle inconsistent document formatting:

- **Style Matching**: Recognizes various heading style names (e.g., "Heading 1", "Title 1", "H1")
- **Text Matching**: Uses case-insensitive, partial matching for section names
- **Products Section Variations**: Detects variations like "Products", "Product List", "Products and Services"
- **Manufacturer Variations**: Recognizes terms like "Manufacturer", "Mfg", "Supplier", "Vendor"

This flexibility ensures the application works with documents created by different authors with varying formatting styles.

## Asana CSV Format

The generated CSV files follow Asana's import format with the following columns:
```
Task Name,Section/Column,Assignee,Due Date,Priority,Notes,Parent Task,Project
```

- **Task Name**: Document name, product type, or manufacturer name
- **Section/Column**: "CA Submittal Check-list" (configurable)
- **Parent Task**: Document name for product types, product type for manufacturers
- **Notes**: Manufacturer info + descriptions

## Configuration

The application can be configured by editing the `config.json` file:

```json
{
    "application": {
        "name": "DOCX to Asana CSV Generator",
        "version": "1.0.0"
    },
    "asana": {
        "default_section": "CA Submittal Check-list",
        "default_project": ""
    },
    "document": {
        "products_heading": "Products",
        "manufacturer_headings": ["Manufacturer", "Manufacturers", "Mfg", "Manufacturing", "Supplier", "Vendors"],
        "heading_styles": {
            "section": "Heading 1",
            "product_type": "Heading 2",
            "manufacturer": "Heading 3",
            "description": "Heading 4"
        },
        "heading_style_variations": {
            "section": ["heading 1", "title 1", "h1", "header 1", "section"],
            "product_type": ["heading 2", "title 2", "h2", "header 2", "subsection"],
            "manufacturer": ["heading 3", "title 3", "h3", "header 3"],
            "description": ["heading 4", "title 4", "h4", "header 4"]
        },
        "products_heading_variations": [
            "products", "product list", "products and services", 
            "product information", "product specs", "product specifications"
        ]
    },
    "output": {
        "csv_encoding": "utf-8",
        "csv_columns": [
            "Task Name",
            "Section/Column",
            "Assignee",
            "Due Date",
            "Priority",
            "Notes",
            "Parent Task",
            "Project"
        ]
    },
    "logging": {
        "level": "INFO",
        "file_path": "app.log",
        "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    }
}
```

## Testing

Run all tests:
```
pytest tests/ -v
```

Run specific test file:
```
pytest tests/test_parser.py -v
```

Run with coverage:
```
pytest tests/ --cov=src/ --cov-report=html
```

## Project Structure

```
docx-to-asana/
├── src/
│   ├── docx_to_asana.py          # Main application
│   ├── document_parser.py        # DOCX parsing logic
│   ├── csv_generator.py          # Asana CSV creation
│   ├── gui_handler.py            # User interface components
│   └── utils/
│       ├── logger.py             # Logging utilities
│       ├── validator.py          # Input validation
│       └── config.py             # Configuration management
├── installer/
│   ├── install.bat               # Batch file installer script
│   └── install.ps1               # PowerShell installer script
├── tests/
│   ├── test_parser.py
│   ├── test_csv_generator.py
│   └── sample_documents/
├── requirements.txt
├── config.json
└── README.md
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
