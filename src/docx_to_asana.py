"""
Main application module for the DOCX to Asana CSV Generator.

This module provides the main entry point for the application,
coordinating between the document parser, CSV generator, and GUI components.
"""

import os
import sys
import argparse
from typing import List, Dict, Any, Optional, Tuple

# Handle imports differently when running as script vs module
if __name__ == "__main__":
    # Add the parent directory to the Python path when running from within src/
    parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if parent_dir not in sys.path:
        sys.path.insert(0, parent_dir)
    
    # Use absolute imports when running as script
    try:
        from src.document_parser import parse_document
        from src.csv_generator import generate_csv
        from src.gui_handler import run_gui
        from src.utils.logger import get_logger
        from src.utils.config import get_config
        from src.utils.validator import validate_file_path, validate_directory_path
    except ImportError:
        # Fallback to direct imports if src module not found
        from document_parser import parse_document
        from csv_generator import generate_csv
        from gui_handler import run_gui
        from utils.logger import get_logger
        from utils.config import get_config
        from utils.validator import validate_file_path, validate_directory_path
else:
    # Use relative imports when imported as module
    from .document_parser import parse_document
    from .csv_generator import generate_csv
    from .gui_handler import run_gui
    from .utils.logger import get_logger
    from .utils.config import get_config
    from .utils.validator import validate_file_path, validate_directory_path

# Initialize logger and config
logger = get_logger(__name__)
config = get_config()


def process_file(file_path: str, output_dir: str) -> bool:
    """
    Process a single DOCX file and generate an Asana CSV file.
    
    Args:
        file_path: Path to the DOCX file
        output_dir: Directory to save the CSV file
    
    Returns:
        True if processing was successful, False otherwise
    """
    try:
        # Validate file path
        if not validate_file_path(file_path, ['.docx']):
            logger.error(f"Invalid document file: {file_path}")
            return False
        
        # Parse the document
        logger.info(f"Parsing document: {file_path}")
        document_data = parse_document(file_path)
        
        if not document_data:
            logger.error(f"Failed to parse document: {file_path}")
            return False
        
        # Generate CSV file
        filename = document_data.get("filename", "")
        if not filename:
            filename = os.path.splitext(os.path.basename(file_path))[0]
        
        output_path = os.path.join(output_dir, f"{filename}.csv")
        logger.info(f"Generating CSV file: {output_path}")
        
        result = generate_csv(document_data, output_path)
        
        if result:
            logger.info(f"Successfully processed file: {file_path}")
            return True
        else:
            logger.error(f"Failed to generate CSV file: {output_path}")
            return False
            
    except Exception as e:
        logger.error(f"Error processing file {file_path}: {e}")
        return False


def process_directory(input_dir: str, output_dir: str) -> Tuple[int, int]:
    """
    Process all DOCX files in a directory.
    
    Args:
        input_dir: Directory containing DOCX files
        output_dir: Directory to save CSV files
    
    Returns:
        A tuple containing (number of successful files, total number of files)
    """
    # Validate directories
    if not validate_directory_path(input_dir):
        logger.error(f"Invalid input directory: {input_dir}")
        return 0, 0
    
    if not validate_directory_path(output_dir, create_if_missing=True):
        logger.error(f"Invalid output directory: {output_dir}")
        return 0, 0
    
    # Find all DOCX files
    docx_files = []
    for file in os.listdir(input_dir):
        if file.lower().endswith(".docx") and not file.startswith("~$"):
            docx_files.append(os.path.join(input_dir, file))
    
    if not docx_files:
        logger.warning(f"No DOCX files found in directory: {input_dir}")
        return 0, 0
    
    # Process each file
    success_count = 0
    total_count = len(docx_files)
    
    for i, file_path in enumerate(docx_files):
        logger.info(f"Processing file {i+1}/{total_count}: {file_path}")
        if process_file(file_path, output_dir):
            success_count += 1
    
    logger.info(f"Processed {success_count}/{total_count} files successfully")
    return success_count, total_count


def parse_arguments() -> argparse.Namespace:
    """
    Parse command-line arguments.
    
    Returns:
        Parsed arguments namespace
    """
    parser = argparse.ArgumentParser(description="DOCX to Asana CSV Generator")
    
    parser.add_argument(
        "--input", "-i",
        help="Input DOCX file or directory"
    )
    
    parser.add_argument(
        "--output", "-o",
        help="Output directory for CSV files"
    )
    
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the graphical user interface"
    )
    
    parser.add_argument(
        "--config", "-c",
        help="Path to configuration file"
    )
    
    return parser.parse_args()


def main() -> int:
    """
    Main entry point for the application.
    
    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    # Parse command-line arguments
    args = parse_arguments()
    
    # Load custom configuration if provided
    if args.config:
        if os.path.exists(args.config):
            config.config_path = args.config
            config.load_config()
            logger.info(f"Loaded configuration from {args.config}")
        else:
            logger.error(f"Configuration file not found: {args.config}")
            return 1
    
    # Launch GUI if requested or if no arguments provided
    if args.gui or (not args.input and not args.output):
        logger.info("Launching GUI")
        run_gui(process_file)
        return 0
    
    # Process files in batch mode
    if not args.input:
        logger.error("No input file or directory specified")
        return 1
    
    if not args.output:
        logger.error("No output directory specified")
        return 1
    
    # Check if input is a file or directory
    if os.path.isfile(args.input):
        # Process single file
        logger.info(f"Processing single file: {args.input}")
        success = process_file(args.input, args.output)
        return 0 if success else 1
    elif os.path.isdir(args.input):
        # Process directory
        logger.info(f"Processing directory: {args.input}")
        success_count, total_count = process_directory(args.input, args.output)
        return 0 if success_count == total_count else 1
    else:
        logger.error(f"Input path does not exist: {args.input}")
        return 1


if __name__ == "__main__":
    try:
        # Configure logging
        logger.info(f"Starting {config.get('application.name', 'DOCX to Asana CSV Generator')} v{config.get('application.version', '1.0.0')}")
        
        # Run the application
        exit_code = main()
        
        # Exit with appropriate code
        sys.exit(exit_code)
    except Exception as e:
        logger.critical(f"Unhandled exception: {e}")
        sys.exit(1)
