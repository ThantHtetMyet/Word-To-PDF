"""
Word to PDF Converter
Converts Microsoft Word documents (.docx, .doc) to PDF format without losing any content or formatting.
"""

import os
import sys
import json
import argparse
from pathlib import Path
from docx2pdf import convert


def convert_word_to_pdf(input_path, output_path=None):
    """
    Convert a Word document to PDF.
    
    Args:
        input_path (str): Path to the input Word document
        output_path (str, optional): Path for the output PDF. If None, uses same name with .pdf extension
    
    Returns:
        str: Path to the generated PDF file
    """
    input_file = Path(input_path)
    
    # Check if input file exists
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    
    # Check if input file is a Word document
    if input_file.suffix.lower() not in ['.docx', '.doc']:
        raise ValueError(f"Input file must be a Word document (.docx or .doc): {input_path}")
    
    # Determine output path
    if output_path is None:
        output_file = input_file.with_suffix('.pdf')
    else:
        output_file = Path(output_path)
        # Create output directory if it doesn't exist
        output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Get file size for user information
    file_size_mb = input_file.stat().st_size / (1024 * 1024)
    
    print(f"Converting: {input_file.name} -> {output_file.name}")
    print(f"File size: {file_size_mb:.2f} MB")
    
    if file_size_mb > 10:
        print("⚠ Large file detected. This may take several minutes. Please be patient...")
    
    try:
        # Perform conversion
        print("Starting conversion (this may appear stuck at 0% for large files)...")
        convert(str(input_file), str(output_file))
        print(f"✓ Successfully converted to: {output_file}")
        return str(output_file)
    except Exception as e:
        print(f"✗ Error converting {input_file.name}: {str(e)}")
        raise


def batch_convert(input_folder, output_folder=None, recursive=False):
    """
    Convert all Word documents in a folder to PDF.
    
    Args:
        input_folder (str): Path to folder containing Word documents
        output_folder (str, optional): Path to output folder. If None, PDFs are saved in input folder
        recursive (bool): If True, search for Word files recursively in subfolders
    """
    input_dir = Path(input_folder)
    
    if not input_dir.exists() or not input_dir.is_dir():
        raise NotADirectoryError(f"Input folder not found: {input_folder}")
    
    # Find all Word documents
    if recursive:
        word_files = list(input_dir.rglob('*.docx')) + list(input_dir.rglob('*.doc'))
    else:
        word_files = list(input_dir.glob('*.docx')) + list(input_dir.glob('*.doc'))
    
    if not word_files:
        print(f"No Word documents found in: {input_folder}")
        return
    
    print(f"Found {len(word_files)} Word document(s) to convert")
    print("-" * 60)
    
    successful = 0
    failed = 0
    
    for word_file in word_files:
        try:
            # Determine output path
            if output_folder:
                output_dir = Path(output_folder)
                # Preserve subfolder structure if recursive
                if recursive:
                    relative_path = word_file.relative_to(input_dir)
                    output_file = output_dir / relative_path.with_suffix('.pdf')
                else:
                    output_file = output_dir / word_file.with_suffix('.pdf').name
            else:
                output_file = word_file.with_suffix('.pdf')
            
            convert_word_to_pdf(word_file, output_file)
            successful += 1
        except Exception as e:
            failed += 1
            continue
    
    print("-" * 60)
    print(f"Conversion complete: {successful} successful, {failed} failed")


def load_config(config_file='config.json'):
    """
    Load configuration from a JSON file.
    
    Args:
        config_file (str): Path to the configuration file
    
    Returns:
        dict: Configuration dictionary
    """
    config_path = Path(config_file)
    
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_file}")
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in configuration file: {str(e)}")


def run_from_config(config_file='config.json'):
    """
    Run conversion using settings from a configuration file.
    
    Args:
        config_file (str): Path to the configuration file
    """
    print(f"Loading configuration from: {config_file}")
    config = load_config(config_file)
    
    batch_mode = config.get('batch_mode', False)
    recursive = config.get('recursive', False)
    
    if batch_mode:
        # Batch conversion mode
        input_folder = config.get('input_folder', '')
        output_folder = config.get('output_folder', None)
        
        if not input_folder:
            raise ValueError("'input_folder' must be specified in config for batch mode")
        
        print(f"Batch Mode: {'Recursive' if recursive else 'Non-recursive'}")
        batch_convert(input_folder, output_folder, recursive)
    else:
        # Single file conversion mode
        input_file = config.get('input_file', '')
        output_file = config.get('output_file', None)
        
        if not input_file:
            raise ValueError("'input_file' must be specified in config")
        
        convert_word_to_pdf(input_file, output_file)


def main():
    """Main function to handle command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Convert Word documents to PDF without losing any content or formatting.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert using config file (default: config.json)
  python word_to_pdf.py --config
  
  # Convert using custom config file
  python word_to_pdf.py --config my_config.json
  
  # Convert a single file
  python word_to_pdf.py document.docx
  
  # Convert a file with custom output path
  python word_to_pdf.py document.docx -o output.pdf
  
  # Convert all Word files in a folder
  python word_to_pdf.py input_folder/ --batch
  
  # Convert all Word files recursively with output folder
  python word_to_pdf.py input_folder/ --batch -o output_folder/ --recursive
        """
    )
    
    parser.add_argument('input', nargs='?', help='Input Word file or folder path')
    parser.add_argument('-o', '--output', help='Output PDF file or folder path')
    parser.add_argument('--batch', action='store_true', help='Batch convert all Word files in a folder')
    parser.add_argument('--recursive', action='store_true', help='Search for Word files recursively in subfolders (use with --batch)')
    parser.add_argument('--config', nargs='?', const='config.json', metavar='CONFIG_FILE', 
                        help='Use configuration file (default: config.json)')
    
    args = parser.parse_args()
    
    try:
        # Check if config mode is requested
        if args.config is not None:
            run_from_config(args.config)
        elif args.input:
            # Command-line mode
            if args.batch:
                # Batch conversion mode
                batch_convert(args.input, args.output, args.recursive)
            else:
                # Single file conversion mode
                convert_word_to_pdf(args.input, args.output)
        else:
            parser.print_help()
            print("\nError: Either provide input file/folder or use --config option", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"\nError: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

