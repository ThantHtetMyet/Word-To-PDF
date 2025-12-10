"""
Advanced Word to PDF Converter with Better Layout Preservation
Uses direct COM interface with optimized settings for maintaining exact formatting
"""

import os
import sys
import json
import argparse
import time
from pathlib import Path
import win32com.client
import pythoncom


def convert_word_to_pdf_advanced(input_path, output_path=None):
    """
    Convert a Word document to PDF using direct COM interface with optimal settings.
    This method preserves images, drawings, and layout better than docx2pdf.
    
    Args:
        input_path (str): Path to the input Word document
        output_path (str, optional): Path for the output PDF
    
    Returns:
        str: Path to the generated PDF file
    """
    input_file = Path(input_path).resolve()
    
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
        output_file = Path(output_path).resolve()
        output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Get file size for user information
    file_size_mb = input_file.stat().st_size / (1024 * 1024)
    
    print(f"Converting: {input_file.name}")
    print(f"Output: {output_file.name}")
    print(f"File size: {file_size_mb:.2f} MB")
    
    if file_size_mb > 10:
        print("⚠ Large file detected. This may take several minutes. Please be patient...")
    
    # Initialize COM
    pythoncom.CoInitialize()
    word = None
    doc = None
    
    try:
        print("Opening Microsoft Word...")
        # Create Word application instance
        word = win32com.client.DispatchEx("Word.Application")
        
        # Configure Word for better performance and reliability
        word.Visible = False  # Run in background
        word.DisplayAlerts = 0  # Don't show alerts (wdAlertsNone = 0)
        
        print(f"Opening document...")
        # Open the document
        doc = word.Documents.Open(str(input_file), ReadOnly=True)
        
        print("Converting to PDF (preserving all formatting, images, and drawings)...")
        
        # ExportAsFixedFormat parameters for best quality and layout preservation
        # wdExportFormatPDF = 17
        # wdExportOptimizeForPrint = 0 (best quality)
        # wdExportAllDocument = 0
        # wdExportDocumentContent = 0
        # wdExportCreateNoBookmarks = 0
        # wdExportCreateHeadingBookmarks = 1
        
        doc.ExportAsFixedFormat(
            OutputFileName=str(output_file),
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,  # wdExportOptimizeForPrint (best quality)
            Range=0,  # wdExportAllDocument
            From=1,
            To=1,
            Item=0,  # wdExportDocumentContent
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )
        
        print(f"✓ Successfully converted to: {output_file}")
        print(f"✓ All images, drawings, and formatting preserved!")
        
        return str(output_file)
        
    except Exception as e:
        error_msg = str(e)
        print(f"\n✗ Error during conversion: {error_msg}")
        
        # Provide helpful error messages
        if "0x800A03EC" in error_msg or "Command failed" in error_msg:
            print("\n⚠ Troubleshooting suggestions:")
            print("  1. Close any open Word documents and try again")
            print("  2. Check if the document is password-protected or corrupted")
            print("  3. Try opening the document in Word manually first")
            print("  4. Restart your computer if the issue persists")
        elif "0x80010001" in error_msg:
            print("\n⚠ Word COM interface is busy. Please:")
            print("  1. Close all Word windows")
            print("  2. Wait a moment and try again")
        
        raise
        
    finally:
        # Clean up COM objects
        print("Cleaning up...")
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        if word:
            try:
                word.Quit()
            except:
                pass
        
        # Give Word time to fully close
        time.sleep(1)
        
        # Uninitialize COM
        pythoncom.CoUninitialize()


def batch_convert_advanced(input_folder, output_folder=None, recursive=False):
    """
    Convert all Word documents in a folder to PDF using advanced method.
    
    Args:
        input_folder (str): Path to folder containing Word documents
        output_folder (str, optional): Path to output folder
        recursive (bool): If True, search for Word files recursively
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
    print("=" * 70)
    
    successful = 0
    failed = 0
    failed_files = []
    
    for i, word_file in enumerate(word_files, 1):
        print(f"\n[{i}/{len(word_files)}] Processing: {word_file.name}")
        print("-" * 70)
        
        try:
            # Determine output path
            if output_folder:
                output_dir = Path(output_folder)
                if recursive:
                    relative_path = word_file.relative_to(input_dir)
                    output_file = output_dir / relative_path.with_suffix('.pdf')
                else:
                    output_file = output_dir / word_file.with_suffix('.pdf').name
            else:
                output_file = word_file.with_suffix('.pdf')
            
            convert_word_to_pdf_advanced(word_file, output_file)
            successful += 1
            
        except Exception as e:
            failed += 1
            failed_files.append(word_file.name)
            print(f"✗ Failed: {word_file.name}")
            continue
    
    print("\n" + "=" * 70)
    print(f"Batch conversion complete:")
    print(f"  ✓ Successful: {successful}")
    print(f"  ✗ Failed: {failed}")
    
    if failed_files:
        print(f"\nFailed files:")
        for fname in failed_files:
            print(f"  - {fname}")


def load_config(config_file='config.json'):
    """Load configuration from a JSON file."""
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
    """Run conversion using settings from a configuration file."""
    print(f"Loading configuration from: {config_file}")
    config = load_config(config_file)
    
    batch_mode = config.get('batch_mode', False)
    recursive = config.get('recursive', False)
    
    if batch_mode:
        input_folder = config.get('input_folder', '')
        output_folder = config.get('output_folder', None)
        
        if not input_folder:
            raise ValueError("'input_folder' must be specified in config for batch mode")
        
        print(f"Batch Mode: {'Recursive' if recursive else 'Non-recursive'}")
        batch_convert_advanced(input_folder, output_folder, recursive)
    else:
        input_file = config.get('input_file', '')
        output_file = config.get('output_file', None)
        
        if not input_file:
            raise ValueError("'input_file' must be specified in config")
        
        convert_word_to_pdf_advanced(input_file, output_file)


def main():
    """Main function to handle command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Advanced Word to PDF Converter with Perfect Layout Preservation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
This advanced converter uses Word's COM interface directly with optimized settings
to ensure images, drawings, and formatting remain exactly as in the original document.

Examples:
  # Convert using config file
  python word_to_pdf_advanced.py --config
  
  # Convert a single file
  python word_to_pdf_advanced.py document.docx
  
  # Convert with custom output
  python word_to_pdf_advanced.py document.docx -o output.pdf
  
  # Batch convert folder
  python word_to_pdf_advanced.py input_folder/ --batch
        """
    )
    
    parser.add_argument('input', nargs='?', help='Input Word file or folder path')
    parser.add_argument('-o', '--output', help='Output PDF file or folder path')
    parser.add_argument('--batch', action='store_true', help='Batch convert all Word files in a folder')
    parser.add_argument('--recursive', action='store_true', help='Search for Word files recursively')
    parser.add_argument('--config', nargs='?', const='config.json', metavar='CONFIG_FILE',
                        help='Use configuration file (default: config.json)')
    
    args = parser.parse_args()
    
    try:
        if args.config is not None:
            run_from_config(args.config)
        elif args.input:
            if args.batch:
                batch_convert_advanced(args.input, args.output, args.recursive)
            else:
                convert_word_to_pdf_advanced(args.input, args.output)
        else:
            parser.print_help()
            print("\nError: Either provide input file/folder or use --config option", file=sys.stderr)
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n⚠ Conversion cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

