"""
Helper script to find recently created PDF files
Useful if you're not sure where your PDF was saved
"""

import os
from pathlib import Path
from datetime import datetime, timedelta

def find_recent_pdfs(directory, hours=1):
    """
    Find PDF files created in the last N hours
    
    Args:
        directory: Directory to search
        hours: How many hours back to search
    """
    directory = Path(directory)
    cutoff_time = datetime.now() - timedelta(hours=hours)
    
    print(f"Searching for PDF files created in the last {hours} hour(s)...")
    print(f"Location: {directory}")
    print("=" * 70)
    
    found_files = []
    
    # Search for PDF files
    for pdf_file in directory.rglob("*.pdf"):
        try:
            # Get file creation time
            creation_time = datetime.fromtimestamp(pdf_file.stat().st_ctime)
            
            if creation_time > cutoff_time:
                file_size_mb = pdf_file.stat().st_size / (1024 * 1024)
                found_files.append((pdf_file, creation_time, file_size_mb))
        except Exception as e:
            continue
    
    # Sort by creation time (newest first)
    found_files.sort(key=lambda x: x[1], reverse=True)
    
    if found_files:
        print(f"\nFound {len(found_files)} PDF file(s):\n")
        
        for i, (pdf_path, creation_time, size) in enumerate(found_files, 1):
            print(f"{i}. {pdf_path.name}")
            print(f"   Size: {size:.2f} MB")
            print(f"   Created: {creation_time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"   Location: {pdf_path}")
            print()
        
        return found_files
    else:
        print(f"\n❌ No PDF files found created in the last {hours} hour(s)")
        print(f"\nTry increasing the search time or checking a different directory.")
        return []


def open_file_location(file_path):
    """Open file location in Windows Explorer"""
    os.system(f'explorer /select,"{file_path}"')


if __name__ == "__main__":
    import sys
    
    # Default to current directory
    search_dir = Path(__file__).parent
    
    # Allow command line argument for directory
    if len(sys.argv) > 1:
        search_dir = Path(sys.argv[1])
    
    print("=" * 70)
    print("PDF Finder - Locate Recently Created PDFs")
    print("=" * 70)
    print()
    
    # Search for PDFs
    found = find_recent_pdfs(search_dir, hours=1)
    
    if found:
        print("=" * 70)
        print("\nWould you like to open the most recent PDF location?")
        
        try:
            response = input("Enter 'y' to open, or file number (1, 2, etc.): ").strip().lower()
            
            if response == 'y':
                # Open most recent
                open_file_location(found[0][0])
                print(f"✓ Opened location: {found[0][0].parent}")
            elif response.isdigit():
                idx = int(response) - 1
                if 0 <= idx < len(found):
                    open_file_location(found[idx][0])
                    print(f"✓ Opened location: {found[idx][0].parent}")
                else:
                    print("Invalid number")
        except KeyboardInterrupt:
            print("\n\nCancelled")
        except Exception as e:
            print(f"Error: {e}")
    
    print("\n" + "=" * 70)

