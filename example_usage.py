"""
Example usage of the Word to PDF converter as a Python module
"""

from word_to_pdf import convert_word_to_pdf, batch_convert

# Example 1: Convert a single file
def example_single_file():
    """Convert a single Word document to PDF"""
    try:
        input_file = "sample.docx"
        output_file = convert_word_to_pdf(input_file)
        print(f"PDF created: {output_file}")
    except Exception as e:
        print(f"Error: {e}")


# Example 2: Convert with custom output path
def example_custom_output():
    """Convert with a specific output path"""
    try:
        input_file = "document.docx"
        output_path = "output/converted_document.pdf"
        output_file = convert_word_to_pdf(input_file, output_path)
        print(f"PDF created: {output_file}")
    except Exception as e:
        print(f"Error: {e}")


# Example 3: Batch convert all files in a folder
def example_batch_conversion():
    """Batch convert all Word files in a folder"""
    try:
        input_folder = "input_documents"
        output_folder = "output_pdfs"
        batch_convert(input_folder, output_folder, recursive=False)
    except Exception as e:
        print(f"Error: {e}")


# Example 4: Recursive batch conversion
def example_recursive_conversion():
    """Recursively convert all Word files in folder and subfolders"""
    try:
        input_folder = "all_documents"
        output_folder = "all_pdfs"
        batch_convert(input_folder, output_folder, recursive=True)
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    print("Word to PDF Converter - Example Usage")
    print("=" * 50)
    
    # Uncomment the example you want to run:
    
    # example_single_file()
    # example_custom_output()
    # example_batch_conversion()
    # example_recursive_conversion()
    
    print("\nNote: Make sure to have Word documents ready before running examples!")

