# Configuration File Guide

## Overview

The Word to PDF converter now supports configuration files, allowing you to specify all conversion settings in a JSON file instead of using command-line arguments.

## Quick Start with Config File

1. **Edit the `config.json` file** with your file paths
2. **Run the converter:** `python word_to_pdf.py --config`

That's it! The converter will read all settings from the config file.

## Configuration File Format

The configuration file is in JSON format. Here's the structure:

### Single File Conversion

```json
{
  "input_file": "C:\\Users\\YourName\\Documents\\MyDocument.docx",
  "output_file": "C:\\Users\\YourName\\Documents\\MyDocument.pdf",
  "batch_mode": false,
  "recursive": false
}
```

### Batch Folder Conversion

```json
{
  "batch_mode": true,
  "recursive": false,
  "input_folder": "C:\\Users\\YourName\\Documents\\WordFiles",
  "output_folder": "C:\\Users\\YourName\\Documents\\PDFs"
}
```

### Recursive Batch Conversion

```json
{
  "batch_mode": true,
  "recursive": true,
  "input_folder": "C:\\Users\\YourName\\Documents\\AllWordFiles",
  "output_folder": "C:\\Users\\YourName\\Documents\\AllPDFs"
}
```

## Configuration Options

| Option | Type | Description |
|--------|------|-------------|
| `input_file` | string | Full path to the Word document to convert (for single file mode) |
| `output_file` | string | Full path for the output PDF (optional, defaults to same name/location) |
| `batch_mode` | boolean | `true` for folder conversion, `false` for single file |
| `recursive` | boolean | `true` to process subfolders, `false` for current folder only |
| `input_folder` | string | Full path to folder containing Word files (for batch mode) |
| `output_folder` | string | Full path to output folder (optional, defaults to input folder) |

## Usage Examples

### Example 1: Using Default Config File

Create/edit `config.json` in the same directory as the script:

```json
{
  "input_file": "C:\\Work\\Report.docx",
  "output_file": "C:\\Work\\Report.pdf",
  "batch_mode": false
}
```

Run:
```bash
python word_to_pdf.py --config
```

### Example 2: Using Custom Config File

Create a custom config file (e.g., `my_conversion.json`):

```json
{
  "batch_mode": true,
  "input_folder": "C:\\Projects\\Documents",
  "output_folder": "C:\\Projects\\PDFs",
  "recursive": true
}
```

Run:
```bash
python word_to_pdf.py --config my_conversion.json
```

### Example 3: Multiple Config Files for Different Projects

Create separate config files for different projects:

**project_a_config.json:**
```json
{
  "input_file": "C:\\ProjectA\\specs.docx",
  "output_file": "C:\\ProjectA\\PDF\\specs.pdf",
  "batch_mode": false
}
```

**project_b_config.json:**
```json
{
  "batch_mode": true,
  "input_folder": "C:\\ProjectB\\docs",
  "output_folder": "C:\\ProjectB\\pdfs",
  "recursive": true
}
```

Run them:
```bash
python word_to_pdf.py --config project_a_config.json
python word_to_pdf.py --config project_b_config.json
```

## Path Formats

Windows paths can be specified in two ways:

### Option 1: Double Backslashes (Recommended)
```json
{
  "input_file": "C:\\Users\\John\\Documents\\file.docx"
}
```

### Option 2: Forward Slashes (Also works on Windows)
```json
{
  "input_file": "C:/Users/John/Documents/file.docx"
}
```

## Tips

1. **Use `config_template.json` as a starting point** - Copy it to `config.json` and modify
2. **Test with a single file first** before batch converting
3. **Leave output paths empty** to save PDFs in the same location as source files
4. **Create multiple config files** for different conversion tasks
5. **Use comments** (lines starting with `_`) in JSON for documentation (they're ignored)

## Command-Line Override

You can still use command-line arguments even when config files exist:

```bash
# This ignores config.json and uses command-line args
python word_to_pdf.py myfile.docx -o output.pdf
```

## Error Handling

If the config file has errors, you'll see helpful messages:

- **File not found:** Check the path to your config file
- **Invalid JSON:** Check for syntax errors (missing commas, quotes, etc.)
- **Missing required fields:** Ensure `input_file` or `input_folder` is specified
- **Invalid paths:** Check that input files/folders exist

## Sample Configs for Common Scenarios

### Scenario 1: Daily Report Conversion
```json
{
  "input_file": "C:\\Reports\\Daily_Report.docx",
  "output_file": "C:\\Reports\\Archive\\Daily_Report.pdf",
  "batch_mode": false
}
```

### Scenario 2: Process Entire Department Folder
```json
{
  "batch_mode": true,
  "recursive": true,
  "input_folder": "C:\\Shared\\Department\\Documents",
  "output_folder": "C:\\Shared\\Department\\PDFs"
}
```

### Scenario 3: Client Presentation (specific file)
```json
{
  "input_file": "C:\\Clients\\ABC Corp\\Presentation_2025.docx",
  "output_file": "C:\\Clients\\ABC Corp\\Presentation_2025_Final.pdf",
  "batch_mode": false
}
```

