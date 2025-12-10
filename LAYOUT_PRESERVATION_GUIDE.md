# Layout Preservation Guide

## Problem: Images and Drawings Moving in PDF

If you're experiencing issues where images, drawings, or other objects shift position when converting Word to PDF, use the **advanced converter** which provides better layout preservation.

## Solution: Use the Advanced Converter

### Quick Start

```bash
python word_to_pdf_advanced.py --config
```

This uses the same `config.json` file but with optimized Word export settings.

## Why the Advanced Version is Better

The advanced converter (`word_to_pdf_advanced.py`) uses Microsoft Word's COM interface directly with these optimizations:

### 1. **OptimizeForPrint Mode**
- Uses highest quality settings
- Preserves exact positioning of all elements
- Maintains image resolution and quality

### 2. **Better Image Handling**
- `BitmapMissingFonts=True` - Ensures fonts render correctly
- `DocStructureTags=True` - Preserves document structure
- No compression or optimization that could shift layouts

### 3. **Enhanced Error Handling**
- Better error messages
- Automatic cleanup
- Retry suggestions

### 4. **Full Formatting Preservation**
- Maintains all document properties
- Keeps bookmarks and headings
- Preserves embedded objects exactly as positioned

## Comparison

| Feature | Original (docx2pdf) | Advanced (direct COM) |
|---------|---------------------|----------------------|
| Speed | Fast | Moderate |
| Layout Accuracy | Good | Excellent |
| Image Preservation | Good | Perfect |
| Drawing Objects | May shift | Exact position |
| Complex Documents | May fail | Better handling |
| Error Recovery | Basic | Advanced |
| Large Files | May timeout | Optimized |

## Usage Examples

### Using Config File (Same as Before)

Your `config.json` works the same way:

```json
{
  "input_file": "C:\\path\\to\\document.docx",
  "output_file": "C:\\path\\to\\output.pdf",
  "batch_mode": false
}
```

Run:
```bash
python word_to_pdf_advanced.py --config
```

### Command Line

```bash
# Single file
python word_to_pdf_advanced.py "your_document.docx"

# With custom output
python word_to_pdf_advanced.py "input.docx" -o "output.pdf"

# Batch convert
python word_to_pdf_advanced.py "folder/" --batch
```

## Troubleshooting

### Issue: "Command failed" Error

**Solution:**
1. Close all Word documents
2. Wait 10 seconds
3. Try again

### Issue: Still Getting Layout Shifts

**Possible causes:**
1. **Document compatibility mode** - Open in Word, save as newest .docx format
2. **Floating objects** - In Word, check if images are set to "In Line with Text" vs "Floating"
3. **Embedded objects** - Some embedded objects may need to be converted to images first

**To fix floating objects in Word:**
1. Select the image/drawing
2. Right-click → Wrap Text → In Line with Text
3. Or: Right-click → Format Picture → Layout → In Line with Text

### Issue: Conversion is Slow

This is normal for large files with many images. The advanced method prioritizes accuracy over speed.

**Speed tips:**
- The script shows file size and progress
- Large files (30+ MB) can take 5-10 minutes
- Don't interrupt - let it complete

## Best Practices for Perfect Conversion

1. **Save Word document properly first**
   - File → Save As → Save as newest .docx format
   - Ensure document isn't corrupted

2. **Check document before converting**
   - Open in Word to ensure it displays correctly
   - Fix any Word errors or warnings first

3. **For critical documents**
   - Make a backup copy first
   - Test with a small section first
   - Compare Word and PDF side-by-side after conversion

4. **Complex layouts**
   - Documents with many floating objects work best with "In Line with Text" setting
   - Tables, charts, and embedded Excel sheets are handled automatically

## When to Use Each Version

### Use `word_to_pdf_advanced.py` when:
- ✅ Layout accuracy is critical
- ✅ Document has images, drawings, or diagrams
- ✅ Original converter failed or shifted content
- ✅ Professional/client-facing documents
- ✅ Complex formatting must be preserved exactly

### Use `word_to_pdf.py` when:
- ✅ Speed is more important than perfect accuracy
- ✅ Simple text documents with minimal formatting
- ✅ Batch converting many simple files
- ✅ Quick previews or drafts

## Technical Details

The advanced converter uses Word's `ExportAsFixedFormat` method with these parameters:

```python
doc.ExportAsFixedFormat(
    OutputFileName=output_path,
    ExportFormat=17,              # PDF format
    OptimizeFor=0,                # Print quality (best)
    IncludeDocProps=True,         # Keep metadata
    BitmapMissingFonts=True,      # Handle fonts correctly
    DocStructureTags=True,        # Preserve structure
    CreateBookmarks=1             # Keep navigation
)
```

This is the same method Word uses internally for "Save as PDF", ensuring identical results to manual conversion.

