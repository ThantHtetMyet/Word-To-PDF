# How to Build the EXE File

## Quick Start (Easiest Method)

### Step 1: Install Requirements

```bash
pip install -r requirements_gui.txt
```

### Step 2: Build the EXE

**On Windows, simply double-click:**
```
build_exe.bat
```

**Or run in PowerShell/CMD:**
```bash
build_exe.bat
```

That's it! The EXE will be created in the `dist` folder.

---

## What You Get

After building, you'll find:

```
dist/
  └── WordToPDF-Converter.exe  ← Your standalone application!
```

This is a **single executable file** that:
- ✅ Can run on any Windows computer
- ✅ Doesn't require Python to be installed
- ✅ Includes all dependencies
- ✅ Has a beautiful GUI interface
- ✅ Preserves perfect layout and formatting

---

## Running the Application

### For You (Development)
Simply run:
```bash
python word_to_pdf_gui.py
```

### For Others (Distribution)
1. Send them `WordToPDF-Converter.exe`
2. They double-click it
3. Done!

**Important:** Microsoft Word must be installed on the computer where the .exe runs.

---

## GUI Features

The application has a modern, user-friendly interface with:

### 📁 File Selection
- Browse button to select Word documents
- Shows full file path
- Supports .docx and .doc files

### 💾 Output Configuration
- Auto-generate PDF filename (same location as input)
- Or manually specify custom output location
- Browse to select specific folder and filename

### 🚀 Conversion Process
- Large "Convert to PDF" button
- Real-time progress bar
- Status messages showing what's happening
- Progress updates for large files

### ✅ Success Actions
After conversion completes:
- Success message with file location
- "Open Output Folder" button
- "Open PDF" button to view result immediately

### 🎨 Modern Design
- Clean, professional interface
- Color-coded buttons
- Progress indicators
- Helpful error messages
- Responsive layout

---

## Manual Build (Alternative Method)

If the batch file doesn't work, build manually:

```bash
# Install dependencies
pip install pyinstaller pywin32

# Build the EXE
pyinstaller --onefile --windowed --name="WordToPDF-Converter" word_to_pdf_gui.py
```

---

## Build Options Explained

The build script uses these PyInstaller options:

| Option | Description |
|--------|-------------|
| `--onefile` | Create a single .exe file (not a folder) |
| `--windowed` | No console window (GUI only) |
| `--name` | Name of the output executable |
| `--icon` | Custom icon (optional) |
| `--hidden-import` | Include COM libraries |

---

## Troubleshooting Build Issues

### Issue: "pyinstaller not found"

**Solution:**
```bash
pip install pyinstaller
```

### Issue: "pywin32 not found"

**Solution:**
```bash
pip install pywin32
```

### Issue: Build succeeds but .exe doesn't work

**Possible causes:**
1. **Antivirus blocking** - Add exception for the .exe
2. **Missing Word** - Microsoft Word must be installed
3. **32-bit vs 64-bit** - Build on same architecture as target system

**Solution:**
```bash
# Rebuild with console window to see errors
pyinstaller --onefile --name="WordToPDF-Converter" word_to_pdf_gui.py
```

### Issue: .exe is too large

The .exe will be around 15-25 MB due to Python and dependencies. This is normal.

To reduce size, you can use:
```bash
pyinstaller --onefile --windowed --strip --name="WordToPDF-Converter" word_to_pdf_gui.py
```

---

## Distribution

### Sharing Your Application

1. **Just the EXE:**
   - Send `WordToPDF-Converter.exe`
   - Size: ~15-25 MB
   - No installation needed

2. **With Instructions:**
   Create a simple README:
   ```
   Word to PDF Converter
   
   Requirements:
   - Windows 10/11
   - Microsoft Word installed
   
   How to Use:
   1. Double-click WordToPDF-Converter.exe
   2. Click "Browse" to select your Word file
   3. Click "Convert to PDF"
   4. Done!
   ```

3. **Professional Package:**
   Create a folder with:
   ```
   WordToPDF-Converter/
     ├── WordToPDF-Converter.exe
     ├── README.txt (instructions)
     └── LICENSE.txt (optional)
   ```

### System Requirements for Users

- **OS:** Windows 10 or 11
- **Required:** Microsoft Word (any recent version)
- **RAM:** 2 GB minimum (4 GB recommended for large files)
- **Disk:** 100 MB free space

---

## Advanced Build Configuration

### Custom Icon

1. Create or download an icon file (`icon.ico`)
2. Place it in the same directory
3. The build script will automatically include it

### Version Information

Create a version file for professional appearance:

```bash
pyinstaller --onefile --windowed ^
    --name="WordToPDF-Converter" ^
    --icon=icon.ico ^
    --version-file=version.txt ^
    word_to_pdf_gui.py
```

### Reduce Build Time

For faster rebuilds during development:

```bash
# First build (full)
pyinstaller word_to_pdf_gui.py

# Subsequent builds (incremental)
pyinstaller --noconfirm word_to_pdf_gui.spec
```

---

## Testing the EXE

### Before Distribution

Test the .exe on a clean system or VM:

1. **Fresh Windows installation**
2. **Only Microsoft Word installed** (no Python)
3. **Run the .exe**
4. **Test with various file sizes:**
   - Small file (< 1 MB)
   - Medium file (5-10 MB)
   - Large file (20-50 MB)
5. **Test with complex documents:**
   - Images
   - Tables
   - Charts
   - Headers/Footers

### Test Checklist

- [ ] .exe runs without errors
- [ ] GUI loads properly
- [ ] File browser works
- [ ] Conversion completes successfully
- [ ] PDF opens correctly
- [ ] Layout is preserved
- [ ] Images are in correct positions
- [ ] No console window appears
- [ ] "Open Folder" button works
- [ ] "Open PDF" button works

---

## Deployment Options

### Option 1: Direct Distribution
Send the .exe file via:
- Email (if size permits)
- File sharing service (Dropbox, Google Drive, etc.)
- USB drive
- Network share

### Option 2: Installer Package
Create a professional installer with tools like:
- **Inno Setup** (free)
- **NSIS** (free)
- **Advanced Installer** (commercial)

Example Inno Setup script:
```ini
[Setup]
AppName=Word to PDF Converter
AppVersion=1.0
DefaultDirName={pf}\WordToPDF
DefaultGroupName=Word to PDF Converter
OutputDir=installer
OutputBaseFilename=WordToPDF_Setup

[Files]
Source: "dist\WordToPDF-Converter.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Word to PDF Converter"; Filename: "{app}\WordToPDF-Converter.exe"
Name: "{userdesktop}\Word to PDF Converter"; Filename: "{app}\WordToPDF-Converter.exe"
```

### Option 3: Portable Package
Create a ZIP file with:
```
WordToPDF_Portable.zip
  ├── WordToPDF-Converter.exe
  ├── README.txt
  └── INSTRUCTIONS.txt
```

---

## Updating the Application

When you make changes:

1. **Edit the code** (`word_to_pdf_gui.py`)
2. **Test the changes** (run with Python first)
3. **Rebuild the .exe** (run `build_exe.bat`)
4. **Test the new .exe**
5. **Distribute updated version**

---

## File Size Comparison

| Method | File Size | Requires Python |
|--------|-----------|-----------------|
| Python Script | ~5 KB | Yes |
| EXE (onefile) | ~15-25 MB | No |
| EXE + Dependencies (folder) | ~30-40 MB | No |

The single-file .exe is recommended for easier distribution.

---

## Legal Considerations

If distributing to others:
- Ensure compliance with Microsoft Word COM interface usage
- Include appropriate license information
- Add disclaimer about requiring Word installation
- Consider liability disclaimers

Example disclaimer:
```
This software requires Microsoft Word to be installed.
Microsoft Word is a registered trademark of Microsoft Corporation.
This software is not affiliated with or endorsed by Microsoft.
```

---

## Summary

✅ **To build EXE:** Run `build_exe.bat`
✅ **Output location:** `dist\WordToPDF-Converter.exe`
✅ **Distribution:** Send the .exe file
✅ **Requirement:** Microsoft Word must be installed
✅ **File size:** ~15-25 MB
✅ **Platform:** Windows only

That's it! You now have a professional, standalone Word to PDF converter application! 🎉

