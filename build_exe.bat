@echo off
echo ========================================
echo Word to PDF Converter - EXE Builder
echo ========================================
echo.

echo [1/4] Installing required packages...
pip install -r requirements_gui.txt
echo.

echo [2/4] Creating build directory...
if not exist "build" mkdir build
if not exist "dist" mkdir dist
echo.

echo [3/4] Building executable with PyInstaller...
echo This may take a few minutes...
echo.

pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name="WordToPDF-Converter" ^
    --icon=icon.ico ^
    --add-data="icon.ico;." ^
    --hidden-import=win32timezone ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    word_to_pdf_gui.py

echo.
echo [4/4] Cleaning up...
echo.

if exist "dist\WordToPDF-Converter.exe" (
    echo ========================================
    echo SUCCESS! 
    echo ========================================
    echo.
    echo Your executable has been created:
    echo   Location: dist\WordToPDF-Converter.exe
    echo.
    echo You can now:
    echo   1. Run the .exe file directly
    echo   2. Copy it to any Windows computer
    echo   3. Share it with others
    echo.
    echo Note: Microsoft Word must be installed
    echo       on the computer where you run this.
    echo.
    echo ========================================
    
    explorer dist
) else (
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
    echo Please check the error messages above.
    echo.
)

pause

