"""
Word to PDF Converter - GUI Application
Modern interface with drag-and-drop, progress tracking, and visual feedback
Uses the advanced conversion method for perfect layout preservation
"""

import os
import sys
import time
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import pythoncom


class WordToPDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to PDF Converter")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        # Set icon (if available)
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready to convert")
        self.is_converting = False
        
        # Configure colors
        self.bg_color = "#f0f0f0"
        self.primary_color = "#4CAF50"
        self.secondary_color = "#2196F3"
        self.danger_color = "#f44336"
        self.text_color = "#333333"
        
        self.root.configure(bg=self.bg_color)
        
        self.create_widgets()
    
    def create_widgets(self):
        """Create and layout all GUI widgets"""
        
        # Title
        title_frame = tk.Frame(self.root, bg=self.primary_color, height=80)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📄 Word to PDF Converter",
            font=("Segoe UI", 24, "bold"),
            bg=self.primary_color,
            fg="white"
        )
        title_label.pack(pady=20)
        
        subtitle_label = tk.Label(
            title_frame,
            text="Convert Word documents to PDF with perfect layout preservation",
            font=("Segoe UI", 10),
            bg=self.primary_color,
            fg="white"
        )
        subtitle_label.pack()
        
        # Main content frame
        content_frame = tk.Frame(self.root, bg=self.bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Input file section
        input_label = tk.Label(
            content_frame,
            text="Select Word Document:",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg=self.text_color
        )
        input_label.pack(anchor=tk.W, pady=(10, 5))
        
        input_frame = tk.Frame(content_frame, bg=self.bg_color)
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.input_entry = tk.Entry(
            input_frame,
            textvariable=self.input_file,
            font=("Segoe UI", 10),
            relief=tk.SOLID,
            borderwidth=1
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        browse_input_btn = tk.Button(
            input_frame,
            text="Browse...",
            command=self.browse_input_file,
            font=("Segoe UI", 10, "bold"),
            bg=self.secondary_color,
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=20,
            pady=8
        )
        browse_input_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # Output file section
        output_label = tk.Label(
            content_frame,
            text="Save PDF As:",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg=self.text_color
        )
        output_label.pack(anchor=tk.W, pady=(10, 5))
        
        output_frame = tk.Frame(content_frame, bg=self.bg_color)
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_file,
            font=("Segoe UI", 10),
            relief=tk.SOLID,
            borderwidth=1
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        browse_output_btn = tk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_file,
            font=("Segoe UI", 10, "bold"),
            bg=self.secondary_color,
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=20,
            pady=8
        )
        browse_output_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # Auto-fill checkbox
        self.auto_output = tk.BooleanVar(value=True)
        auto_check = tk.Checkbutton(
            content_frame,
            text="Auto-generate output filename (same location as input)",
            variable=self.auto_output,
            command=self.toggle_auto_output,
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg=self.text_color,
            activebackground=self.bg_color,
            selectcolor=self.bg_color
        )
        auto_check.pack(anchor=tk.W, pady=(0, 20))
        
        # Convert button
        self.convert_btn = tk.Button(
            content_frame,
            text="🚀 Convert to PDF",
            command=self.start_conversion,
            font=("Segoe UI", 14, "bold"),
            bg=self.primary_color,
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=40,
            pady=15
        )
        self.convert_btn.pack(pady=20)
        
        # Progress section
        progress_frame = tk.Frame(content_frame, bg=self.bg_color)
        progress_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='indeterminate',
            length=400
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Status label
        self.status_label = tk.Label(
            progress_frame,
            textvariable=self.status_text,
            font=("Segoe UI", 10),
            bg=self.bg_color,
            fg=self.text_color,
            wraplength=600
        )
        self.status_label.pack()
        
        # Result buttons frame (hidden initially)
        self.result_frame = tk.Frame(content_frame, bg=self.bg_color)
        
        self.open_folder_btn = tk.Button(
            self.result_frame,
            text="📁 Open Output Folder",
            command=self.open_output_folder,
            font=("Segoe UI", 10, "bold"),
            bg=self.secondary_color,
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.open_folder_btn.pack(side=tk.LEFT, padx=5)
        
        self.open_pdf_btn = tk.Button(
            self.result_frame,
            text="📄 Open PDF",
            command=self.open_pdf_file,
            font=("Segoe UI", 10, "bold"),
            bg=self.primary_color,
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.open_pdf_btn.pack(side=tk.LEFT, padx=5)
        
        # Footer
        footer_label = tk.Label(
            self.root,
            text="© 2025 | Preserves all images, drawings, and formatting perfectly",
            font=("Segoe UI", 8),
            bg=self.bg_color,
            fg="#666666"
        )
        footer_label.pack(side=tk.BOTTOM, pady=10)
        
        # Configure styles
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "TProgressbar",
            troughcolor=self.bg_color,
            background=self.primary_color,
            thickness=20
        )
        
        # Initially disable output entry if auto-fill is checked
        if self.auto_output.get():
            self.output_entry.config(state='disabled')
            browse_output_btn.config(state='disabled')
    
    def toggle_auto_output(self):
        """Toggle auto-output filename generation"""
        if self.auto_output.get():
            self.output_entry.config(state='disabled')
            self.output_file.set("")
        else:
            self.output_entry.config(state='normal')
    
    def browse_input_file(self):
        """Open file dialog to select input Word file"""
        filename = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[
                ("Word Documents", "*.docx *.doc"),
                ("All Files", "*.*")
            ]
        )
        
        if filename:
            self.input_file.set(filename)
            
            # Auto-generate output filename if enabled
            if self.auto_output.get():
                output = str(Path(filename).with_suffix('.pdf'))
                self.output_file.set(output)
    
    def browse_output_file(self):
        """Open file dialog to select output PDF location"""
        if not self.input_file.get():
            messagebox.showwarning("No Input File", "Please select an input file first.")
            return
        
        initial_name = Path(self.input_file.get()).stem + ".pdf"
        initial_dir = str(Path(self.input_file.get()).parent)
        
        filename = filedialog.asksaveasfilename(
            title="Save PDF As",
            initialfile=initial_name,
            initialdir=initial_dir,
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if filename:
            self.output_file.set(filename)
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        if self.is_converting:
            return
        
        # Validate input
        if not self.input_file.get():
            messagebox.showerror("Error", "Please select a Word document to convert.")
            return
        
        input_path = Path(self.input_file.get())
        if not input_path.exists():
            messagebox.showerror("Error", "Input file does not exist.")
            return
        
        # Generate output path if auto-fill is enabled
        if self.auto_output.get():
            output_path = input_path.with_suffix('.pdf')
            self.output_file.set(str(output_path))
        elif not self.output_file.get():
            messagebox.showerror("Error", "Please specify an output location.")
            return
        
        # Verify output directory is writable
        output_test_path = Path(self.output_file.get())
        if not output_test_path.parent.exists():
            try:
                output_test_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory: {str(e)}")
                return
        
        # Hide result buttons
        self.result_frame.pack_forget()
        
        # Start conversion in separate thread
        self.is_converting = True
        self.convert_btn.config(state='disabled', bg="#cccccc")
        self.progress_bar.start(10)
        
        thread = threading.Thread(target=self.convert_file, daemon=True)
        thread.start()
    
    def convert_file(self):
        """Convert Word file to PDF (runs in separate thread)"""
        try:
            input_path = Path(self.input_file.get())
            output_path = Path(self.output_file.get())
            
            # Update status
            self.update_status(f"Opening {input_path.name}...")
            
            # Get file size
            file_size_mb = input_path.stat().st_size / (1024 * 1024)
            
            if file_size_mb > 10:
                self.update_status(f"Converting large file ({file_size_mb:.1f} MB)... This may take several minutes.")
            else:
                self.update_status(f"Converting {input_path.name}...")
            
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            word = None
            doc = None
            
            try:
                # Create Word application
                self.update_status("Starting Microsoft Word...")
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                
                # Open document
                self.update_status(f"Opening document...")
                doc = word.Documents.Open(str(input_path.resolve()), ReadOnly=True)
                
                # Convert to PDF
                self.update_status("Converting to PDF with optimized settings...")
                self.update_status("Preserving all images, drawings, and formatting...")
                
                # Create output directory if needed
                output_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Export as PDF with best quality settings
                doc.ExportAsFixedFormat(
                    OutputFileName=str(output_path.resolve()),
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
                
                self.update_status("Cleaning up...")
                
            finally:
                # Clean up COM objects properly to release file locks
                self.update_status("Releasing file locks and cleaning up...")
                
                # Close document
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                        # Release COM reference
                        doc = None
                    except Exception as e:
                        print(f"Warning: Error closing document: {e}")
                
                # Quit Word
                if word:
                    try:
                        word.Quit()
                        # Release COM reference
                        word = None
                    except Exception as e:
                        print(f"Warning: Error quitting Word: {e}")
                
                # Force garbage collection to release COM objects
                import gc
                gc.collect()
                
                # Give Word time to fully close and release files
                time.sleep(2)
                
                # Uninitialize COM
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
            
            # Verify the PDF was actually created
            if output_path.exists():
                file_size = output_path.stat().st_size / (1024 * 1024)
                self.conversion_complete(str(output_path), file_size)
            else:
                raise FileNotFoundError(f"PDF was not created at expected location: {output_path}")
            
        except Exception as e:
            self.conversion_error(str(e))
    
    def update_status(self, message):
        """Update status message (thread-safe)"""
        self.root.after(0, lambda: self.status_text.set(message))
    
    def conversion_complete(self, output_path, file_size=0):
        """Handle successful conversion (thread-safe)"""
        def update_ui():
            self.progress_bar.stop()
            self.is_converting = False
            self.convert_btn.config(state='normal', bg=self.primary_color)
            
            # Show file size in status
            size_text = f" ({file_size:.2f} MB)" if file_size > 0 else ""
            self.status_text.set(f"✅ Conversion successful!{size_text}\nPDF saved to:\n{output_path}")
            
            # Show result buttons
            self.result_frame.pack(pady=10)
            
            # Show success message with full path
            messagebox.showinfo(
                "Success!",
                f"Your document has been converted successfully!\n\n"
                f"File: {Path(output_path).name}\n"
                f"Size: {file_size:.2f} MB{size_text}\n\n"
                f"Location:\n{output_path}"
            )
        
        self.root.after(0, update_ui)
    
    def conversion_error(self, error_message):
        """Handle conversion error (thread-safe)"""
        def update_ui():
            self.progress_bar.stop()
            self.is_converting = False
            self.convert_btn.config(state='normal', bg=self.primary_color)
            self.status_text.set(f"❌ Conversion failed")
            
            # Show error dialog with helpful suggestions
            error_text = "Conversion failed with the following error:\n\n"
            error_text += error_message + "\n\n"
            error_text += "Troubleshooting suggestions:\n"
            error_text += "• Close all Word documents and try again\n"
            error_text += "• Ensure the file is not corrupted\n"
            error_text += "• Check if the file is password-protected\n"
            error_text += "• Try opening the file in Word first\n"
            error_text += "• Restart your computer if the issue persists"
            
            messagebox.showerror("Conversion Error", error_text)
        
        self.root.after(0, update_ui)
    
    def open_output_folder(self):
        """Open the folder containing the output PDF"""
        if self.output_file.get():
            output_path = Path(self.output_file.get())
            if output_path.exists():
                # Open folder and select the file
                os.system(f'explorer /select,"{output_path}"')
            elif output_path.parent.exists():
                # If file doesn't exist but folder does, open folder
                os.startfile(output_path.parent)
                messagebox.showwarning(
                    "File Not Found",
                    f"The PDF file was not found at:\n{output_path}\n\n"
                    f"Opening the folder instead."
                )
            else:
                messagebox.showerror(
                    "Location Not Found",
                    f"Neither the file nor folder exists:\n{output_path}"
                )
    
    def open_pdf_file(self):
        """Open the generated PDF file"""
        if self.output_file.get():
            output_path = Path(self.output_file.get())
            if output_path.exists():
                try:
                    os.startfile(str(output_path))
                except Exception as e:
                    messagebox.showerror(
                        "Cannot Open PDF",
                        f"Failed to open PDF file:\n{str(e)}\n\n"
                        f"File location:\n{output_path}"
                    )
            else:
                messagebox.showerror(
                    "File Not Found",
                    f"The PDF file does not exist:\n{output_path}\n\n"
                    f"The conversion may have failed or the file was moved."
                )


def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    app = WordToPDFConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

