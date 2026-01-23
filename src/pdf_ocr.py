"""
PDF OCR Tool

Add OCR (text recognition) to PDF files and convert images to searchable PDFs.
Supports PDF files and image files (JPG, PNG, TIFF, BMP).

Copyright 2025-2026 Andre Lorbach

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import os
import sys
import glob
import shutil
import re
import subprocess
from pathlib import Path
from typing import List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk

# Drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("[WARNING] tkinterdnd2 not available, drag and drop disabled")

# OCR support
try:
    import ocrmypdf
    OCRMYPDF_AVAILABLE = True
except ImportError:
    OCRMYPDF_AVAILABLE = False
    print("[WARNING] ocrmypdf not available. Install with: pip install ocrmypdf")

# Image processing
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("[WARNING] Pillow not available. Install with: pip install pillow")

# Image to PDF conversion
try:
    import img2pdf
    IMG2PDF_AVAILABLE = True
except ImportError:
    IMG2PDF_AVAILABLE = False
    print("[WARNING] img2pdf not available. Install with: pip install img2pdf")


# ============================================================================
# Modern UI Styling Constants
# ============================================================================

class UIColors:
    """Modern color palette for consistent UI styling."""
    PRIMARY = "#2563eb"
    PRIMARY_HOVER = "#1d4ed8"
    PRIMARY_LIGHT = "#dbeafe"
    
    SECONDARY = "#64748b"
    SECONDARY_HOVER = "#475569"
    
    SUCCESS = "#16a34a"
    SUCCESS_LIGHT = "#dcfce7"
    SUCCESS_HOVER = "#15803d"
    ERROR = "#dc2626"
    ERROR_LIGHT = "#fee2e2"
    ERROR_HOVER = "#b91c1c"
    WARNING = "#f59e0b"
    WARNING_LIGHT = "#fef3c7"
    
    BG_PRIMARY = "#ffffff"
    BG_SECONDARY = "#f8fafc"
    BG_TERTIARY = "#f1f5f9"
    BORDER = "#e2e8f0"
    BORDER_DARK = "#cbd5e1"
    TEXT_PRIMARY = "#1e293b"
    TEXT_SECONDARY = "#64748b"
    TEXT_MUTED = "#94a3b8"
    
    DROP_ZONE_BG = "#f8fafc"
    DROP_ZONE_BORDER = "#94a3b8"
    DROP_ZONE_ACTIVE = "#dbeafe"
    DROP_ZONE_BORDER_ACTIVE = "#2563eb"


class UIFonts:
    """Font configurations for consistent typography."""
    TITLE = ("Segoe UI", 18, "bold")
    SUBTITLE = ("Segoe UI", 14, "bold")
    HEADING = ("Segoe UI", 12, "bold")
    BODY = ("Segoe UI", 10)
    BODY_BOLD = ("Segoe UI", 10, "bold")
    SMALL = ("Segoe UI", 9)
    SMALL_BOLD = ("Segoe UI", 9, "bold")
    MONO = ("Consolas", 9)
    BUTTON = ("Segoe UI", 10, "bold")
    BUTTON_SMALL = ("Segoe UI", 9)


class UISpacing:
    """Consistent spacing values."""
    XS = 2
    SM = 5
    MD = 10
    LG = 15
    XL = 20
    XXL = 30


def create_rounded_button(parent, text, command, style="primary", width=None):
    """Create a styled button with hover effect."""
    colors = {
        "primary": (UIColors.PRIMARY, UIColors.PRIMARY_HOVER, "#ffffff"),
        "secondary": (UIColors.BG_TERTIARY, UIColors.BORDER, UIColors.TEXT_PRIMARY),
        "success": (UIColors.SUCCESS, UIColors.SUCCESS_HOVER, "#ffffff"),
        "danger": (UIColors.ERROR, UIColors.ERROR_HOVER, "#ffffff"),
        "ghost": (UIColors.BG_PRIMARY, UIColors.BG_TERTIARY, UIColors.TEXT_PRIMARY),
    }
    
    bg, hover_bg, fg = colors.get(style, colors["primary"])
    
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        font=UIFonts.BUTTON,
        bg=bg,
        fg=fg,
        activebackground=hover_bg,
        activeforeground=fg,
        relief="flat",
        cursor="hand2",
        padx=UISpacing.MD,
        pady=UISpacing.SM,
        bd=0,
        highlightthickness=0,
    )
    
    if width:
        btn.config(width=width)
    
    def on_enter(e):
        btn.config(bg=hover_bg)
    def on_leave(e):
        btn.config(bg=bg)
    
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    
    return btn


# ============================================================================
# OCR Processing Functions
# ============================================================================

def convert_images_to_pdf(image_files: List[str], output_pdf: str) -> bool:
    """Convert multiple image files to a single PDF."""
    if not PIL_AVAILABLE or not IMG2PDF_AVAILABLE:
        return False
    
    try:
        # Convert images to PDF-compatible format
        converted_images = []
        for img_path in image_files:
            with Image.open(img_path) as img:
                # Convert to RGB if necessary
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                # Save as temporary file if conversion was needed
                if img.format not in ['JPEG', 'PNG']:
                    temp_path = f"{os.path.splitext(img_path)[0]}_converted.jpg"
                    img.save(temp_path, 'JPEG')
                    converted_images.append(temp_path)
                else:
                    converted_images.append(img_path)
        
        # Convert images to PDF
        with open(output_pdf, "wb") as f:
            f.write(img2pdf.convert(converted_images))
        
        # Clean up temporary files
        for img_path in converted_images:
            if img_path.endswith('_converted.jpg'):
                try:
                    os.remove(img_path)
                except:
                    pass
        
        return True
    except Exception as e:
        print(f"Error converting images to PDF: {str(e)}")
        return False


def process_pdf_with_ocr(pdf_path: str, language: str = 'eng', output_callback=None) -> bool:
    """Process a single PDF file with OCR.
    
    Args:
        pdf_path: Path to PDF file
        language: OCR language code
        output_callback: Optional callback function(line) to receive output lines
    
    Returns:
        True if successful, False otherwise
    """
    if not OCRMYPDF_AVAILABLE:
        return False
    
    try:
        # In frozen (exe) mode, avoid spawning sys.executable which relaunches the app.
        if getattr(sys, 'frozen', False):
            try:
                ocrmypdf.ocr(
                    pdf_path,
                    pdf_path,
                    skip_text=True,
                    output_type='pdf',
                    language=language
                )
                return True
            except Exception as e:
                error_msg = str(e)
                if "already contains OCR text" in error_msg.lower() or "already has text" in error_msg.lower():
                    return True
                if output_callback:
                    output_callback(f"Error: {error_msg}")
                print(f"Error processing {pdf_path}: {error_msg}")
                return False

        # Use subprocess to capture ocrmypdf output in real-time
        # This allows us to display progress in the GUI
        cmd = [
            sys.executable, "-m", "ocrmypdf",
            "--skip-text",  # Skip pages that already have text
            "--output-type", "pdf",
            "--language", language,
            pdf_path,
            pdf_path  # Output to same file
        ]
        
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True
        )
        
        # Read output line by line and call callback if provided
        # Handle both \n and \r\n line endings, and progress bars with \r
        last_line = ""
        for line in process.stdout:
            # Handle progress bars that use \r to overwrite the same line
            if '\r' in line:
                # Split by \r and take the last part (most recent progress update)
                parts = line.split('\r')
                line = parts[-1].strip()
                if not line:
                    continue
            
            line = line.rstrip()
            if line and line != last_line:  # Avoid duplicate lines from progress bars
                if output_callback:
                    output_callback(line)
                else:
                    print(line)  # Fallback to console
                last_line = line
        
        process.wait()
        
        if process.returncode == 0:
            return True
        elif process.returncode == 6:  # ocrmypdf exit code 6 = already has text
            return True  # Consider this a success
        else:
            return False
            
    except Exception as e:
        error_msg = str(e)
        # Check if PDF already has OCR text
        if "already contains OCR text" in error_msg.lower() or "already has text" in error_msg.lower():
            return True  # Consider this a success
        
        # Fallback to direct API call if subprocess fails
        try:
            ocrmypdf.ocr(
                pdf_path,
                pdf_path,
                skip_text=True,
                output_type='pdf',
                language=language
            )
            return True
        except Exception as e2:
            if output_callback:
                output_callback(f"Error: {str(e2)}")
            print(f"Error processing {pdf_path}: {error_msg}")
            return False


def process_directory_images(directory: str, language: str = 'eng', output_callback=None) -> bool:
    """Process all images in a directory and convert them to a single PDF with OCR.
    
    Args:
        directory: Directory containing image files
        language: OCR language code
        output_callback: Optional callback function(line) to receive output lines
    
    Returns:
        True if successful, False otherwise
    """
    if not PIL_AVAILABLE or not IMG2PDF_AVAILABLE:
        return False
    
    # Supported image formats
    image_extensions = ('.jpg', '.jpeg', '.png', '.tiff', '.bmp')
    
    # Find all image files in the directory (not recursive)
    image_files = []
    for ext in image_extensions:
        image_files.extend(glob.glob(os.path.join(directory, f'*{ext}')))
        image_files.extend(glob.glob(os.path.join(directory, f'*{ext.upper()}')))
    
    if not image_files:
        return False
    
    # Sort image files and use first image's name for the PDF
    sorted_images = sorted(image_files)
    first_image_name = os.path.splitext(os.path.basename(sorted_images[0]))[0]
    pdf_path = os.path.join(directory, f"{first_image_name}.pdf")
    
    # Convert images to PDF
    if not convert_images_to_pdf(sorted_images, pdf_path):
        return False
    
    # Perform OCR on the combined PDF
    if not process_pdf_with_ocr(pdf_path, language, output_callback=output_callback):
        return False
    
    return True


# ============================================================================
# GUI Application
# ============================================================================

class PDFOCRTool:
    """Main GUI application for PDF OCR."""
    
    def __init__(self):
        # Use TkinterDnD if available
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("PDF OCR Tool")
        self.root.configure(bg=UIColors.BG_SECONDARY)
        
        # Position window
        self.position_window()
        
        # Processing state
        self.processing = False
        self.language = 'eng'
        
        # Setup UI
        self.setup_ui()
        
        # Setup drag and drop
        if HAS_DND:
            self.setup_drag_drop()
    
    def position_window(self):
        """Position window using launcher environment variables."""
        try:
            if 'TOOL_WINDOW_X' in os.environ:
                x = int(os.environ.get('TOOL_WINDOW_X', 100))
                y = int(os.environ.get('TOOL_WINDOW_Y', 100))
                width = int(os.environ.get('TOOL_WINDOW_WIDTH', 900))
                height = int(os.environ.get('TOOL_WINDOW_HEIGHT', 700))
                
                width = max(width, 1280)  # Match launcher minimum width (1280px)
                height = max(height, 600)
                
                self.root.geometry(f"{width}x{height}+{x}+{y}")
                print(f"[INFO] Window positioned at {x},{y} with size {width}x{height}")
            else:
                self.root.geometry("900x700")
                print("[INFO] Running standalone (not from launcher)")
        except (ValueError, TypeError) as e:
            print(f"[WARNING] Could not position window: {e}")
            self.root.geometry("900x700")
        
        # Set minimum size to match launcher
        self.root.minsize(1280, 600)
    
    def setup_ui(self):
        """Setup the main UI."""
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        main_frame = tk.Frame(self.root, bg=UIColors.BG_SECONDARY)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=UISpacing.SM, pady=UISpacing.SM)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Header
        self.create_header(main_frame)
        
        # Drop zone
        self.create_drop_zone(main_frame)
        
        # Options panel
        self.create_options_panel(main_frame)
        
        # Results area
        self.create_results_area(main_frame)
        
        # Buttons
        self.create_button_frame(main_frame)
        
        # Status bar
        self.create_status_bar(main_frame)
        
        # Welcome message
        self.show_welcome_message()
    
    def create_header(self, parent):
        """Create header section."""
        header_frame = tk.Frame(parent, bg=UIColors.BG_PRIMARY, pady=UISpacing.XS)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, UISpacing.XS))
        header_frame.grid_columnconfigure(0, weight=1)
        
        title_label = tk.Label(
            header_frame,
            text="🔍 PDF OCR Tool",
            font=UIFonts.TITLE,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.PRIMARY
        )
        title_label.grid(row=0, column=0, pady=(UISpacing.XS, 0))
        
        desc_label = tk.Label(
            header_frame,
            text="Add OCR text recognition to PDFs and convert images to searchable PDFs",
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_SECONDARY
        )
        desc_label.grid(row=1, column=0, pady=(0, UISpacing.XS))
    
    def create_drop_zone(self, parent):
        """Create drop zone for files."""
        self.drop_frame = tk.Frame(
            parent,
            bg=UIColors.DROP_ZONE_BG,
            bd=2,
            relief="flat",
            padx=UISpacing.MD,
            pady=UISpacing.SM
        )
        self.drop_frame.grid(row=1, column=0, sticky="ew", pady=UISpacing.XS)
        self.drop_frame.config(
            highlightbackground=UIColors.DROP_ZONE_BORDER,
            highlightthickness=2
        )
        
        main_text = "📄 Drag and drop PDF/image files here" if HAS_DND else "📄 Click to select PDF/image files"
        if HAS_DND:
            main_text += " or click to browse"
        
        self.drop_label = tk.Label(
            self.drop_frame,
            text=main_text,
            font=UIFonts.BODY,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_PRIMARY,
            cursor="hand2"
        )
        self.drop_label.pack(pady=UISpacing.XS)
        
        # Click handlers
        for widget in [self.drop_frame, self.drop_label]:
            widget.bind('<Button-1>', lambda e: self.select_files())
        
        self._drop_zone_widgets = [self.drop_frame, self.drop_label]
    
    def create_options_panel(self, parent):
        """Create options panel."""
        options_frame = tk.LabelFrame(
            parent,
            text="  ⚙️ Options  ",
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid"
        )
        options_frame.grid(row=2, column=0, sticky="ew", pady=UISpacing.XS)
        options_frame.grid_columnconfigure(1, weight=1)
        
        # Language selection
        tk.Label(
            options_frame,
            text="OCR Language:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        ).grid(row=0, column=0, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        self.language_var = tk.StringVar(value="eng")
        language_frame = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        language_frame.grid(row=0, column=1, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        languages = [
            ("English", "eng"),
            ("German", "deu"),
            ("English + German", "eng+deu"),
            ("French", "fra"),
            ("Spanish", "spa"),
        ]
        
        for text, value in languages:
            rb = tk.Radiobutton(
                language_frame,
                text=text,
                variable=self.language_var,
                value=value,
                font=UIFonts.BODY,
                bg=UIColors.BG_PRIMARY,
                fg=UIColors.TEXT_PRIMARY
            )
            rb.pack(side=tk.LEFT, padx=(0, UISpacing.MD))
    
    def create_results_area(self, parent):
        """Create results text area."""
        result_frame = tk.LabelFrame(
            parent,
            text="  📋 Processing Log  ",
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid"
        )
        result_frame.grid(row=3, column=0, sticky="nsew", pady=UISpacing.XS)
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)
        
        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            wrap=tk.WORD,
            font=UIFonts.MONO,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            relief="flat",
            padx=UISpacing.SM,
            pady=UISpacing.XS,
            height=10
        )
        self.result_text.grid(row=0, column=0, sticky="nsew", padx=UISpacing.XS, pady=UISpacing.XS)
    
    def create_button_frame(self, parent):
        """Create button frame."""
        button_frame = tk.Frame(parent, bg=UIColors.BG_SECONDARY)
        button_frame.grid(row=4, column=0, pady=UISpacing.XS)
        
        select_files_btn = create_rounded_button(
            button_frame,
            "📂 Select Files",
            self.select_files,
            style="primary",
            width=18
        )
        select_files_btn.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        select_folder_btn = create_rounded_button(
            button_frame,
            "📁 Select Folder",
            self.select_folder,
            style="secondary"
        )
        select_folder_btn.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        clear_btn = create_rounded_button(
            button_frame,
            "Clear Log",
            self.clear_results,
            style="ghost"
        )
        clear_btn.pack(side=tk.LEFT, padx=UISpacing.SM)
    
    def create_status_bar(self, parent):
        """Create status bar."""
        status_frame = tk.Frame(parent, bg=UIColors.BG_TERTIARY, pady=UISpacing.XS)
        status_frame.grid(row=5, column=0, sticky="ew")
        
        # Status indicators
        status_items = [
            ("OCRmyPDF", OCRMYPDF_AVAILABLE),
            ("Pillow", PIL_AVAILABLE),
            ("img2pdf", IMG2PDF_AVAILABLE),
            ("Drag&Drop", HAS_DND)
        ]
        
        status_parts = []
        for name, available in status_items:
            icon = "✓" if available else "✗"
            status_parts.append(f"{icon} {name}")
        
        status_text = "  •  ".join(status_parts)
        
        self.status_label = tk.Label(
            status_frame,
            text=status_text,
            font=UIFonts.SMALL,
            bg=UIColors.BG_TERTIARY,
            fg=UIColors.TEXT_SECONDARY
        )
        self.status_label.pack(fill="x", padx=UISpacing.SM)
    
    def setup_drag_drop(self):
        """Setup drag and drop handlers."""
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        self.drop_frame.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.drop_frame.dnd_bind('<<DragLeave>>', self.on_drag_leave)
    
    def on_drop(self, event):
        """Handle dropped files."""
        files = self.parse_dropped_files(event.data)
        if files:
            self.process_files(files)
    
    def on_drag_enter(self, event):
        """Visual feedback when dragging over."""
        self.drop_frame.config(
            bg=UIColors.DROP_ZONE_ACTIVE,
            highlightbackground=UIColors.DROP_ZONE_BORDER_ACTIVE
        )
        for widget in self._drop_zone_widgets:
            if isinstance(widget, tk.Label):
                widget.config(bg=UIColors.DROP_ZONE_ACTIVE)
    
    def on_drag_leave(self, event):
        """Reset visual feedback."""
        self.drop_frame.config(
            bg=UIColors.DROP_ZONE_BG,
            highlightbackground=UIColors.DROP_ZONE_BORDER
        )
        for widget in self._drop_zone_widgets:
            if isinstance(widget, tk.Label):
                widget.config(bg=UIColors.DROP_ZONE_BG)
    
    def parse_dropped_files(self, data):
        """Parse dropped file paths."""
        files = []
        if '{' in data:
            files = re.findall(r'\{([^}]+)\}', data)
            remaining = re.sub(r'\{[^}]+\}', '', data).strip()
            if remaining:
                files.extend(remaining.split())
        else:
            files = data.split()
        
        # Filter for PDF and image files
        supported_extensions = ('.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.bmp')
        return [f for f in files if f.lower().endswith(supported_extensions)]
    
    def select_files(self):
        """Open file dialog to select PDFs and images."""
        files = filedialog.askopenfilenames(
            title="Select PDF and/or Image Files",
            filetypes=[
                ("All supported files", "*.pdf *.jpg *.jpeg *.png *.tiff *.bmp"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg *.jpeg *.png *.tiff *.bmp"),
                ("All files", "*.*")
            ]
        )
        
        if files:
            self.process_files(list(files))
    
    def select_folder(self):
        """Select folder to process."""
        folder = filedialog.askdirectory(title="Select folder with PDF and/or image files")
        
        if folder:
            # Ask about recursive processing
            recursive = messagebox.askyesno(
                "Recursive Search",
                "Search subfolders recursively?",
                parent=self.root
            )
            
            files = self.find_files_in_folder(folder, recursive)
            
            if files:
                self.process_files(files)
            else:
                messagebox.showwarning("No Files", "No PDF or image files found in the selected folder.")
    
    def find_files_in_folder(self, folder: str, recursive: bool = False) -> List[str]:
        """Find PDF and image files in folder."""
        supported_extensions = ('.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.bmp')
        files = []
        
        if recursive:
            for ext in supported_extensions:
                files.extend(glob.glob(os.path.join(folder, '**', f'*{ext}'), recursive=True))
                files.extend(glob.glob(os.path.join(folder, '**', f'*{ext.upper()}'), recursive=True))
        else:
            for ext in supported_extensions:
                files.extend(glob.glob(os.path.join(folder, f'*{ext}')))
                files.extend(glob.glob(os.path.join(folder, f'*{ext.upper()}')))
        
        return files
    
    def process_files(self, files: List[str]):
        """Process a list of files."""
        if self.processing:
            messagebox.showwarning("Processing", "Already processing files. Please wait.")
            return
        
        if not OCRMYPDF_AVAILABLE:
            messagebox.showerror(
                "Missing Dependency",
                "OCRmyPDF is not available.\n\n"
                "The tool will attempt to install it automatically on next launch.\n\n"
                "Alternatively, install manually with:\n"
                "  pip install ocrmypdf\n\n"
                "Note: OCRmyPDF requires Tesseract OCR to be installed on your system."
            )
            return
        
        # Check for image files
        has_images = any(f.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.bmp')) for f in files)
        if has_images and (not PIL_AVAILABLE or not IMG2PDF_AVAILABLE):
            missing = []
            if not PIL_AVAILABLE:
                missing.append("pillow")
            if not IMG2PDF_AVAILABLE:
                missing.append("img2pdf")
            
            messagebox.showerror(
                "Missing Dependencies",
                f"Image processing requires: {', '.join(missing)}\n\n"
                "The tool will attempt to install them automatically on next launch.\n\n"
                f"Alternatively, install manually with:\n"
                f"  pip install {' '.join(missing)}"
            )
            return
        
        self.processing = True
        self.language = self.language_var.get()
        
        # Set wait cursor
        self.root.config(cursor="wait")
        self.root.update()
        
        try:
            self.result_text.insert(tk.END, f"\n{'='*60}\n")
            self.result_text.insert(tk.END, f"Processing {len(files)} file(s)\n")
            self.result_text.insert(tk.END, f"Language: {self.language}\n")
            self.result_text.insert(tk.END, f"{'='*60}\n\n")
            self.result_text.see(tk.END)
            self.root.update()
            
            # Separate PDFs and images
            pdf_files = [f for f in files if f.lower().endswith('.pdf')]
            image_files = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.bmp'))]
            
            processed_count = 0
            skipped_count = 0
            error_count = 0
            
            # Process PDFs
            for i, pdf_file in enumerate(pdf_files, 1):
                self.result_text.insert(tk.END, f"[{i}/{len(pdf_files)}] Processing PDF: {os.path.basename(pdf_file)}\n")
                self.result_text.see(tk.END)
                self.root.update()
                
                try:
                    # Define callback to display ocrmypdf output in Processing Log
                    def log_output(line):
                        """Callback to display ocrmypdf output in the Processing Log"""
                        self.result_text.insert(tk.END, f"  {line}\n")
                        self.result_text.see(tk.END)
                        self.root.update()
                    
                    if process_pdf_with_ocr(pdf_file, self.language, output_callback=log_output):
                        self.result_text.insert(tk.END, f"  ✓ Successfully processed\n")
                        processed_count += 1
                    else:
                        self.result_text.insert(tk.END, f"  ⏭ Skipped (may already have OCR text)\n")
                        skipped_count += 1
                except Exception as e:
                    self.result_text.insert(tk.END, f"  ✗ Error: {str(e)}\n")
                    error_count += 1
                
                self.result_text.see(tk.END)
                self.root.update()
            
            # Process images (group by directory)
            if image_files:
                # Group images by directory
                image_groups = {}
                for img in image_files:
                    img_dir = os.path.dirname(img)
                    if img_dir not in image_groups:
                        image_groups[img_dir] = []
                    image_groups[img_dir].append(img)
                
                for img_dir, imgs in image_groups.items():
                    self.result_text.insert(tk.END, f"\nProcessing {len(imgs)} image(s) from: {os.path.basename(img_dir)}\n")
                    self.result_text.see(tk.END)
                    self.root.update()
                    
                    try:
                        # Create temporary directory for images
                        temp_dir = os.path.join(img_dir, "temp_ocr")
                        os.makedirs(temp_dir, exist_ok=True)
                        
                        # Copy images to temp directory
                        for img in imgs:
                            shutil.copy2(img, temp_dir)
                        
                        # Define callback to display ocrmypdf output in Processing Log
                        def log_output(line):
                            """Callback to display ocrmypdf output in the Processing Log"""
                            self.result_text.insert(tk.END, f"  {line}\n")
                            self.result_text.see(tk.END)
                            self.root.update()
                        
                        # Process the temporary directory
                        if process_directory_images(temp_dir, self.language, output_callback=log_output):
                            # Find the created PDF
                            pdf_files_in_dir = glob.glob(os.path.join(temp_dir, '*.pdf'))
                            if pdf_files_in_dir:
                                created_pdf = pdf_files_in_dir[0]
                                # Move PDF to original directory
                                final_pdf = os.path.join(img_dir, os.path.basename(created_pdf))
                                if os.path.exists(final_pdf):
                                    os.remove(final_pdf)
                                shutil.move(created_pdf, final_pdf)
                                self.result_text.insert(tk.END, f"  ✓ Created PDF: {os.path.basename(final_pdf)}\n")
                                processed_count += 1
                        
                        # Clean up temp directory
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except Exception as e:
                        self.result_text.insert(tk.END, f"  ✗ Error: {str(e)}\n")
                        error_count += 1
                        # Clean up temp directory on error
                        try:
                            shutil.rmtree(temp_dir, ignore_errors=True)
                        except:
                            pass
                    
                    self.result_text.see(tk.END)
                    self.root.update()
            
            # Summary
            self.result_text.insert(tk.END, f"\n{'='*60}\n")
            self.result_text.insert(tk.END, f"SUMMARY\n")
            self.result_text.insert(tk.END, f"{'='*60}\n")
            self.result_text.insert(tk.END, f"Total: {len(files)} | ")
            self.result_text.insert(tk.END, f"✓ Success: {processed_count} | ")
            self.result_text.insert(tk.END, f"⏭ Skipped: {skipped_count} | ")
            self.result_text.insert(tk.END, f"✗ Failed: {error_count}\n")
            self.result_text.insert(tk.END, f"{'='*60}\n\n")
            self.result_text.see(tk.END)
            
            # Show summary dialog
            if error_count == 0:
                messagebox.showinfo(
                    "Complete",
                    f"Processed {len(files)} file(s).\n\n"
                    f"✓ Success: {processed_count}\n"
                    f"⏭ Skipped: {skipped_count}"
                )
            else:
                messagebox.showwarning(
                    "Complete with Errors",
                    f"Processed {len(files)} file(s).\n\n"
                    f"✓ Success: {processed_count}\n"
                    f"⏭ Skipped: {skipped_count}\n"
                    f"✗ Failed: {error_count}"
                )
        finally:
            self.processing = False
            self.root.config(cursor="")
            self.root.update()
    
    def clear_results(self):
        """Clear results text."""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Log cleared.\n\n")
    
    def show_welcome_message(self):
        """Show welcome message."""
        self.result_text.insert(tk.END, "Welcome to PDF OCR Tool!\n")
        self.result_text.insert(tk.END, "━" * 45 + "\n")
        
        self.result_text.insert(tk.END, "📋 Features: ")
        features = []
        if OCRMYPDF_AVAILABLE:
            features.append("✓ OCR Processing")
        else:
            features.append("✗ OCR Processing")
        if PIL_AVAILABLE and IMG2PDF_AVAILABLE:
            features.append("✓ Image to PDF")
        else:
            features.append("✗ Image to PDF")
        self.result_text.insert(tk.END, " | ".join(features) + "\n")
        
        self.result_text.insert(tk.END, "📄 Supported: PDF, JPG, PNG, TIFF, BMP\n")
        
        if HAS_DND:
            self.result_text.insert(tk.END, "💡 Drag and drop files to begin.\n")
        else:
            self.result_text.insert(tk.END, "💡 Click 'Select Files' to begin.\n")
    
    def run(self):
        """Run the application."""
        self.root.mainloop()


def install_missing_dependencies():
    """Automatically install missing dependencies"""
    if getattr(sys, 'frozen', False):
        print("[INFO] Running as executable - skipping auto-install of dependencies.")
        return False

    missing_deps = []
    
    if not OCRMYPDF_AVAILABLE:
        missing_deps.append("ocrmypdf>=16.0.0")
    if not PIL_AVAILABLE:
        missing_deps.append("pillow>=10.0.0")
    if not IMG2PDF_AVAILABLE:
        missing_deps.append("img2pdf>=0.5.0")
    if not HAS_DND:
        missing_deps.append("tkinterdnd2>=0.3.0")
    
    if missing_deps:
        print(f"[INFO] Missing dependencies detected: {', '.join(missing_deps)}")
        print("[INFO] Attempting to install missing dependencies...")
        
        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet", "--upgrade"] + missing_deps,
                capture_output=True,
                text=True,
                timeout=120
            )
            
            if result.returncode == 0:
                print("[SUCCESS] Missing dependencies installed successfully.")
                print("[INFO] Please restart the application for changes to take effect.")
                return True
            else:
                print(f"[WARNING] Failed to install some dependencies: {result.stderr}")
                print(f"[INFO] Please install manually with: pip install {' '.join(missing_deps)}")
                return False
        except Exception as e:
            print(f"[ERROR] Failed to install dependencies: {e}")
            print(f"[INFO] Please install manually with: pip install {' '.join(missing_deps)}")
            return False
    
    return True


def main():
    """Main entry point."""
    print("[INFO] Starting PDF OCR Tool")
    
    # Try to install missing dependencies automatically
    install_missing_dependencies()
    
    # Re-check availability after installation attempt
    # Note: We can't re-import here, but the user will need to restart
    print(f"[INFO] OCRmyPDF: {'Available' if OCRMYPDF_AVAILABLE else 'Not available'}")
    print(f"[INFO] Pillow: {'Available' if PIL_AVAILABLE else 'Not available'}")
    print(f"[INFO] img2pdf: {'Available' if IMG2PDF_AVAILABLE else 'Not available'}")
    print(f"[INFO] Drag&Drop: {'Available' if HAS_DND else 'Not available'}")
    
    app = PDFOCRTool()
    app.run()


if __name__ == "__main__":
    main()
