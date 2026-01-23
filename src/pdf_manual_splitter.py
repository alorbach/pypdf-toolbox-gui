"""
PDF Manual Splitter Tool

Manually split PDF files into multiple documents by interactively
selecting split points with visual thumbnail preview.

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

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolledtext
import logging
import argparse
import os
from pathlib import Path
from datetime import datetime
import json
import re

# PDF manipulation support
try:
    from PyPDF2 import PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# PIL for thumbnail display
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Drag and drop support (optional)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# PDF to image conversion (for on-the-fly thumbnails)
# Try PyMuPDF first (no external dependencies), then fall back to pdf2image
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

# Combined flag for any thumbnail generation capability
THUMBNAIL_GENERATION_AVAILABLE = PYMUPDF_AVAILABLE or PDF2IMAGE_AVAILABLE

# Parse command line arguments
parser = argparse.ArgumentParser(description='PDF Manual Splitter - Split PDF files manually into multiple documents')
parser.add_argument('pdf_file', nargs='?', help='Path to PDF file')
parser.add_argument('--output-folder', '-o', help='Output folder for split PDFs')
parser.add_argument('--debug', action='store_true', help='Enable debug output')
args = parser.parse_args()

# Configure logging
log_level = logging.DEBUG if args.debug else logging.INFO
logging.basicConfig(
    level=log_level,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Log startup information
print(f"[INFO] Starting PDF Manual Splitter")
print(f"[INFO] PyPDF2: {'Available' if PDF_AVAILABLE else 'Not available'}")
print(f"[INFO] PIL/Pillow: {'Available' if PIL_AVAILABLE else 'Not available'}")
print(f"[INFO] Drag&Drop: {'Available' if DND_AVAILABLE else 'Not available'}")
print(f"[INFO] PyMuPDF: {'Available' if PYMUPDF_AVAILABLE else 'Not available'}")
print(f"[INFO] pdf2image: {'Available' if PDF2IMAGE_AVAILABLE else 'Not available'}")
if THUMBNAIL_GENERATION_AVAILABLE:
    print(f"[INFO] Thumbnail generation: Available via {'PyMuPDF' if PYMUPDF_AVAILABLE else 'pdf2image'}")
else:
    print(f"[INFO] Thumbnail generation: Not available (install PyMuPDF: pip install pymupdf)")
logger.debug(f"Starting PDF Manual Splitter")
logger.debug(f"PDF_AVAILABLE: {PDF_AVAILABLE}")
logger.debug(f"PIL_AVAILABLE: {PIL_AVAILABLE}")
logger.debug(f"DND_AVAILABLE: {DND_AVAILABLE}")
logger.debug(f"PYMUPDF_AVAILABLE: {PYMUPDF_AVAILABLE}")
logger.debug(f"PDF2IMAGE_AVAILABLE: {PDF2IMAGE_AVAILABLE}")

# Default output folder
DEFAULT_OUTPUT_FOLDER = "_splitted"

# ============================================================================
# Modern UI Styling Constants
# ============================================================================

class UIColors:
    """Modern color palette for consistent UI styling."""
    # Primary colors
    PRIMARY = "#2563eb"          # Blue 600
    PRIMARY_HOVER = "#1d4ed8"    # Blue 700
    PRIMARY_LIGHT = "#dbeafe"    # Blue 100
    
    # Secondary/Accent colors
    SECONDARY = "#64748b"        # Slate 500
    SECONDARY_HOVER = "#475569"  # Slate 600
    
    # Success/Error/Warning
    SUCCESS = "#16a34a"          # Green 600
    SUCCESS_LIGHT = "#dcfce7"    # Green 100
    SUCCESS_HOVER = "#15803d"    # Green 700
    ERROR = "#dc2626"            # Red 600
    ERROR_LIGHT = "#fee2e2"      # Red 100
    ERROR_HOVER = "#b91c1c"      # Red 700
    WARNING = "#f59e0b"          # Amber 500
    WARNING_LIGHT = "#fef3c7"    # Amber 100
    
    # Neutral colors
    BG_PRIMARY = "#ffffff"       # White
    BG_SECONDARY = "#f8fafc"     # Slate 50
    BG_TERTIARY = "#f1f5f9"      # Slate 100
    BORDER = "#e2e8f0"           # Slate 200
    BORDER_DARK = "#cbd5e1"      # Slate 300
    TEXT_PRIMARY = "#1e293b"     # Slate 800
    TEXT_SECONDARY = "#64748b"   # Slate 500
    TEXT_MUTED = "#94a3b8"       # Slate 400
    
    # Special
    SPLIT_ACTIVE = "#ef4444"     # Red 500 - for active split points
    SPLIT_HOVER = "#fca5a5"      # Red 300
    THUMBNAIL_BG = "#ffffff"
    THUMBNAIL_BORDER = "#e2e8f0"
    THUMBNAIL_HOVER = "#dbeafe"
    
    # Drop zone
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
    """Create a styled button with consistent appearance.
    
    Styles: primary, secondary, success, danger, ghost
    """
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
    
    # Hover effects
    def on_enter(e):
        btn.config(bg=hover_bg)
    
    def on_leave(e):
        btn.config(bg=bg)
    
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    
    return btn


def create_card_frame(parent, title=None, padding=True):
    """Create a card-like frame with optional title."""
    if title:
        frame = tk.LabelFrame(
            parent,
            text=title,
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid",
            padx=UISpacing.MD if padding else 0,
            pady=UISpacing.SM if padding else 0,
        )
    else:
        frame = tk.Frame(
            parent,
            bg=UIColors.BG_PRIMARY,
            bd=1,
            relief="solid",
            padx=UISpacing.MD if padding else 0,
            pady=UISpacing.SM if padding else 0,
        )
    return frame


def auto_scroll_text_widget(text_widget):
    """Auto-scroll text widget to bottom."""
    text_widget.see(tk.END)
    text_widget.update()


def check_thumbnails_folder(pdf_path):
    """Check if .thumbs folder exists for the given PDF."""
    pdf_dir = os.path.dirname(pdf_path)
    thumbs_folder = os.path.join(pdf_dir, ".thumbs")
    return os.path.exists(thumbs_folder) and os.path.isdir(thumbs_folder)


def generate_pdf_thumbnails(pdf_path, max_size=(150, 200)):
    """Generate thumbnails for all pages of a PDF.
    
    Tries PyMuPDF first (no external dependencies), then falls back to pdf2image.
    
    Returns a dict mapping page numbers (1-based) to PIL Image objects.
    Returns empty dict if no thumbnail generation is available or generation fails.
    """
    if not PIL_AVAILABLE:
        return {}
    
    # Try PyMuPDF first (no external dependencies like poppler)
    if PYMUPDF_AVAILABLE:
        try:
            print(f"[INFO] Generating thumbnails using PyMuPDF for {os.path.basename(pdf_path)}...")
            
            doc = fitz.open(pdf_path)
            thumbnails = {}
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                # Render at a scale that gives us roughly thumbnail size
                # Default page size is 612x792 points (letter), scale to ~150 width
                zoom = max_size[0] / page.rect.width
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                
                # Convert to PIL Image
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                # Ensure it fits within max_size
                img.thumbnail(max_size, Image.Resampling.LANCZOS)
                thumbnails[page_num + 1] = img  # 1-based page numbers
            
            doc.close()
            print(f"[INFO] Generated {len(thumbnails)} thumbnails via PyMuPDF")
            return thumbnails
            
        except Exception as e:
            print(f"[WARNING] PyMuPDF thumbnail generation failed: {e}")
            logger.warning(f"PyMuPDF failed for {pdf_path}: {e}")
            # Fall through to try pdf2image
    
    # Fall back to pdf2image (requires poppler)
    if PDF2IMAGE_AVAILABLE:
        try:
            print(f"[INFO] Generating thumbnails using pdf2image for {os.path.basename(pdf_path)}...")
            
            # Convert PDF to images at low DPI for thumbnails
            images = convert_from_path(
                pdf_path,
                dpi=72,  # Low DPI for faster thumbnail generation
                fmt='png',
                thread_count=2
            )
            
            thumbnails = {}
            for i, img in enumerate(images, start=1):
                # Resize to thumbnail size
                img.thumbnail(max_size, Image.Resampling.LANCZOS)
                thumbnails[i] = img
                
            print(f"[INFO] Generated {len(thumbnails)} thumbnails via pdf2image")
            return thumbnails
            
        except Exception as e:
            print(f"[WARNING] pdf2image thumbnail generation failed: {e}")
            logger.warning(f"pdf2image failed for {pdf_path}: {e}")
    
    return {}


def get_thumbnail_path(pdf_path, page_number):
    """Get the path to the thumbnail for a specific page."""
    pdf_dir = os.path.dirname(pdf_path)
    pdf_basename = os.path.splitext(os.path.basename(pdf_path))[0]
    thumbs_folder = os.path.join(pdf_dir, ".thumbs")
    
    # Thumbnail naming convention: {basename}_t{page_num}.webp
    thumbnail_filename = f"{pdf_basename}_t{page_number}.webp"
    thumbnail_path = os.path.join(thumbs_folder, thumbnail_filename)
    
    if os.path.exists(thumbnail_path):
        return thumbnail_path
    return None


def load_thumbnail_image(thumbnail_path, max_size=(150, 200), master=None):
    """Load and resize thumbnail image for display."""
    if not PIL_AVAILABLE:
        return None
    
    try:
        if not os.path.exists(thumbnail_path):
            return None
        
        image = Image.open(thumbnail_path)
        image.thumbnail(max_size, Image.Resampling.LANCZOS)
        
        if master:
            photo = ImageTk.PhotoImage(image, master=master)
        else:
            photo = ImageTk.PhotoImage(image)
        
        return photo
        
    except Exception as e:
        logger.error(f"Failed to load thumbnail {thumbnail_path}: {str(e)}")
        return None


def get_pdf_page_count(pdf_path):
    """Get the total number of pages in a PDF file."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            return len(reader.pages)
    except Exception as e:
        logger.error(f"Error getting page count for {pdf_path}: {str(e)}")
        return 0


def render_pdf_page_to_image(pdf_path, page_number, max_size=None):
    """Render a PDF page to a PIL Image using PyMuPDF.
    
    Args:
        pdf_path: Path to the PDF file
        page_number: 1-based page number
        max_size: Optional tuple (width, height) to limit size
    
    Returns:
        PIL Image or None if rendering fails
    """
    if not PYMUPDF_AVAILABLE or not PIL_AVAILABLE:
        return None
    
    try:
        doc = fitz.open(pdf_path)
        if page_number < 1 or page_number > len(doc):
            doc.close()
            return None
        
        page = doc[page_number - 1]  # Convert to 0-based index
        
        # Render at good quality (150 DPI equivalent)
        zoom = 2.0  # 2x zoom for better quality
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
        
        # Resize if max_size specified
        if max_size:
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
        
        return img
        
    except Exception as e:
        logger.error(f"Failed to render PDF page {page_number}: {e}")
        return None


def show_full_size_image(thumbnail_path, pdf_path=None, page_num=None, split_callback=None, parent=None):
    """Show full size image in a popup window with navigation and split functionality."""
    if not PIL_AVAILABLE:
        return
    
    try:
        # Get all available pages for navigation if pdf_path is provided
        available_pages = []
        current_page_index = 0
        
        if pdf_path and page_num:
            total_pages = get_pdf_page_count(pdf_path)
            available_pages = list(range(1, total_pages + 1))
            try:
                current_page_index = available_pages.index(page_num)
            except ValueError:
                current_page_index = 0
        
        # Create popup window
        popup = tk.Toplevel(parent)
        title_file = os.path.basename(pdf_path) if pdf_path else (os.path.basename(thumbnail_path) if thumbnail_path else "PDF")
        popup.title(f"Full View: {title_file}")
        
        # Make it modal and on top of parent
        if parent:
            popup.transient(parent)
        popup.grab_set()
        popup.lift()
        popup.focus_force()
        
        popup.grid_rowconfigure(1, weight=1)
        popup.grid_columnconfigure(0, weight=1)
        
        current_page = tk.IntVar(value=page_num if page_num else 1)
        
        image_frame = tk.Frame(popup, bg=UIColors.BG_SECONDARY)
        image_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        
        current_image_label = tk.Label(image_frame, bg=UIColors.BG_SECONDARY)
        current_image_label.pack(expand=True, fill="both")
        
        def load_and_display_image(page_number):
            """Load and display image for the given page number."""
            try:
                image = None
                
                # First try to render directly from PDF (best quality)
                if pdf_path and PYMUPDF_AVAILABLE:
                    screen_width = popup.winfo_screenwidth() - 200
                    screen_height = popup.winfo_screenheight() - 200
                    image = render_pdf_page_to_image(pdf_path, page_number, (screen_width, screen_height))
                
                # Fall back to thumbnail file if direct rendering failed
                if image is None and thumbnail_path:
                    thumb_path = get_thumbnail_path(pdf_path, page_number) if pdf_path else thumbnail_path
                    if thumb_path and os.path.exists(thumb_path):
                        image = Image.open(thumb_path)
                        screen_width = popup.winfo_screenwidth() - 200
                        screen_height = popup.winfo_screenheight() - 200
                        if image.width > screen_width or image.height > screen_height:
                            image.thumbnail((screen_width, screen_height), Image.Resampling.LANCZOS)
                
                if image is None:
                    logger.warning(f"Could not load image for page {page_number}")
                    return False
                
                photo = ImageTk.PhotoImage(image)
                
                current_image_label.config(image=photo)
                current_image_label.image = photo
                
                popup.title(f"Page {page_number} - {os.path.basename(pdf_path) if pdf_path else 'PDF'}")
                current_page.set(page_number)
                
                return True
                
            except Exception as e:
                logger.warning(f"Failed to load image for page {page_number}: {str(e)}")
                return False
        
        prev_button = None
        next_button = None
        split_button = None
        page_info_label = None
        
        def go_previous():
            nonlocal current_page_index
            if available_pages and current_page_index > 0:
                new_index = current_page_index - 1
                new_page = available_pages[new_index]
                if load_and_display_image(new_page):
                    current_page_index = new_index
                    update_button_states()
        
        def go_next():
            nonlocal current_page_index
            if available_pages and current_page_index < len(available_pages) - 1:
                new_index = current_page_index + 1
                new_page = available_pages[new_index]
                if load_and_display_image(new_page):
                    current_page_index = new_index
                    update_button_states()
        
        def split_after_current_page():
            if split_callback and current_page.get() < len(available_pages):
                try:
                    split_callback(current_page.get())
                    confirmation_label = tk.Label(nav_frame, text="✓ Split added!", 
                                                 fg="green", font=("Arial", 9, "bold"))
                    confirmation_label.grid(row=1, column=0, columnspan=5, pady=2)
                    popup.after(2000, lambda: confirmation_label.destroy())
                except Exception as e:
                    logger.error(f"Error adding split from fullscreen: {e}")
        
        def update_button_states():
            if available_pages:
                if prev_button:
                    prev_button.config(state="normal" if current_page_index > 0 else "disabled")
                if next_button:
                    next_button.config(state="normal" if current_page_index < len(available_pages) - 1 else "disabled")
                if page_info_label:
                    page_info_label.config(text=f"Page {current_page.get()} of {len(available_pages)}")
                if split_callback and split_button:
                    can_split = current_page.get() < len(available_pages)
                    split_button.config(state="normal" if can_split else "disabled")
        
        if (available_pages and len(available_pages) > 1) or split_callback:
            nav_frame = tk.Frame(popup)
            nav_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
            nav_frame.grid_columnconfigure(3, weight=1)
            
            col_index = 0
            
            if available_pages and len(available_pages) > 1:
                prev_button = tk.Button(nav_frame, text="◀ Previous", command=go_previous, 
                                      font=("Arial", 9), width=8)
                prev_button.grid(row=0, column=col_index, padx=2)
                col_index += 1
            
            if split_callback and available_pages:
                split_button = tk.Button(nav_frame, text="🔪 Add Split", 
                                       command=split_after_current_page,
                                       bg="orange", font=("Arial", 9, "bold"), width=12)
                split_button.grid(row=0, column=col_index, padx=2)
                col_index += 1
            
            page_info_label = tk.Label(nav_frame, 
                                     text=f"Page {current_page.get()} of {len(available_pages)}" if available_pages else f"Page {current_page.get()}", 
                                     font=("Arial", 10, "bold"))
            page_info_label.grid(row=0, column=col_index, padx=10)
            col_index += 1
            
            if available_pages and len(available_pages) > 1:
                next_button = tk.Button(nav_frame, text="Next ▶", command=go_next, 
                                      font=("Arial", 9), width=8)
                next_button.grid(row=0, column=col_index, padx=2)
        
        button_frame = tk.Frame(popup)
        button_frame.grid(row=2, column=0, pady=10)
        
        help_texts = []
        if available_pages and len(available_pages) > 1:
            help_texts.append("← → : Navigate")
        if split_callback:
            help_texts.append("S/Space : Add Split")
        help_texts.append("Esc : Close")
        
        if help_texts:
            help_text = " • ".join(help_texts)
            help_label = tk.Label(button_frame, text=f"⌨️ {help_text}", 
                                font=("Arial", 9), fg="gray")
            help_label.pack(pady=(0, 5))
        
        close_button = tk.Button(button_frame, text="Close", command=popup.destroy, 
                               font=("Arial", 10), padx=20)
        close_button.pack()
        
        initial_page = page_num if page_num else 1
        if not load_and_display_image(initial_page):
            try:
                image = Image.open(thumbnail_path)
                screen_width = popup.winfo_screenwidth() - 200
                screen_height = popup.winfo_screenheight() - 200
                
                if image.width > screen_width or image.height > screen_height:
                    image.thumbnail((screen_width, screen_height), Image.Resampling.LANCZOS)
                
                photo = ImageTk.PhotoImage(image)
                current_image_label.config(image=photo)
                current_image_label.image = photo
            except Exception as e:
                current_image_label.config(text=f"Error loading image:\n{str(e)}", 
                                         justify="center", wraplength=400)
        
        if (available_pages and len(available_pages) > 1) or split_callback:
            update_button_states()
        
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (popup.winfo_width() // 2)
        y = (popup.winfo_screenheight() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{x}+{y}")
        
        popup.bind('<Escape>', lambda e: popup.destroy())
        if available_pages and len(available_pages) > 1:
            popup.bind('<Left>', lambda e: go_previous())
            popup.bind('<Right>', lambda e: go_next())
        
        if split_callback:
            popup.bind('<KeyPress-s>', lambda e: split_after_current_page())
            popup.bind('<KeyPress-S>', lambda e: split_after_current_page())
            popup.bind('<space>', lambda e: split_after_current_page())
        
        popup.focus_set()
        
    except Exception as e:
        logger.error(f"Failed to show full size image {thumbnail_path}: {str(e)}")
        messagebox.showerror("Error", f"Could not display image: {str(e)}")


def split_pdf_file(pdf_path, splits, output_folder, move_original=False, result_text=None):
    """Split a PDF file based on user-defined splits."""
    
    def safe_write_result(text):
        try:
            if result_text and result_text.winfo_exists():
                result_text.insert(tk.END, text)
                auto_scroll_text_widget(result_text)
        except tk.TclError:
            pass
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            total_pages = len(reader.pages)

            # Determine output directory
            if output_folder and os.path.isabs(output_folder):
                output_dir = Path(output_folder)
            else:
                pdf_dir = os.path.dirname(pdf_path)
                folder_name = output_folder or DEFAULT_OUTPUT_FOLDER
                output_dir = Path(pdf_dir) / folder_name
            
            output_dir.mkdir(parents=True, exist_ok=True)
            
            safe_write_result(f"  → Output directory: {output_dir}\n")
            safe_write_result(f"  → Creating {len(splits)} document(s)\n")

            documents_meta = []

            for idx, (start, end, name) in enumerate(splits, 1):
                writer = PdfWriter()
                for page_num in range(start - 1, min(end, total_pages)):
                    if 0 <= page_num < total_pages:
                        writer.add_page(reader.pages[page_num])

                if not name:
                    name = f"Document_{idx}.pdf"
                if not name.lower().endswith('.pdf'):
                    name += '.pdf'

                out_path = output_dir / name
                with open(out_path, 'wb') as out_file:
                    writer.write(out_file)

                logger.info(f"Created: {out_path} (Pages {start}-{end})")
                safe_write_result(f"  ✓ Created: {name} (Pages {start}-{end})\n")

                documents_meta.append({
                    'document_number': idx,
                    'start_page': start,
                    'end_page': end,
                    'filename': name
                })

            # Write summary JSON
            summary = {
                'original_file': os.path.basename(pdf_path),
                'split_timestamp': datetime.now().isoformat(),
                'total_pages': total_pages,
                'documents': documents_meta,
                'output_folder': str(output_dir)
            }
            summary_path = output_dir / f"{Path(pdf_path).stem}_split_summary.json"
            with open(summary_path, 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)

            safe_write_result(f"  ✓ Summary saved: {Path(pdf_path).stem}_split_summary.json\n")

        return True
    except Exception as e:
        logger.error(f"Error splitting PDF {pdf_path}: {str(e)}")
        safe_write_result(f"  ✗ Error splitting: {str(e)}\n")
        return False


class ManualSplitDialog:
    """Manual split dialog with thumbnail preview."""
    
    def __init__(self, pdf_path, output_folder, parent=None):
        self.pdf_path = pdf_path
        self.output_folder = output_folder or DEFAULT_OUTPUT_FOLDER
        self.total_pages = get_pdf_page_count(pdf_path)
        self.parent = parent
        
        # Check for pre-existing thumbnails in .thumbs folder
        thumbnails_available = check_thumbnails_folder(pdf_path)
        self.has_thumbs_folder = thumbnails_available and PIL_AVAILABLE
        
        # Try to generate thumbnails on-the-fly if PyMuPDF or pdf2image is available
        self.generated_thumbnails = {}
        if not self.has_thumbs_folder and THUMBNAIL_GENERATION_AVAILABLE and PIL_AVAILABLE:
            self.generated_thumbnails = generate_pdf_thumbnails(pdf_path)
        
        # We have thumbnails if either source is available
        self.has_thumbnails = self.has_thumbs_folder or bool(self.generated_thumbnails)
        
        self.split_points = [1]
        self.split_names = {}
        self.result = None
        
        self.create_dialog()
    
    def create_dialog(self):
        """Create the main dialog window."""
        if self.total_pages < 1:
            messagebox.showerror("Error", "PDF could not be read or has no pages.")
            return
        
        # Create as Toplevel if we have a parent, otherwise as Tk root
        if self.parent:
            self.root = tk.Toplevel(self.parent)
            self.root.transient(self.parent)  # Keep on top of parent
            self.root.grab_set()  # Make modal
            self._is_toplevel = True
        else:
            self.root = tk.Tk()
            self._is_toplevel = False
        
        self.root.title("PDF Manual Splitter")
        self.root.geometry("1400x900")
        self.root.minsize(1280, 700)  # Match launcher minimum width (1280px)
        self.root.resizable(True, True)
        self.root.configure(bg=UIColors.BG_SECONDARY)
        
        # Ensure it's on top and focused
        self.root.lift()
        self.root.focus_force()
        
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Header section
        header_frame = tk.Frame(self.root, bg=UIColors.BG_PRIMARY, pady=UISpacing.MD)
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)
        
        # Title with icon
        title_label = tk.Label(
            header_frame, 
            text="✂️  PDF Manual Splitter", 
            font=UIFonts.TITLE,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.PRIMARY
        )
        title_label.grid(row=0, column=0, pady=(UISpacing.SM, 0))
        
        # File info with styled badge
        file_info_frame = tk.Frame(header_frame, bg=UIColors.BG_PRIMARY)
        file_info_frame.grid(row=1, column=0, pady=UISpacing.SM)
        
        file_icon = tk.Label(
            file_info_frame, 
            text="📄", 
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY
        )
        file_icon.pack(side=tk.LEFT, padx=(0, UISpacing.XS))
        
        file_name_label = tk.Label(
            file_info_frame, 
            text=os.path.basename(self.pdf_path),
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        file_name_label.pack(side=tk.LEFT)
        
        # Page count badge
        page_badge = tk.Label(
            file_info_frame,
            text=f" {self.total_pages} pages ",
            font=UIFonts.SMALL_BOLD,
            bg=UIColors.PRIMARY_LIGHT,
            fg=UIColors.PRIMARY,
            padx=UISpacing.SM,
            pady=UISpacing.XS
        )
        page_badge.pack(side=tk.LEFT, padx=(UISpacing.SM, 0))
        
        # Main content frame
        main_frame = tk.Frame(self.root, bg=UIColors.BG_SECONDARY)
        main_frame.grid(row=2, column=0, sticky="nsew", padx=UISpacing.MD, pady=UISpacing.SM)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=2)
        main_frame.grid_columnconfigure(1, weight=1)
        
        self.create_thumbnail_area(main_frame)
        self.create_control_panel(main_frame)
        self.create_config_frame()
        self.create_button_frame()
        
        self.update_split_display()
        
        self.root.bind('<Return>', lambda e: self.do_split())
        self.root.bind('<Escape>', lambda e: self.cancel_dialog())
        
        self.center_window()
    
    def create_thumbnail_area(self, parent):
        """Create the scrollable thumbnail area."""
        thumb_frame = tk.LabelFrame(
            parent, 
            text="  📖 Page Preview  ", 
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid"
        )
        thumb_frame.grid(row=0, column=0, sticky="nsew", padx=(0, UISpacing.SM))
        thumb_frame.grid_rowconfigure(0, weight=1)
        thumb_frame.grid_columnconfigure(0, weight=1)
        
        self.thumb_canvas = tk.Canvas(
            thumb_frame, 
            bg=UIColors.BG_SECONDARY,
            highlightthickness=0
        )
        thumb_scrollbar = ttk.Scrollbar(thumb_frame, orient="vertical", command=self.thumb_canvas.yview)
        self.scrollable_thumb_frame = tk.Frame(self.thumb_canvas, bg=UIColors.BG_SECONDARY)
        
        self.thumb_canvas.configure(yscrollcommand=thumb_scrollbar.set)
        self.thumb_canvas.grid(row=0, column=0, sticky="nsew", padx=UISpacing.XS, pady=UISpacing.XS)
        thumb_scrollbar.grid(row=0, column=1, sticky="ns")
        
        canvas_frame_id = self.thumb_canvas.create_window((0, 0), window=self.scrollable_thumb_frame, anchor="nw")
        
        # Bind resize event to reflow thumbnails when window is resized
        self.thumb_canvas.bind('<Configure>', self._on_canvas_resize)
        
        # Info bar at bottom
        info_frame = tk.Frame(thumb_frame, bg=UIColors.BG_TERTIARY, pady=UISpacing.SM)
        info_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        
        if self.has_thumbnails:
            info_icon = "💡"
            info_bg = UIColors.PRIMARY_LIGHT
            info_fg = UIColors.PRIMARY
            if self.generated_thumbnails:
                info_text = "Click thumbnails for full view  •  Use split buttons below each page"
            else:
                info_text = "Click thumbnails for full view  •  Use split buttons below each page"
        else:
            info_icon = "⚠️"
            info_bg = UIColors.WARNING_LIGHT
            info_fg = UIColors.TEXT_PRIMARY
            if not PIL_AVAILABLE:
                info_text = "Install Pillow for preview: pip install Pillow"
            elif not THUMBNAIL_GENERATION_AVAILABLE:
                info_text = "Install PyMuPDF for thumbnails: pip install pymupdf"
            else:
                info_text = "Could not generate thumbnails. Check console for errors."
        
        info_label = tk.Label(
            info_frame, 
            text=f"  {info_icon}  {info_text}  ", 
            font=UIFonts.SMALL,
            bg=info_bg,
            fg=info_fg,
            padx=UISpacing.SM,
            pady=UISpacing.XS
        )
        info_label.pack(pady=UISpacing.XS)
        
        self.create_thumbnails()
        
        def configure_scroll_region(event=None):
            self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))
            canvas_width = self.thumb_canvas.winfo_width()
            if canvas_width > 1:
                self.thumb_canvas.itemconfig(canvas_frame_id, width=canvas_width)
        
        self.scrollable_thumb_frame.bind("<Configure>", configure_scroll_region)
        self.thumb_canvas.bind("<Configure>", configure_scroll_region)
        
        if self.has_thumbnails:
            self.root.after(100, self.load_thumbnail_images_delayed)
        
        def on_mousewheel(event):
            self.thumb_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        self.thumb_canvas.bind("<MouseWheel>", on_mousewheel)
        self.scrollable_thumb_frame.bind("<MouseWheel>", on_mousewheel)
    
    def create_control_panel(self, parent):
        """Create the control panel for split management."""
        control_frame = tk.LabelFrame(
            parent, 
            text="  ⚙️ Split Controls  ", 
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid"
        )
        control_frame.grid(row=0, column=1, sticky="nsew", padx=(UISpacing.SM, 0))
        control_frame.grid_rowconfigure(2, weight=1)
        control_frame.grid_columnconfigure(0, weight=1)
        
        # Instructions card
        instructions_frame = tk.Frame(control_frame, bg=UIColors.BG_TERTIARY, padx=UISpacing.MD, pady=UISpacing.SM)
        instructions_frame.grid(row=0, column=0, sticky="ew", padx=UISpacing.SM, pady=UISpacing.SM)
        
        instructions_title = tk.Label(
            instructions_frame, 
            text="How to use:",
            font=UIFonts.SMALL_BOLD,
            bg=UIColors.BG_TERTIARY,
            fg=UIColors.TEXT_PRIMARY,
            anchor="w"
        )
        instructions_title.pack(anchor="w")
        
        instructions_items = [
            "• Click 'Split after Page X' to add split points",
            "• Split points create separate documents",
            "• Edit filenames before splitting"
        ]
        
        for item in instructions_items:
            tk.Label(
                instructions_frame, 
                text=item,
                font=UIFonts.SMALL,
                bg=UIColors.BG_TERTIARY,
                fg=UIColors.TEXT_SECONDARY,
                anchor="w"
            ).pack(anchor="w", pady=(UISpacing.XS, 0))
        
        # Split summary header
        summary_header = tk.Frame(control_frame, bg=UIColors.BG_PRIMARY)
        summary_header.grid(row=1, column=0, sticky="ew", padx=UISpacing.SM, pady=(UISpacing.SM, 0))
        
        self.split_count_label = tk.Label(
            summary_header,
            text="📋 Document Preview",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        self.split_count_label.pack(side=tk.LEFT)
        
        # Split display area
        split_display_frame = tk.Frame(
            control_frame, 
            bg=UIColors.BG_PRIMARY,
            bd=1,
            relief="solid"
        )
        split_display_frame.grid(row=2, column=0, sticky="nsew", padx=UISpacing.SM, pady=UISpacing.SM)
        split_display_frame.grid_rowconfigure(0, weight=1)
        split_display_frame.grid_columnconfigure(0, weight=1)
        
        self.split_display_text = scrolledtext.ScrolledText(
            split_display_frame, 
            width=35, 
            height=12, 
            font=UIFonts.SMALL,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            relief="flat",
            wrap=tk.WORD,
            padx=UISpacing.SM,
            pady=UISpacing.SM
        )
        self.split_display_text.grid(row=0, column=0, sticky="nsew")
        
        # Control buttons
        control_buttons_frame = tk.Frame(control_frame, bg=UIColors.BG_PRIMARY, pady=UISpacing.SM)
        control_buttons_frame.grid(row=3, column=0, sticky="ew", padx=UISpacing.SM)
        control_buttons_frame.grid_columnconfigure(0, weight=1)
        control_buttons_frame.grid_columnconfigure(1, weight=1)
        
        reset_button = create_rounded_button(
            control_buttons_frame, 
            "🔄 Reset Splits", 
            self.clear_all_splits,
            style="secondary"
        )
        reset_button.grid(row=0, column=0, padx=UISpacing.XS, sticky="ew")
        
        edit_button = create_rounded_button(
            control_buttons_frame, 
            "✏️ Edit Names", 
            self.edit_names,
            style="ghost"
        )
        edit_button.grid(row=0, column=1, padx=UISpacing.XS, sticky="ew")
    
    def create_config_frame(self):
        """Create configuration frame for output settings."""
        config_frame = tk.Frame(self.root, bg=UIColors.BG_PRIMARY, pady=UISpacing.SM)
        config_frame.grid(row=3, column=0, sticky="ew", padx=UISpacing.MD, pady=(0, UISpacing.SM))
        config_frame.grid_columnconfigure(1, weight=1)
        
        # Output folder row
        folder_label = tk.Label(
            config_frame, 
            text="📁 Output Folder:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        folder_label.grid(row=0, column=0, padx=(UISpacing.SM, UISpacing.MD), pady=UISpacing.SM, sticky="w")
        
        self.output_folder_var = tk.StringVar(value=self.output_folder)
        output_entry = tk.Entry(
            config_frame, 
            textvariable=self.output_folder_var, 
            font=UIFonts.BODY,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            relief="solid",
            bd=1
        )
        output_entry.grid(row=0, column=1, sticky="ew", padx=UISpacing.SM, pady=UISpacing.SM, ipady=UISpacing.XS)
        
        def browse_output_folder():
            # Use PDF file directory as initial directory
            initial_dir = None
            if self.pdf_path:
                initial_dir = os.path.dirname(os.path.abspath(self.pdf_path))
            elif os.path.exists(os.getcwd()):
                initial_dir = os.getcwd()
            
            folder = filedialog.askdirectory(
                title="Select output folder for split PDFs",
                initialdir=initial_dir
            )
            if folder:
                self.output_folder_var.set(folder)
        
        browse_button = create_rounded_button(
            config_frame, 
            "Browse...", 
            browse_output_folder,
            style="secondary"
        )
        browse_button.grid(row=0, column=2, padx=UISpacing.SM, pady=UISpacing.SM)
    
    def create_button_frame(self):
        """Create the button frame."""
        button_frame = tk.Frame(self.root, bg=UIColors.BG_SECONDARY, pady=UISpacing.MD)
        button_frame.grid(row=4, column=0, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        
        # Center the buttons
        buttons_container = tk.Frame(button_frame, bg=UIColors.BG_SECONDARY)
        buttons_container.grid(row=0, column=0)
        
        # Split button (primary action)
        split_button = tk.Button(
            buttons_container, 
            text="✂️  Split PDF Now", 
            command=self.do_split,
            font=UIFonts.BUTTON,
            bg=UIColors.SUCCESS,
            fg="white",
            activebackground=UIColors.SUCCESS_HOVER,
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            padx=UISpacing.XL,
            pady=UISpacing.SM,
            bd=0,
            width=18
        )
        split_button.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        # Hover effect for split button
        def on_split_enter(e):
            split_button.config(bg=UIColors.SUCCESS_HOVER)
        def on_split_leave(e):
            split_button.config(bg=UIColors.SUCCESS)
        split_button.bind("<Enter>", on_split_enter)
        split_button.bind("<Leave>", on_split_leave)
        
        # Cancel button
        cancel_button = tk.Button(
            buttons_container, 
            text="Cancel", 
            command=self.cancel_dialog,
            font=UIFonts.BUTTON,
            bg=UIColors.BG_TERTIARY,
            fg=UIColors.TEXT_PRIMARY,
            activebackground=UIColors.ERROR_LIGHT,
            activeforeground=UIColors.ERROR,
            relief="flat",
            cursor="hand2",
            padx=UISpacing.LG,
            pady=UISpacing.SM,
            bd=0,
            width=12
        )
        cancel_button.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        # Hover effect for cancel button
        def on_cancel_enter(e):
            cancel_button.config(bg=UIColors.ERROR_LIGHT, fg=UIColors.ERROR)
        def on_cancel_leave(e):
            cancel_button.config(bg=UIColors.BG_TERTIARY, fg=UIColors.TEXT_PRIMARY)
        cancel_button.bind("<Enter>", on_cancel_enter)
        cancel_button.bind("<Leave>", on_cancel_leave)
        
        split_button.focus_set()
    
    def create_thumbnails(self):
        """Create the thumbnail grid."""
        self.thumbnail_widgets = {}
        
        for page in range(1, self.total_pages + 1):
            # Card-style page frame
            page_frame = tk.Frame(
                self.scrollable_thumb_frame, 
                bg=UIColors.THUMBNAIL_BG,
                bd=1,
                relief="solid",
                padx=UISpacing.SM,
                pady=UISpacing.SM
            )
            
            thumb_label = self.create_text_placeholder(page, page_frame)
            
            # Mark pages that have thumbnails available (either generated or from .thumbs folder)
            if page in self.generated_thumbnails:
                thumb_label.has_generated_thumb = True
            elif self.has_thumbs_folder:
                thumbnail_path = get_thumbnail_path(self.pdf_path, page)
                if thumbnail_path and os.path.exists(thumbnail_path):
                    thumb_label.thumbnail_path = thumbnail_path
            
            thumb_label.pack(pady=(0, UISpacing.XS))
            
            # Page number label
            page_label = tk.Label(
                page_frame, 
                text=f"Page {page}",
                font=UIFonts.SMALL_BOLD,
                bg=UIColors.THUMBNAIL_BG,
                fg=UIColors.TEXT_PRIMARY
            )
            page_label.pack()
            
            split_button = None
            if page < self.total_pages:
                split_button = tk.Button(
                    page_frame, 
                    text=f"Split after Page {page}",
                    command=lambda p=page: self.toggle_split_point_and_update_layout(p),
                    font=UIFonts.SMALL,
                    bg=UIColors.BG_TERTIARY,
                    fg=UIColors.TEXT_SECONDARY,
                    activebackground=UIColors.SPLIT_HOVER,
                    relief="flat",
                    cursor="hand2",
                    padx=UISpacing.SM,
                    pady=UISpacing.XS,
                    bd=0
                )
                split_button.pack(pady=(UISpacing.XS, 0), fill="x")
                
                # Hover effect for split button
                def make_hover_handlers(btn, page_num):
                    def on_enter(e):
                        if (page_num + 1) not in self.split_points:
                            btn.config(bg=UIColors.PRIMARY_LIGHT, fg=UIColors.PRIMARY)
                    def on_leave(e):
                        if (page_num + 1) not in self.split_points:
                            btn.config(bg=UIColors.BG_TERTIARY, fg=UIColors.TEXT_SECONDARY)
                    return on_enter, on_leave
                
                enter_handler, leave_handler = make_hover_handlers(split_button, page)
                split_button.bind("<Enter>", enter_handler)
                split_button.bind("<Leave>", leave_handler)
            
            # Card hover effect
            def make_card_hover(frame):
                def on_enter(e):
                    frame.config(bg=UIColors.THUMBNAIL_HOVER)
                    for child in frame.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg=UIColors.THUMBNAIL_HOVER)
                def on_leave(e):
                    frame.config(bg=UIColors.THUMBNAIL_BG)
                    for child in frame.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg=UIColors.THUMBNAIL_BG)
                return on_enter, on_leave
            
            card_enter, card_leave = make_card_hover(page_frame)
            page_frame.bind("<Enter>", card_enter)
            page_frame.bind("<Leave>", card_leave)
            
            self.thumbnail_widgets[page] = {
                'frame': page_frame,
                'thumbnail': thumb_label,
                'page_label': page_label,
                'split_button': split_button
            }
        
        self.update_thumbnail_layout()
    
    def create_text_placeholder(self, page_num, parent):
        """Create a text placeholder for a page."""
        lbl = tk.Label(
            parent, 
            text=f"📄\n\nPage {page_num}", 
            width=14, 
            height=9,
            relief="flat",
            bg=UIColors.BG_TERTIARY,
            fg=UIColors.TEXT_SECONDARY,
            font=UIFonts.BODY
        )
        lbl._is_placeholder = True
        return lbl
    
    def toggle_split_point(self, page):
        """Toggle a split point after the specified page."""
        split_after_page = page + 1
        
        if split_after_page in self.split_points:
            self.split_points.remove(split_after_page)
        else:
            self.split_points.append(split_after_page)
        
        self.split_points.sort()
        self.update_split_display()
    
    def toggle_split_point_and_update_layout(self, page):
        """Toggle a split point and update the thumbnail layout."""
        self.toggle_split_point(page)
        self.update_thumbnail_layout()
        self.update_split_button_appearance()
    
    def update_thumbnail_layout(self):
        """Update the thumbnail layout based on current split points and available width."""
        for widget_info in self.thumbnail_widgets.values():
            widget_info['frame'].grid_forget()
        
        # Calculate how many thumbnails fit per row based on canvas width
        # Each thumbnail frame is approximately 140px wide + 10px padding
        thumbnail_width = 150  # Approximate width including padding
        canvas_width = self.thumb_canvas.winfo_width()
        if canvas_width < 100:  # Not yet rendered, use a reasonable default
            canvas_width = 800
        
        max_cols = max(1, canvas_width // thumbnail_width)
        
        current_row = 0
        current_col = 0
        
        for page in range(1, self.total_pages + 1):
            frame = self.thumbnail_widgets[page]['frame']
            
            # Start new row if we hit a split point (after page 1) or exceed max columns
            if page in self.split_points and page > 1:
                current_row += 1
                current_col = 0
            elif current_col >= max_cols:
                current_row += 1
                current_col = 0
            
            frame.grid(row=current_row, column=current_col, padx=5, pady=5, sticky="n")
            current_col += 1
        
        self.scrollable_thumb_frame.update_idletasks()
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))
    
    def _on_canvas_resize(self, event):
        """Handle canvas resize to reflow thumbnails."""
        # Only update if width changed significantly (avoid excessive updates)
        if not hasattr(self, '_last_canvas_width'):
            self._last_canvas_width = 0
        
        if abs(event.width - self._last_canvas_width) > 50:
            self._last_canvas_width = event.width
            self.update_thumbnail_layout()
    
    def update_split_button_appearance(self):
        """Update the appearance of split buttons based on current split points."""
        for page in range(1, self.total_pages):
            split_button = self.thumbnail_widgets[page]['split_button']
            if split_button:
                split_after_page = page + 1
                if split_after_page in self.split_points:
                    split_button.config(
                        bg=UIColors.SPLIT_ACTIVE, 
                        fg="white",
                        text=f"✂️ Remove Split",
                        activebackground=UIColors.ERROR_HOVER
                    )
                else:
                    split_button.config(
                        bg=UIColors.BG_TERTIARY, 
                        fg=UIColors.TEXT_SECONDARY,
                        text=f"Split after Page {page}",
                        activebackground=UIColors.PRIMARY_LIGHT
                    )
    
    def update_split_display(self):
        """Update the split display in the control panel."""
        self.split_display_text.delete(1.0, tk.END)
        
        if len(self.split_points) <= 1:
            # Update header
            if hasattr(self, 'split_count_label'):
                self.split_count_label.config(text="📋 Document Preview (1 document)")
            
            self.split_display_text.insert(tk.END, "No splits defined\n\n")
            self.split_display_text.insert(tk.END, "━" * 30 + "\n\n")
            self.split_display_text.insert(tk.END, "📄  Document 1\n")
            self.split_display_text.insert(tk.END, f"     Pages 1 – {self.total_pages}\n")
            self.split_display_text.insert(tk.END, "     (entire PDF)\n")
        else:
            splits = list(zip(self.split_points, self.split_points[1:] + [self.total_pages + 1]))
            
            # Update header
            if hasattr(self, 'split_count_label'):
                self.split_count_label.config(text=f"📋 Document Preview ({len(splits)} documents)")
            
            self.split_display_text.insert(tk.END, f"Splitting into {len(splits)} documents:\n\n")
            self.split_display_text.insert(tk.END, "━" * 30 + "\n")
            
            for i, (start, end) in enumerate(splits, 1):
                end_page = end - 1
                page_count = end_page - start + 1
                pdf_basename = Path(self.pdf_path).stem
                default_name = f"{pdf_basename}_Part{i}.pdf"
                name = self.split_names.get(i, default_name)
                if not name.lower().endswith('.pdf'):
                    name += '.pdf'
                
                self.split_display_text.insert(tk.END, f"\n📄  Document {i}\n")
                self.split_display_text.insert(tk.END, f"     Pages {start} – {end_page} ({page_count} page{'s' if page_count > 1 else ''})\n")
                self.split_display_text.insert(tk.END, f"     → {name}\n")
        
        self.split_display_text.see(1.0)
    
    def clear_all_splits(self):
        """Clear all split points."""
        self.split_points = [1]
        self.split_names = {}
        self.update_split_display()
        self.update_thumbnail_layout()
        self.update_split_button_appearance()
        messagebox.showinfo("Splits Reset", "All split points have been removed.")
    
    def load_thumbnail_images_delayed(self):
        """Load thumbnail images after UI is fully constructed."""
        try:
            for page, widgets in self.thumbnail_widgets.items():
                thumb_label = widgets['thumbnail']
                thumb_img = None
                
                # First, try to use generated thumbnails (from pdf2image)
                if page in self.generated_thumbnails:
                    try:
                        pil_img = self.generated_thumbnails[page]
                        thumb_img = ImageTk.PhotoImage(pil_img, master=self.root)
                    except Exception as e:
                        logger.error(f"Error converting generated thumbnail for page {page}: {e}")
                
                # Fall back to loading from .thumbs folder
                elif hasattr(thumb_label, 'thumbnail_path'):
                    thumbnail_path = thumb_label.thumbnail_path
                    try:
                        thumb_img = load_thumbnail_image(thumbnail_path, max_size=(120, 160), master=self.root)
                    except Exception as e:
                        logger.error(f"Error loading thumbnail from file for page {page}: {e}")
                
                # Apply the thumbnail if we have one
                if thumb_img:
                    try:
                        thumb_label.config(image=thumb_img, text="", width=0, height=0, 
                                         cursor="hand2", relief="raised", bd=2)
                        thumb_label.image = thumb_img
                        
                        if not hasattr(self, '_image_references'):
                            self._image_references = []
                        self._image_references.append(thumb_img)
                        
                        def make_click_handler(page_num):
                            return lambda e: self.show_full_size_image(None, page_num)
                        
                        thumb_label.bind("<Button-1>", make_click_handler(page))
                        
                    except Exception as e:
                        logger.error(f"Error applying thumbnail for page {page}: {e}")
                        
        except Exception as e:
            logger.error(f"Error during delayed thumbnail loading: {e}")
    
    def edit_names(self):
        """Open dialog to edit document names."""
        if len(self.split_points) <= 1:
            messagebox.showinfo("No Splits", "Create split points first.")
            return
        
        splits = list(zip(self.split_points, self.split_points[1:] + [self.total_pages + 1]))
        
        for i, (start, end) in enumerate(splits):
            end_page = end - 1
            pdf_basename = Path(self.pdf_path).stem
            default_name = f"{pdf_basename}_Part{i+1}.pdf"
            
            name = simpledialog.askstring(
                "Edit Filename", 
                f"Name for Document {i+1} (Pages {start}-{end_page}):", 
                initialvalue=default_name, 
                parent=self.root
            )
            
            if name is not None:
                if name.strip():
                    self.split_names[i] = name.strip()
                else:
                    self.split_names[i] = default_name
        
        self.update_split_display()
    
    def show_full_size_image(self, thumbnail_path, page_num):
        """Show full size image with navigation and split functionality."""
        if not PIL_AVAILABLE:
            return
        
        def split_callback(page):
            self.toggle_split_point_and_update_layout(page)
            return True
        
        show_full_size_image(thumbnail_path, self.pdf_path, page_num, split_callback, parent=self.root)
    
    def do_split(self):
        """Execute the PDF splitting."""
        if len(self.split_points) <= 1:
            self.result = {
                'split': False,
                'message': 'No splits defined. PDF remains unchanged.'
            }
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass
            return
        
        splits = []
        split_ranges = list(zip(self.split_points, self.split_points[1:] + [self.total_pages + 1]))
        
        for i, (start, end) in enumerate(split_ranges):
            end_page = end - 1
            pdf_basename = Path(self.pdf_path).stem
            default_name = f"{pdf_basename}_Part{i+1}.pdf"
            name = self.split_names.get(i, default_name)
            if not name.lower().endswith('.pdf'):
                name += '.pdf'
            splits.append((start, end_page, name))
        
        self.result = {
            'split': True,
            'splits': splits,
            'output_folder': self.output_folder_var.get() or self.output_folder
        }
        
        try:
            self.root.quit()
            self.root.destroy()
        except:
            pass
    
    def cancel_dialog(self):
        """Cancel the dialog."""
        self.result = {'split': False, 'message': 'Cancelled by user.'}
        try:
            self.root.quit()
            self.root.destroy()
        except:
            pass
    
    def center_window(self):
        """Center the window on screen."""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def run(self):
        """Run the dialog and return the result."""
        if self._is_toplevel:
            # For Toplevel, we need to wait for the window to close
            self.root.wait_window()
        else:
            # For Tk root, use mainloop
            self.root.mainloop()
        return self.result


class PDFManualSplitterApp:
    """Main application window with drag and drop support."""
    
    def __init__(self):
        # Use TkinterDnD if available
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("PDF Manual Splitter")
        self.root.geometry("900x700")
        self.root.minsize(1280, 500)  # Match launcher minimum width (1280px)
        self.root.resizable(True, True)
        
        # Position window using environment variables if available
        self.position_window()
        
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        self.output_folder = DEFAULT_OUTPUT_FOLDER
        
        self.setup_ui()
        
        if DND_AVAILABLE:
            self.setup_drag_drop()
    
    def position_window(self):
        """Position window using launcher environment variables if available."""
        try:
            # Check if launched from launcher (environment variables set)
            if 'TOOL_WINDOW_X' in os.environ:
                x = int(os.environ.get('TOOL_WINDOW_X', 100))
                y = int(os.environ.get('TOOL_WINDOW_Y', 100))
                width = int(os.environ.get('TOOL_WINDOW_WIDTH', 900))
                height = int(os.environ.get('TOOL_WINDOW_HEIGHT', 700))
                
                # Ensure minimum size (match launcher minimum width)
                width = max(width, 1280)  # Match launcher minimum width (1280px)
                height = max(height, 600)
                
                self.root.geometry(f"{width}x{height}+{x}+{y}")
                print(f"[INFO] Window positioned at {x},{y} with size {width}x{height}")
            else:
                # Not launched from launcher, use defaults
                self.root.geometry("900x700")
                print("[INFO] Running standalone (not from launcher)")
        except (ValueError, TypeError) as e:
            print(f"[WARNING] Could not position window: {e}")
            self.root.geometry("900x700")
    
    def setup_ui(self):
        """Setup the main UI."""
        self.root.configure(bg=UIColors.BG_SECONDARY)
        
        main_frame = tk.Frame(self.root, bg=UIColors.BG_SECONDARY)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=UISpacing.MD, pady=UISpacing.MD)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Header card
        header_frame = tk.Frame(main_frame, bg=UIColors.BG_PRIMARY, pady=UISpacing.MD)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, UISpacing.MD))
        header_frame.grid_columnconfigure(0, weight=1)
        
        # Title
        title_label = tk.Label(
            header_frame, 
            text="✂️  PDF Manual Splitter", 
            font=UIFonts.TITLE,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.PRIMARY
        )
        title_label.grid(row=0, column=0, pady=(UISpacing.SM, UISpacing.XS))
        
        # Description
        desc_label = tk.Label(
            header_frame, 
            text="Split PDF files manually by selecting split points interactively",
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_SECONDARY
        )
        desc_label.grid(row=1, column=0, pady=(0, UISpacing.SM))
        
        # Drop zone / file selection
        self.create_drop_zone(main_frame)
        
        # Results text area
        result_frame = tk.LabelFrame(
            main_frame, 
            text="  📋 Results  ",
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid"
        )
        result_frame.grid(row=3, column=0, sticky="nsew", pady=(UISpacing.MD, 0))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)
        
        self.result_text = scrolledtext.ScrolledText(
            result_frame, 
            wrap=tk.WORD, 
            font=UIFonts.MONO,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            relief="flat",
            padx=UISpacing.MD,
            pady=UISpacing.SM
        )
        self.result_text.grid(row=0, column=0, sticky="nsew", padx=UISpacing.SM, pady=UISpacing.SM)
        
        # Buttons
        button_frame = tk.Frame(main_frame, bg=UIColors.BG_SECONDARY)
        button_frame.grid(row=4, column=0, pady=UISpacing.MD)
        
        select_button = create_rounded_button(
            button_frame,
            "📂  Select PDF File",
            self.select_file,
            style="primary",
            width=18
        )
        select_button.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        clear_button = create_rounded_button(
            button_frame,
            "Clear Results",
            self.clear_results,
            style="secondary"
        )
        clear_button.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        # Status bar
        status_frame = tk.Frame(main_frame, bg=UIColors.BG_TERTIARY, pady=UISpacing.XS)
        status_frame.grid(row=5, column=0, sticky="ew")
        
        # Status indicators
        status_items = []
        status_items.append(("PyPDF2", PDF_AVAILABLE))
        status_items.append(("PIL", PIL_AVAILABLE))
        status_items.append(("Thumbnails", THUMBNAIL_GENERATION_AVAILABLE))
        status_items.append(("Drag&Drop", DND_AVAILABLE))
        
        status_text_parts = []
        for name, available in status_items:
            icon = "✓" if available else "✗"
            color = UIColors.SUCCESS if available else UIColors.TEXT_MUTED
            status_text_parts.append(f"{icon} {name}")
        
        status_text = "  •  ".join(status_text_parts)
        
        self.status_label = tk.Label(
            status_frame, 
            text=status_text, 
            font=UIFonts.SMALL,
            bg=UIColors.BG_TERTIARY,
            fg=UIColors.TEXT_SECONDARY,
            anchor="w"
        )
        self.status_label.pack(fill="x", padx=UISpacing.SM)
        
        # Welcome message
        self.result_text.insert(tk.END, "Welcome to PDF Manual Splitter!\n")
        self.result_text.insert(tk.END, "━" * 45 + "\n\n")
        
        if DND_AVAILABLE:
            self.result_text.insert(tk.END, "📁 Drag and drop PDF files onto the drop zone\n")
            self.result_text.insert(tk.END, "   or click 'Select PDF File' to begin.\n\n")
        else:
            self.result_text.insert(tk.END, "Click 'Select PDF File' to begin.\n\n")
            self.result_text.insert(tk.END, "💡 Install tkinterdnd2 for drag & drop support:\n")
            self.result_text.insert(tk.END, "   pip install tkinterdnd2\n\n")
    
    def create_drop_zone(self, parent):
        """Create a visual drop zone for files."""
        self.drop_frame = tk.Frame(
            parent, 
            bg=UIColors.DROP_ZONE_BG,
            bd=2,
            relief="flat",
            padx=UISpacing.XL,
            pady=UISpacing.XL
        )
        self.drop_frame.grid(row=2, column=0, sticky="ew", pady=UISpacing.SM)
        
        # Create dashed border effect using a label
        self.drop_frame.config(
            highlightbackground=UIColors.DROP_ZONE_BORDER,
            highlightthickness=2
        )
        
        # Icon
        icon_label = tk.Label(
            self.drop_frame,
            text="📄",
            font=("Segoe UI", 32),
            bg=UIColors.DROP_ZONE_BG
        )
        icon_label.pack(pady=(UISpacing.SM, UISpacing.XS))
        
        # Main text
        if DND_AVAILABLE:
            main_text = "Drag and drop PDF file here"
        else:
            main_text = "Click to select PDF file"
        
        self.drop_label = tk.Label(
            self.drop_frame, 
            text=main_text,
            font=UIFonts.SUBTITLE,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_PRIMARY,
            cursor="hand2"
        )
        self.drop_label.pack()
        
        # Sub text
        sub_label = tk.Label(
            self.drop_frame,
            text="or click to browse",
            font=UIFonts.SMALL,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_MUTED,
            cursor="hand2"
        )
        sub_label.pack(pady=(UISpacing.XS, UISpacing.SM))
        
        # Click handlers
        for widget in [self.drop_frame, icon_label, self.drop_label, sub_label]:
            widget.bind('<Button-1>', lambda e: self.select_file())
            
        # Store references for hover effect
        self._drop_zone_widgets = [self.drop_frame, icon_label, self.drop_label, sub_label]
    
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
            self.process_file(files[0])  # Process first PDF
    
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
        
        return [f for f in files if f.lower().endswith('.pdf')]
    
    def select_file(self):
        """Open file dialog to select PDF."""
        pdf_file = filedialog.askopenfilename(
            title="Select PDF file for manual splitting",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if pdf_file:
            self.process_file(pdf_file)
    
    def process_file(self, pdf_path):
        """Process a PDF file."""
        print(f"[INFO] Processing file: {pdf_path}")
        self.result_text.insert(tk.END, f"\n{'='*60}\n")
        self.result_text.insert(tk.END, f"Manual PDF Splitting\n")
        self.result_text.insert(tk.END, f"{'='*60}\n")
        self.result_text.insert(tk.END, f"File: {pdf_path}\n")
        auto_scroll_text_widget(self.result_text)
        
        # Check page count
        total_pages = get_pdf_page_count(pdf_path)
        if total_pages < 2:
            self.result_text.insert(tk.END, f"⚠ PDF has only {total_pages} page(s) - no splitting needed\n\n")
            return
        
        self.result_text.insert(tk.END, f"Pages: {total_pages}\n")
        self.result_text.insert(tk.END, "Opening split dialog...\n\n")
        auto_scroll_text_widget(self.result_text)
        
        # Show manual split dialog (pass parent to keep it on top)
        dialog = ManualSplitDialog(pdf_path, self.output_folder, parent=self.root)
        result = dialog.run()
        
        if not result or not result.get('split', False):
            message = result.get('message', 'No splitting performed') if result else 'Dialog closed'
            self.result_text.insert(tk.END, f"→ {message}\n\n")
            return
        
        # Process the split
        splits = result.get('splits', [])
        output_folder = result.get('output_folder', self.output_folder)
        
        self.result_text.insert(tk.END, f"→ Starting split into {len(splits)} documents...\n")
        
        success = split_pdf_file(pdf_path, splits, output_folder, result_text=self.result_text)
        
        if success:
            self.result_text.insert(tk.END, f"\n✓ PDF successfully split into {len(splits)} documents!\n")
            self.result_text.insert(tk.END, f"→ Output folder: {output_folder}\n\n")
            messagebox.showinfo("Success", f"PDF split into {len(splits)} documents!\n\nOutput: {output_folder}")
        else:
            self.result_text.insert(tk.END, f"\n✗ Error splitting PDF\n\n")
            messagebox.showerror("Error", "Error splitting PDF")
    
    def clear_results(self):
        """Clear the results text area."""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Results cleared.\n\n")
    
    def run(self):
        """Run the application."""
        self.root.mainloop()


def main():
    """Main entry point."""
    print("[INFO] PDF Manual Splitter initialized")
    logger.info("Starting PDF Manual Splitter")
    
    if not PDF_AVAILABLE:
        print("[ERROR] PyPDF2 not available. Please install: pip install PyPDF2")
        sys.exit(1)
    
    # Command line mode
    if args.pdf_file:
        pdf_file = args.pdf_file
        if not os.path.exists(pdf_file):
            print(f"❌ File not found: {pdf_file}")
            sys.exit(1)
        
        output_folder = args.output_folder or DEFAULT_OUTPUT_FOLDER
        
        print(f"📄 PDF Manual Splitter - Command Line Mode")
        print(f"📁 File: {pdf_file}")
        print(f"📂 Output folder: {output_folder}")
        
        total_pages = get_pdf_page_count(pdf_file)
        if total_pages < 2:
            print(f"⚠ PDF has only {total_pages} page(s) - no splitting needed")
            sys.exit(0)
        
        print(f"📄 PDF has {total_pages} pages")
        print("🔧 Starting manual split dialog...")
        
        dialog = ManualSplitDialog(pdf_file, output_folder)
        result = dialog.run()
        
        if not result or not result.get('split', False):
            message = result.get('message', 'No splitting performed') if result else 'Dialog cancelled'
            print(f"ℹ {message}")
            sys.exit(0)
        
        splits = result.get('splits', [])
        output_folder = result.get('output_folder', output_folder)
        
        print(f"📋 Splitting into {len(splits)} documents...")
        
        success = split_pdf_file(pdf_file, splits, output_folder)
        
        if success:
            print(f"✅ PDF successfully split into {len(splits)} documents!")
            print(f"📂 Output folder: {output_folder}")
        else:
            print("❌ Error splitting PDF")
            sys.exit(1)
    
    else:
        # GUI mode
        print("[INFO] Starting GUI mode...")
        app = PDFManualSplitterApp()
        print("[INFO] GUI window opened")
        app.run()
        print("[INFO] PDF Manual Splitter closed")


if __name__ == "__main__":
    main()
