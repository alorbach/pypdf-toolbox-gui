"""
PDF Combiner Tool

Combine multiple PDF files by visually selecting pages from thumbnails.
Supports drag and drop, auto-selection patterns, and configurable thumbnail sizes.

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
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import logging
import argparse
import os
import re
import io
from pathlib import Path

# PDF manipulation support
try:
    from PyPDF2 import PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# PIL for image conversion
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# PyMuPDF for better PDF rendering
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# Drag and drop support (optional)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# Parse command line arguments
parser = argparse.ArgumentParser(description='PDF Combiner - Visual page selection for combining PDFs')
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
print(f"[INFO] Starting PDF Combiner")
print(f"[INFO] PyPDF2: {'Available' if PDF_AVAILABLE else 'Not available'}")
print(f"[INFO] PIL/Pillow: {'Available' if PIL_AVAILABLE else 'Not available'}")
print(f"[INFO] PyMuPDF: {'Available' if PYMUPDF_AVAILABLE else 'Not available'}")
print(f"[INFO] Drag&Drop: {'Available' if DND_AVAILABLE else 'Not available'}")

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
    
    THUMBNAIL_BG = "#ffffff"
    THUMBNAIL_BORDER = "#e2e8f0"
    THUMBNAIL_SELECTED = "#dbeafe"
    THUMBNAIL_SELECTED_BORDER = "#2563eb"
    SELECTION_BADGE = "#ef4444"


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
    """Create a styled button with consistent appearance."""
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


# ============================================================================
# GUI Application
# ============================================================================

class PDFCombinerApp:
    """Main GUI application for PDF Combiner with visual page selection."""
    
    def __init__(self):
        # Use TkinterDnD if available
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("PDF Combiner - Visual Page Selection")
        self.root.geometry("1800x1200")
        self.root.minsize(1200, 800)
        self.root.resizable(True, True)
        
        # Position window using environment variables if available
        self.position_window()
        
        # Variables
        self.pdf_files = []
        self.all_pages = []  # List of page data dictionaries
        self.selected_pages = []  # List of selected page data in order
        self.pages_by_file = []  # List of lists: pages grouped by file for auto-selection
        
        # Preview size configuration
        self.preview_sizes = {
            'small': (120, 150),
            'big': (200, 250),
            'biggest': (640, 480),
            'huge': (800, 600),
            'massive': (1024, 768),
            'giant': (1280, 960)
        }
        self.current_preview_size = 'big'  # Default to big size
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Setup GUI
        self.setup_gui()
        
        if DND_AVAILABLE:
            self.setup_drag_drop()
    
    def position_window(self):
        """Position window using launcher environment variables if available."""
        try:
            # Check if launched from launcher (environment variables set)
            if 'TOOL_WINDOW_X' in os.environ:
                x = int(os.environ.get('TOOL_WINDOW_X', 100))
                y = int(os.environ.get('TOOL_WINDOW_Y', 100))
                width = int(os.environ.get('TOOL_WINDOW_WIDTH', 1800))
                height = int(os.environ.get('TOOL_WINDOW_HEIGHT', 1200))
                
                # Ensure minimum size
                width = max(width, 1200)
                height = max(height, 800)
                
                self.root.geometry(f"{width}x{height}+{x}+{y}")
                print(f"[INFO] Window positioned at {x},{y} with size {width}x{height}")
            else:
                # Not launched from launcher, use defaults
                self.root.geometry("1800x1200")
                print("[INFO] Running standalone (not from launcher)")
        except (ValueError, TypeError) as e:
            print(f"[WARNING] Could not position window: {e}")
            self.root.geometry("1800x1200")
    
    def setup_gui(self):
        """Setup the main GUI."""
        self.root.configure(bg=UIColors.BG_SECONDARY)
        
        # Main frame
        main_frame = tk.Frame(self.root, bg=UIColors.BG_SECONDARY, padx=UISpacing.MD, pady=UISpacing.MD)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = tk.Label(
            main_frame,
            text="📄 PDF Combiner - Click thumbnails to select pages in order",
            font=UIFonts.TITLE,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.PRIMARY
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, UISpacing.MD))
        
        # Control panel
        self.create_control_panel(main_frame)
        
        # Thumbnail panel
        self.create_thumbnail_panel(main_frame)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Load PDF files to begin")
        status_bar = tk.Label(
            main_frame,
            textvariable=self.status_var,
            font=UIFonts.SMALL,
            bg=UIColors.BG_TERTIARY,
            fg=UIColors.TEXT_SECONDARY,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=UISpacing.SM,
            pady=UISpacing.XS
        )
        status_bar.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(UISpacing.MD, 0))
    
    def create_control_panel(self, parent):
        """Create the left control panel."""
        control_frame = create_card_frame(parent, "  ⚙️ Controls  ", padding=True)
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, UISpacing.MD))
        control_frame.grid_rowconfigure(2, weight=1)
        
        # Load files button with drop zone
        self.create_drop_zone(control_frame)
        
        # File list
        file_list_label = tk.Label(
            control_frame,
            text="Loaded Files:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        file_list_label.pack(anchor=tk.W, pady=(UISpacing.MD, UISpacing.XS))
        
        self.file_listbox = tk.Listbox(
            control_frame,
            height=8,
            font=UIFonts.BODY,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            selectbackground=UIColors.PRIMARY_LIGHT,
            selectforeground=UIColors.TEXT_PRIMARY,
            relief="solid",
            bd=1
        )
        self.file_listbox.pack(fill=tk.BOTH, expand=True, pady=(UISpacing.XS, UISpacing.MD))
        
        # Selection info
        selection_label = tk.Label(
            control_frame,
            text="Selected Pages:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        selection_label.pack(anchor=tk.W, pady=(0, UISpacing.XS))
        
        self.selection_text = tk.Text(
            control_frame,
            height=6,
            wrap=tk.WORD,
            font=UIFonts.SMALL,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            relief="solid",
            bd=1,
            padx=UISpacing.SM,
            pady=UISpacing.SM
        )
        self.selection_text.pack(fill=tk.BOTH, expand=True, pady=(UISpacing.XS, UISpacing.MD))
        
        # Preview size selection
        preview_frame = create_card_frame(control_frame, "  📐 Thumbnail Size  ", padding=True)
        preview_frame.pack(fill=tk.X, pady=(0, UISpacing.MD))
        
        self.preview_size_var = tk.StringVar(value=self.current_preview_size)
        size_frame = tk.Frame(preview_frame, bg=UIColors.BG_PRIMARY)
        size_frame.pack(fill=tk.X, pady=(UISpacing.XS, 0))
        
        for size_name in ['small', 'big', 'biggest', 'huge', 'massive', 'giant']:
            rb = tk.Radiobutton(
                size_frame,
                text=size_name.title(),
                variable=self.preview_size_var,
                value=size_name,
                command=self.change_preview_size,
                font=UIFonts.SMALL,
                bg=UIColors.BG_PRIMARY,
                fg=UIColors.TEXT_PRIMARY,
                activebackground=UIColors.BG_PRIMARY,
                activeforeground=UIColors.PRIMARY,
                selectcolor=UIColors.BG_PRIMARY
            )
            rb.pack(side=tk.LEFT, padx=(0, UISpacing.SM))
        
        # Auto selection buttons
        auto_frame = create_card_frame(control_frame, "  🔄 Auto Selection  ", padding=True)
        auto_frame.pack(fill=tk.X, pady=(0, UISpacing.MD))
        
        self.auto_alternate_btn = create_rounded_button(
            auto_frame,
            "Auto: Alternate Pages",
            self.auto_select_alternate,
            style="secondary"
        )
        self.auto_alternate_btn.pack(fill=tk.X, pady=(UISpacing.XS, UISpacing.XS))
        self.auto_alternate_btn.config(state=tk.DISABLED)
        self._create_tooltip(self.auto_alternate_btn, "Select pages alternating between files:\nFile1 Page1, File2 Page1, File1 Page2, File2 Page2...")
        
        self.auto_reverse_btn = create_rounded_button(
            auto_frame,
            "Auto: Alternate + Reverse",
            self.auto_select_reverse,
            style="secondary"
        )
        self.auto_reverse_btn.pack(fill=tk.X, pady=(0, UISpacing.XS))
        self.auto_reverse_btn.config(state=tk.DISABLED)
        self._create_tooltip(self.auto_reverse_btn, "Alternate with 2nd file in reverse:\nFile1 Page1, File2 LastPage, File1 Page2, File2 2nd-LastPage...")
        
        # Action buttons
        self.clear_btn = create_rounded_button(
            control_frame,
            "Clear Selection",
            self.clear_selection,
            style="danger"
        )
        self.clear_btn.pack(fill=tk.X, pady=(UISpacing.MD, UISpacing.SM))
        self.clear_btn.config(state=tk.DISABLED)
        
        self.save_btn = create_rounded_button(
            control_frame,
            "💾 Save Combined PDF",
            self.save_combined_pdf,
            style="success"
        )
        self.save_btn.pack(fill=tk.X)
        self.save_btn.config(state=tk.DISABLED)
    
    def create_drop_zone(self, parent):
        """Create a visual drop zone for files."""
        drop_frame = tk.Frame(
            parent,
            bg=UIColors.DROP_ZONE_BG,
            bd=2,
            relief="flat",
            padx=UISpacing.MD,
            pady=UISpacing.MD
        )
        drop_frame.pack(fill=tk.X, pady=(0, UISpacing.MD))
        drop_frame.config(
            highlightbackground=UIColors.DROP_ZONE_BORDER,
            highlightthickness=2
        )
        
        if DND_AVAILABLE:
            main_text = "Drag and drop PDF files here"
        else:
            main_text = "Click to select PDF files"
        
        self.drop_label = tk.Label(
            drop_frame,
            text=main_text,
            font=UIFonts.BODY_BOLD,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_PRIMARY,
            cursor="hand2"
        )
        self.drop_label.pack(pady=UISpacing.XS)
        
        sub_label = tk.Label(
            drop_frame,
            text="or click to browse",
            font=UIFonts.SMALL,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_MUTED,
            cursor="hand2"
        )
        sub_label.pack(pady=(0, UISpacing.XS))
        
        # Click handlers
        for widget in [drop_frame, self.drop_label, sub_label]:
            widget.bind('<Button-1>', lambda e: self.load_files())
        
        self._drop_zone_widgets = [drop_frame, self.drop_label, sub_label]
        self.drop_frame = drop_frame
    
    def create_thumbnail_panel(self, parent):
        """Create the right thumbnail panel."""
        thumbnail_frame = create_card_frame(parent, "  📖 PDF Pages - Click to select in order  ", padding=True)
        thumbnail_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        thumbnail_frame.columnconfigure(0, weight=1)
        thumbnail_frame.rowconfigure(0, weight=1)
        
        # Create canvas and scrollbar for thumbnails
        self.canvas = tk.Canvas(
            thumbnail_frame,
            bg=UIColors.BG_SECONDARY,
            highlightthickness=0
        )
        scrollbar = ttk.Scrollbar(thumbnail_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=UIColors.BG_SECONDARY)
        
        canvas_frame_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        def configure_scroll_region(event=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            canvas_width = self.canvas.winfo_width()
            if canvas_width > 1:
                self.canvas.itemconfig(canvas_frame_id, width=canvas_width)
        
        self.scrollable_frame.bind("<Configure>", configure_scroll_region)
        self.canvas.bind("<Configure>", configure_scroll_region)
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=UISpacing.SM, pady=UISpacing.SM)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Bind mouse wheel to canvas
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
    
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
            self.load_files_from_list(files)
    
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
        
        # Filter for PDF files only
        return [f for f in files if f.lower().endswith('.pdf')]
    
    def _on_mousewheel(self, event):
        """Enhanced mouse wheel scrolling."""
        try:
            # Handle different mouse wheel events across platforms
            if event.delta:
                delta = int(-1 * (event.delta / 120))
            else:
                delta = -1 if event.num == 4 else 1
                
            self.canvas.yview_scroll(delta, "units")
        except:
            # Fallback for any mouse wheel issues
            self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
    
    def load_files(self):
        """Open file dialog to load PDF files."""
        file_paths = filedialog.askopenfilenames(
            title="Select PDF files to combine",
            filetypes=[("PDF files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if file_paths:
            self.load_files_from_list(list(file_paths))
    
    def load_files_from_list(self, file_paths):
        """Load PDF files from a list of paths."""
        if not file_paths:
            return
        
        if not PDF_AVAILABLE:
            messagebox.showerror("Error", "PyPDF2 not available. Please install: pip install PyPDF2")
            return
        
        if not PYMUPDF_AVAILABLE:
            messagebox.showerror("Error", "PyMuPDF not available. Please install: pip install pymupdf")
            return
        
        if not PIL_AVAILABLE:
            messagebox.showerror("Error", "PIL/Pillow not available. Please install: pip install Pillow")
            return
        
        self.pdf_files = list(file_paths)
        self.all_pages = []
        self.selected_pages = []
        self.pages_by_file = []
        
        # Clear previous thumbnails
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # Update file list
        self.file_listbox.delete(0, tk.END)
        for file_path in self.pdf_files:
            self.file_listbox.insert(tk.END, os.path.basename(file_path))
        
        self.status_var.set("Loading PDF pages and generating thumbnails...")
        self.root.update()
        
        try:
            self.load_pdf_thumbnails()
            self.status_var.set(f"Loaded {len(self.all_pages)} pages from {len(self.pdf_files)} files")
            self.clear_btn.config(state=tk.NORMAL)
            self.auto_alternate_btn.config(state=tk.NORMAL if len(self.pdf_files) >= 2 else tk.DISABLED)
            self.auto_reverse_btn.config(state=tk.NORMAL if len(self.pdf_files) >= 2 else tk.DISABLED)
        except Exception as e:
            error_msg = f"Error loading PDFs: {str(e)}"
            messagebox.showerror("Error", error_msg)
            logger.error(error_msg)
            self.status_var.set("Error loading files")
    
    def change_preview_size(self):
        """Change the preview size and regenerate thumbnails if files are loaded."""
        new_size = self.preview_size_var.get()
        if new_size != self.current_preview_size and self.pdf_files:
            self.current_preview_size = new_size
            self.status_var.set("Regenerating thumbnails with new size...")
            self.root.update()
            
            # Clear current thumbnails
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            
            # Regenerate thumbnails with new size
            self.load_pdf_thumbnails()
    
    def load_pdf_thumbnails(self):
        """Load PDF thumbnails with visual preview."""
        row = 0
        # Adjust max columns based on preview size
        if self.current_preview_size == 'small':
            max_cols = 8
        elif self.current_preview_size == 'big':
            max_cols = 5
        elif self.current_preview_size == 'biggest':
            max_cols = 2
        elif self.current_preview_size == 'huge':
            max_cols = 1
        elif self.current_preview_size == 'massive':
            max_cols = 1
        else:  # giant (1280x960)
            max_cols = 1
        
        for file_index, file_path in enumerate(self.pdf_files):
            # Initialize pages list for this file
            file_pages = []
            
            # Add file separator label
            file_label = tk.Label(
                self.scrollable_frame,
                text=f"📄 {os.path.basename(file_path)}",
                font=UIFonts.BODY_BOLD,
                bg=UIColors.BG_SECONDARY,
                fg=UIColors.TEXT_PRIMARY
            )
            file_label.grid(row=row, column=0, columnspan=max_cols+1, sticky=tk.W, padx=UISpacing.SM, pady=(UISpacing.MD, UISpacing.SM))
            row += 1
            
            # Open PDF with PyMuPDF for better rendering
            try:
                pdf_doc = fitz.open(file_path)
                total_pages = len(pdf_doc)
                
                self.status_var.set(f"Loading thumbnails for {os.path.basename(file_path)} ({total_pages} pages)...")
                self.root.update()
                
                # Process pages in rows
                current_row_start = row
                col = 0
                
                for page_index in range(total_pages):
                    try:
                        page = pdf_doc[page_index]
                        
                        # Render page as image with quality based on target size
                        target_size = self.preview_sizes[self.current_preview_size]
                        
                        # Calculate appropriate matrix scale based on target size
                        if self.current_preview_size in ['small', 'big']:
                            mat_scale = 0.4
                        elif self.current_preview_size == 'biggest':
                            mat_scale = 0.6
                        elif self.current_preview_size == 'huge':
                            mat_scale = 0.8
                        elif self.current_preview_size == 'massive':
                            mat_scale = 1.0
                        else:  # giant
                            mat_scale = 1.2
                        
                        mat = fitz.Matrix(mat_scale, mat_scale)
                        pix = page.get_pixmap(matrix=mat)
                        img_data = pix.tobytes("ppm")
                        
                        # Convert to PIL Image
                        pil_image = Image.open(io.BytesIO(img_data))
                        
                        # Resize to configurable thumbnail size maintaining aspect ratio
                        pil_image.thumbnail(target_size, Image.Resampling.LANCZOS)
                        
                        # Convert to PhotoImage for tkinter
                        photo = ImageTk.PhotoImage(pil_image)
                        
                        # Create thumbnail button
                        page_data = {
                            'file_index': file_index,
                            'page_index': page_index,
                            'file_path': file_path,
                            'photo': photo,
                            'pil_image': pil_image,
                            'selected': False
                        }
                        
                        # Calculate row for this thumbnail
                        thumb_row = current_row_start + (col // (max_cols + 1))
                        thumb_col = col % (max_cols + 1)
                        
                        thumb_frame = tk.Frame(
                            self.scrollable_frame,
                            relief=tk.RAISED,
                            borderwidth=2,
                            bg=UIColors.THUMBNAIL_BG,
                            highlightbackground=UIColors.THUMBNAIL_BORDER,
                            highlightthickness=1
                        )
                        thumb_frame.grid(row=thumb_row, column=thumb_col, padx=UISpacing.XS, pady=UISpacing.XS, sticky="n")
                        
                        # Thumbnail button
                        thumb_btn = tk.Button(
                            thumb_frame,
                            image=photo,
                            command=lambda pd=page_data, tf=thumb_frame: self.toggle_page_selection(pd, tf),
                            bg=UIColors.THUMBNAIL_BG,
                            relief=tk.FLAT,
                            cursor="hand2",
                            bd=0
                        )
                        thumb_btn.pack(padx=UISpacing.XS, pady=UISpacing.XS)
                        
                        # Page info label
                        page_info = tk.Label(
                            thumb_frame,
                            text=f"Page {page_index + 1}",
                            font=UIFonts.SMALL,
                            bg=UIColors.THUMBNAIL_BG,
                            fg=UIColors.TEXT_SECONDARY
                        )
                        page_info.pack()
                        
                        # Selection number label (initially hidden)
                        selection_label = tk.Label(
                            thumb_frame,
                            text="",
                            font=UIFonts.BODY_BOLD,
                            bg=UIColors.SELECTION_BADGE,
                            fg="white",
                            relief=tk.RAISED,
                            bd=2
                        )
                        
                        page_data['thumb_frame'] = thumb_frame
                        page_data['selection_label'] = selection_label
                        
                        self.all_pages.append(page_data)
                        file_pages.append(page_data)
                        
                        col += 1
                        
                    except Exception as e:
                        logger.error(f"Error processing page {page_index + 1} of {os.path.basename(file_path)}: {e}")
                        continue
                
                # Add this file's pages to the organized structure
                self.pages_by_file.append(file_pages)
                
                # Update row counter for next file
                row = current_row_start + ((col - 1) // (max_cols + 1)) + 2
                
                pdf_doc.close()
                
            except Exception as e:
                error_msg = f"Error loading file {file_path}: {e}"
                logger.error(error_msg)
                messagebox.showwarning("File Error", f"Could not load {os.path.basename(file_path)}: {e}")
                continue
        
        # Ensure canvas scroll region is updated
        self.root.after(100, self._update_scroll_region)
    
    def _update_scroll_region(self):
        """Update canvas scroll region after all widgets are rendered."""
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        # Ensure canvas can be scrolled with mouse wheel
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    
    def toggle_page_selection(self, page_data, thumb_frame):
        """Toggle page selection state."""
        if page_data['selected']:
            # Unselect page
            page_data['selected'] = False
            page_data['selection_label'].pack_forget()
            thumb_frame.config(
                bg=UIColors.THUMBNAIL_BG,
                relief=tk.RAISED,
                borderwidth=2,
                highlightbackground=UIColors.THUMBNAIL_BORDER
            )
            
            # Remove from selected list
            self.selected_pages = [p for p in self.selected_pages if p != page_data]
            
            # Renumber remaining selected pages
            self._renumber_selected_pages()
        else:
            # Select page
            page_data['selected'] = True
            selection_number = len(self.selected_pages) + 1
            page_data['selection_label'].config(text=str(selection_number))
            page_data['selection_label'].pack()
            thumb_frame.config(
                bg=UIColors.THUMBNAIL_SELECTED,
                relief=tk.SUNKEN,
                borderwidth=3,
                highlightbackground=UIColors.THUMBNAIL_SELECTED_BORDER
            )
            
            # Add to selected list
            self.selected_pages.append(page_data)
        
        # Update selection display
        self.update_selection_display()
        
        # Enable/disable save button
        self.save_btn.config(state=tk.NORMAL if self.selected_pages else tk.DISABLED)
    
    def _renumber_selected_pages(self):
        """Renumber all selected pages to maintain sequential numbering."""
        for i, page_data in enumerate(self.selected_pages):
            page_data['selection_label'].config(text=str(i + 1))
    
    def clear_selection(self):
        """Clear all page selections."""
        for page_data in self.selected_pages:
            page_data['selected'] = False
            page_data['selection_label'].pack_forget()
            page_data['thumb_frame'].config(
                bg=UIColors.THUMBNAIL_BG,
                relief=tk.RAISED,
                borderwidth=2,
                highlightbackground=UIColors.THUMBNAIL_BORDER
            )
        
        self.selected_pages = []
        self.update_selection_display()
        self.save_btn.config(state=tk.DISABLED)
    
    def update_selection_display(self):
        """Update the selection text display."""
        if not self.selected_pages:
            self.selection_text.delete(1.0, tk.END)
            self.selection_text.insert(tk.END, "No pages selected")
            return
        
        text = f"Selected {len(self.selected_pages)} pages:\n\n"
        for i, page_data in enumerate(self.selected_pages):
            filename = os.path.basename(page_data['file_path'])
            text += f"{i+1}. {filename} - Page {page_data['page_index'] + 1}\n"
        
        self.selection_text.delete(1.0, tk.END)
        self.selection_text.insert(tk.END, text)
    
    def save_combined_pdf(self):
        """Save the combined PDF with selected pages."""
        if not self.selected_pages:
            messagebox.showwarning("No Selection", "Please select pages to combine")
            return
        
        # Get save location
        output_path = filedialog.asksaveasfilename(
            title="Save Combined PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if not output_path:
            return
        
        try:
            self.status_var.set("Creating combined PDF...")
            self.root.update()
            
            pdf_writer = PdfWriter()
            
            # Process selected pages in order
            for page_data in self.selected_pages:
                # Open the PDF file and get the specific page
                pdf_reader = PdfReader(page_data['file_path'])
                page = pdf_reader.pages[page_data['page_index']]
                pdf_writer.add_page(page)
            
            # Save the combined PDF
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            self.status_var.set(f"Successfully saved {len(self.selected_pages)} pages to {os.path.basename(output_path)}")
            messagebox.showinfo("Success", f"Combined PDF saved successfully!\n\nLocation: {output_path}")
            
        except Exception as e:
            error_msg = f"Error saving PDF: {str(e)}"
            messagebox.showerror("Error", error_msg)
            logger.error(error_msg)
            self.status_var.set("Error saving file")
    
    def auto_select_alternate(self):
        """Auto-select pages alternating between files: file1 p1, file2 p1, file1 p2, file2 p2, etc."""
        if len(self.pages_by_file) < 2:
            messagebox.showinfo("Info", "You need at least 2 PDF files loaded for auto-selection.")
            return
        
        # Clear current selection first
        self.clear_selection()
        
        # Find the maximum number of pages in any file
        max_pages = max(len(file_pages) for file_pages in self.pages_by_file)
        
        # Alternate between files page by page
        for page_num in range(max_pages):
            for file_index, file_pages in enumerate(self.pages_by_file):
                if page_num < len(file_pages):
                    page_data = file_pages[page_num]
                    if not page_data['selected']:
                        self._select_page(page_data)
        
        self.status_var.set(f"Auto-selected {len(self.selected_pages)} pages in alternating pattern")
    
    def auto_select_reverse(self):
        """Auto-select pages: file1 forward, file2 reverse: file1 p1, file2 last, file1 p2, file2 second-last, etc."""
        if len(self.pages_by_file) < 2:
            messagebox.showinfo("Info", "You need at least 2 PDF files loaded for auto-selection.")
            return
        
        # Clear current selection first
        self.clear_selection()
        
        # Get first two files (main pattern)
        file1_pages = self.pages_by_file[0]
        file2_pages = self.pages_by_file[1]
        
        # Find the minimum number of pages to ensure we have pairs
        min_pages = min(len(file1_pages), len(file2_pages))
        
        # Alternate: file1 forward, file2 reverse
        for i in range(min_pages):
            # Select from file1 (forward direction)
            if i < len(file1_pages):
                page_data = file1_pages[i]
                if not page_data['selected']:
                    self._select_page(page_data)
            
            # Select from file2 (reverse direction)
            reverse_index = len(file2_pages) - 1 - i
            if reverse_index >= 0 and reverse_index < len(file2_pages):
                page_data = file2_pages[reverse_index]
                if not page_data['selected']:
                    self._select_page(page_data)
        
        # Handle additional files if any (continue with normal alternating)
        if len(self.pages_by_file) > 2:
            max_pages = max(len(file_pages) for file_pages in self.pages_by_file[2:])
            for page_num in range(max_pages):
                for file_index in range(2, len(self.pages_by_file)):
                    file_pages = self.pages_by_file[file_index]
                    if page_num < len(file_pages):
                        page_data = file_pages[page_num]
                        if not page_data['selected']:
                            self._select_page(page_data)
        
        self.status_var.set(f"Auto-selected {len(self.selected_pages)} pages in alternating + reverse pattern")
    
    def _select_page(self, page_data):
        """Helper method to select a page programmatically."""
        if not page_data['selected']:
            page_data['selected'] = True
            selection_number = len(self.selected_pages) + 1
            page_data['selection_label'].config(text=str(selection_number))
            page_data['selection_label'].pack()
            page_data['thumb_frame'].config(
                bg=UIColors.THUMBNAIL_SELECTED,
                relief=tk.SUNKEN,
                borderwidth=3,
                highlightbackground=UIColors.THUMBNAIL_SELECTED_BORDER
            )
            
            self.selected_pages.append(page_data)
        
        # Update display
        self.update_selection_display()
        self.save_btn.config(state=tk.NORMAL if self.selected_pages else tk.DISABLED)
    
    def _create_tooltip(self, widget, text):
        """Create a tooltip for a widget."""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = tk.Label(
                tooltip,
                text=text,
                background=UIColors.WARNING_LIGHT,
                relief=tk.SOLID,
                borderwidth=1,
                font=UIFonts.SMALL,
                padx=UISpacing.SM,
                pady=UISpacing.XS
            )
            label.pack()
            
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
    
    def run(self):
        """Run the application."""
        self.root.mainloop()


def main():
    """Main entry point."""
    print("[INFO] PDF Combiner initialized")
    logger.info("Starting PDF Combiner")
    
    if not PDF_AVAILABLE:
        print("[ERROR] PyPDF2 not available. Please install: pip install PyPDF2")
        messagebox.showerror("Error", "PyPDF2 not available. Please install: pip install PyPDF2")
        sys.exit(1)
    
    if not PYMUPDF_AVAILABLE:
        print("[ERROR] PyMuPDF not available. Please install: pip install pymupdf")
        messagebox.showerror("Error", "PyMuPDF not available. Please install: pip install pymupdf")
        sys.exit(1)
    
    if not PIL_AVAILABLE:
        print("[ERROR] PIL/Pillow not available. Please install: pip install Pillow")
        messagebox.showerror("Error", "PIL/Pillow not available. Please install: pip install Pillow")
        sys.exit(1)
    
    print("[INFO] Starting GUI mode...")
    app = PDFCombinerApp()
    print("[INFO] GUI window opened")
    app.run()
    print("[INFO] PDF Combiner closed")


if __name__ == "__main__":
    main()
