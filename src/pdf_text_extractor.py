"""
PDF Text Extractor Tool

Extract text from PDF files using Python-based OCR or Azure AI Document Intelligence.
Supports multiple output formats: Text, Markdown, and JSON.

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
import time
import logging
import argparse
import os
import json
import re
import tempfile
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolledtext

# HTTP requests for Azure API
try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# YAML for config loading
try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False

# Azure Identity support for Entra ID authentication
try:
    from azure.identity import ClientSecretCredential
    AZURE_IDENTITY_AVAILABLE = True
except ImportError:
    AZURE_IDENTITY_AVAILABLE = False

# PDF text extraction with PyMuPDF
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# OCR support
try:
    import ocrmypdf
    OCRMYPDF_AVAILABLE = True
except ImportError:
    OCRMYPDF_AVAILABLE = False

# PIL for image support
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# Import shared Azure config
try:
    from utils.azure_config import AzureAIConfig, get_azure_config
    AZURE_CONFIG_AVAILABLE = True
except ImportError:
    AZURE_CONFIG_AVAILABLE = False

# Parse command line arguments
parser = argparse.ArgumentParser(
    description='PDF Text Extractor - Extract text from PDFs using Python or Azure AI',
    epilog='''
Examples:
  %(prog)s document.pdf                     # Extract text from single PDF
  %(prog)s folder/                          # Process all PDFs in folder
  %(prog)s folder/ --recursive              # Process PDFs recursively
  %(prog)s --method python document.pdf     # Use Python-based extraction
  %(prog)s --method azure document.pdf      # Use Azure AI extraction
  %(prog)s --format markdown document.pdf   # Output as markdown
  %(prog)s --gui                            # Start GUI mode
  %(prog)s                                  # Start GUI mode (default)
    ''',
    formatter_class=argparse.RawDescriptionHelpFormatter
)
parser.add_argument('input_path', nargs='?', help='PDF file or folder to process')
parser.add_argument('--debug', action='store_true', help='Enable debug output')
parser.add_argument('--gui', action='store_true', help='Start GUI mode')
parser.add_argument('--recursive', '-r', action='store_true', help='Process folders recursively')
parser.add_argument('--force', action='store_true', help='Overwrite existing output files')
parser.add_argument('--format', choices=['text', 'markdown', 'json'], default='text',
                    help='Output format: text (default), markdown, or json')
parser.add_argument('--method', choices=['python', 'azure', 'ocr'], default='python',
                    help='Extraction method: python (default), azure, or ocr')
parser.add_argument('--output-dir', '-o', help='Output directory for extracted files')
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
print(f"[INFO] Starting PDF Text Extractor")
print(f"[INFO] PyMuPDF: {'Available' if PYMUPDF_AVAILABLE else 'Not available'}")
print(f"[INFO] OCRmyPDF: {'Available' if OCRMYPDF_AVAILABLE else 'Not available'}")
print(f"[INFO] Requests: {'Available' if REQUESTS_AVAILABLE else 'Not available'}")
print(f"[INFO] Azure Identity: {'Available' if AZURE_IDENTITY_AVAILABLE else 'Not available'}")
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


def auto_scroll_text_widget(text_widget):
    """Auto-scroll text widget to bottom."""
    text_widget.see(tk.END)
    text_widget.update()


# ============================================================================
# Text Extraction Functions
# ============================================================================

def extract_text_python(pdf_path, output_format='text'):
    """
    Extract text from PDF using PyMuPDF (fitz).
    
    This is a pure Python method that doesn't require Azure.
    
    Args:
        pdf_path: Path to the PDF file
        output_format: 'text', 'markdown', or 'json'
    
    Returns:
        Extracted content in the specified format
    """
    if not PYMUPDF_AVAILABLE:
        raise Exception("PyMuPDF not available. Install with: pip install pymupdf")
    
    logger.info(f"Extracting text with PyMuPDF: {pdf_path}")
    
    try:
        doc = fitz.open(pdf_path)
        pages_data = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text()
            
            pages_data.append({
                'page_number': page_num + 1,
                'content': text.strip(),
                'width': page.rect.width,
                'height': page.rect.height
            })
        
        doc.close()
        
        if output_format == 'json':
            return {
                'method': 'python',
                'source_file': os.path.basename(pdf_path),
                'total_pages': len(pages_data),
                'extraction_timestamp': datetime.now().isoformat(),
                'pages': pages_data
            }
        
        elif output_format == 'markdown':
            md_content = f"# {Path(pdf_path).stem}\n\n"
            md_content += f"*Extracted from {os.path.basename(pdf_path)}*\n\n"
            md_content += "---\n\n"
            
            for page in pages_data:
                md_content += f"## Page {page['page_number']}\n\n"
                md_content += page['content'] + "\n\n"
            
            return md_content
        
        else:  # text format
            text_content = ""
            for page in pages_data:
                header_line = "=" * 80
                page_header = f"Page {page['page_number']}"
                centered_header = page_header.center(80)
                text_content += f"{header_line}\n{centered_header}\n{header_line}\n"
                text_content += page['content'] + "\n\n"
            
            return text_content.strip()
        
    except Exception as e:
        logger.error(f"PyMuPDF extraction failed: {e}")
        raise


def extract_text_ocr(pdf_path, output_format='text'):
    """
    Extract text from PDF using OCR (OCRmyPDF + PyMuPDF).
    
    Good for scanned documents where text is in images.
    
    Args:
        pdf_path: Path to the PDF file
        output_format: 'text', 'markdown', or 'json'
    
    Returns:
        Extracted content in the specified format
    """
    if not OCRMYPDF_AVAILABLE:
        raise Exception("OCRmyPDF not available. Install with: pip install ocrmypdf")
    if not PYMUPDF_AVAILABLE:
        raise Exception("PyMuPDF not available. Install with: pip install pymupdf")
    
    logger.info(f"Extracting text with OCR: {pdf_path}")
    
    try:
        # Create temp file for OCR'd PDF
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_path = temp_file.name
        
        try:
            # Run OCR on the PDF
            logger.debug("Running OCR on PDF...")
            try:
                ocrmypdf.ocr(
                    pdf_path,
                    temp_path,
                    skip_text=True,
                    output_type='pdf',
                    language='eng+deu',  # English and German
                    deskew=True
                )
            except Exception as ocr_error:
                logger.warning(f"Advanced OCR failed, trying basic: {ocr_error}")
                ocrmypdf.ocr(
                    pdf_path,
                    temp_path,
                    skip_text=True,
                    output_type='pdf',
                    language='eng'
                )
            
            # Now extract text from the OCR'd PDF
            return extract_text_python(temp_path, output_format)
            
        finally:
            if os.path.exists(temp_path):
                os.unlink(temp_path)
                
    except Exception as e:
        logger.error(f"OCR extraction failed: {e}")
        raise


def extract_text_azure(pdf_path, output_format='text', config=None):
    """
    Extract text from PDF using Azure AI Document Intelligence.
    
    Provides high-quality extraction with layout preservation.
    
    Args:
        pdf_path: Path to the PDF file
        output_format: 'text', 'markdown', or 'json'
        config: AzureAIConfig instance
    
    Returns:
        Extracted content in the specified format
    """
    if not REQUESTS_AVAILABLE:
        raise Exception("Requests library not available. Install with: pip install requests")
    
    if config is None:
        if AZURE_CONFIG_AVAILABLE:
            config = get_azure_config()
        else:
            raise Exception("Azure configuration not available")
    
    if not config.is_doc_intel_configured():
        raise Exception("Azure Document Intelligence not configured. Please configure endpoint and API key.")
    
    logger.info(f"Extracting text with Azure AI: {pdf_path}")
    
    try:
        endpoint = config.doc_intel_endpoint.rstrip('/')
        api_key = config.doc_intel_api_key
        
        # Determine URL based on format
        if output_format == 'markdown':
            url = f"{endpoint}/formrecognizer/documentModels/prebuilt-layout:analyze?api-version=2023-07-31&outputContentFormat=markdown"
        else:
            url = f"{endpoint}/formrecognizer/documentModels/prebuilt-read:analyze?api-version=2023-07-31"
        
        headers = {
            'Content-Type': 'application/pdf',
            'Ocp-Apim-Subscription-Key': api_key,
        }
        
        # Get timeout settings (default to longer for large PDFs)
        request_timeout = config.timeout
        polling_timeout = config.polling_timeout
        max_polling_attempts = polling_timeout // 3  # Poll every 3 seconds
        
        # Send PDF for analysis
        logger.info(f"Submitting PDF to Azure AI (timeout: {request_timeout}s)...")
        with open(pdf_path, 'rb') as f:
            response = requests.post(url, headers=headers, data=f, timeout=request_timeout)
        
        response.raise_for_status()
        operation_url = response.headers.get("Operation-Location")
        
        if not operation_url:
            raise Exception("No operation URL in response")
        
        # Poll for results with longer timeout
        logger.info(f"Waiting for Azure AI processing (max {polling_timeout}s)...")
        status_headers = {'Ocp-Apim-Subscription-Key': api_key}
        
        start_time = time.time()
        for attempt in range(max_polling_attempts):
            # Check if we've exceeded total timeout
            elapsed = time.time() - start_time
            if elapsed > polling_timeout:
                raise Exception(f"Azure AI processing timed out after {elapsed:.0f} seconds")
            
            try:
                result = requests.get(operation_url, headers=status_headers, timeout=30)
                result.raise_for_status()
                result_json = result.json()
                status = result_json.get('status', '')
                
                if attempt % 10 == 0 or status != 'running':  # Log every 10th attempt or status changes
                    logger.info(f"Status check {attempt + 1}/{max_polling_attempts}: {status} (elapsed: {elapsed:.0f}s)")
                else:
                    logger.debug(f"Status check {attempt + 1}: {status}")
                
                if status == 'succeeded':
                    logger.info(f"Azure AI processing completed in {elapsed:.0f} seconds")
                    break
                elif status == 'failed':
                    error = result_json.get('error', {})
                    raise Exception(f"Azure AI processing failed: {error.get('message', 'Unknown error')}")
                
            except requests.exceptions.Timeout:
                logger.warning(f"Polling request {attempt + 1} timed out, retrying...")
                continue
            except requests.exceptions.RequestException as e:
                logger.warning(f"Polling request {attempt + 1} failed: {e}, retrying...")
                time.sleep(2)  # Wait a bit longer on error
                continue
            
            # Wait before next poll (3 seconds for large PDFs)
            time.sleep(3)
        else:
            elapsed = time.time() - start_time
            raise Exception(f"Azure AI processing timed out after {elapsed:.0f} seconds ({max_polling_attempts} attempts)")
        
        # Process results
        analyze_result = result_json.get('analyzeResult', {})
        
        if output_format == 'json':
            return {
                'method': 'azure',
                'source_file': os.path.basename(pdf_path),
                'extraction_timestamp': datetime.now().isoformat(),
                'azure_result': result_json
            }
        
        elif output_format == 'markdown':
            content = analyze_result.get('content', '')
            return post_process_markdown(content)
        
        else:  # text format
            return extract_text_with_layout(result_json)
        
    except Exception as e:
        logger.error(f"Azure AI extraction failed: {e}")
        raise


def extract_text_with_layout(result_json):
    """Extract text from Azure response preserving layout."""
    try:
        analyze_result = result_json.get('analyzeResult', {})
        pages = analyze_result.get('pages', [])
        
        if not pages:
            content = analyze_result.get('content', '')
            return content if content else "No content extracted"
        
        formatted_pages = []
        
        for page_num, page in enumerate(pages, 1):
            page_width = page.get('width', 100)
            lines = page.get('lines', [])
            
            if not lines:
                header_line = "=" * 80
                centered_header = f"Page {page_num}".center(80)
                formatted_pages.append(f"{header_line}\n{centered_header}\n{header_line}\n(No content)\n")
                continue
            
            # Group lines by Y coordinate
            line_groups = []
            for line in lines:
                polygon = line.get('polygon', [])
                content = line.get('content', '')
                
                if len(polygon) >= 4 and content.strip():
                    y_coord = (polygon[1] + polygon[3]) / 2
                    x_coord = min(polygon[0], polygon[6]) if len(polygon) >= 8 else polygon[0]
                    line_groups.append({
                        'y': y_coord,
                        'x': x_coord,
                        'content': content
                    })
            
            # Sort by Y then X
            line_groups.sort(key=lambda l: (l['y'], l['x']))
            
            # Group into rows
            tolerance = 0.1
            rows = []
            current_row = []
            current_y = None
            
            for line_data in line_groups:
                if current_y is None or abs(line_data['y'] - current_y) <= tolerance:
                    current_row.append(line_data)
                    current_y = line_data['y'] if current_y is None else current_y
                else:
                    if current_row:
                        rows.append(current_row)
                    current_row = [line_data]
                    current_y = line_data['y']
            
            if current_row:
                rows.append(current_row)
            
            # Format page
            header_line = "=" * 80
            centered_header = f"Page {page_num}".center(80)
            page_content = f"{header_line}\n{centered_header}\n{header_line}\n"
            
            for row in rows:
                if len(row) == 1:
                    page_content += row[0]['content'] + '\n'
                else:
                    row.sort(key=lambda l: l['x'])
                    row_text = ""
                    last_x = 0
                    
                    for i, line_data in enumerate(row):
                        if i > 0:
                            char_width = page_width / 80
                            space_chars = max(1, int((line_data['x'] - last_x) / char_width))
                            row_text += ' ' * min(space_chars, 20)
                        row_text += line_data['content']
                        last_x = line_data['x'] + len(line_data['content']) * (page_width / 80)
                    
                    page_content += row_text + '\n'
            
            formatted_pages.append(page_content)
        
        return '\n'.join(formatted_pages)
        
    except Exception as e:
        logger.error(f"Layout extraction failed: {e}")
        # Fallback to simple content
        content = result_json.get('analyzeResult', {}).get('content', '')
        return content if content else "Error extracting text"


def post_process_markdown(content):
    """Post-process markdown content for better quality."""
    if not content:
        return content
    
    # Reduce excessive blank lines
    content = re.sub(r'\n{4,}', '\n\n\n', content)
    
    # Fix heading spacing
    content = re.sub(r'\n(#{1,6})\s+([^\n]+)\n([^\n#])', r'\n\1 \2\n\n\3', content)
    
    # Clean up excessive spaces
    content = re.sub(r' {3,}', ' ', content)
    
    return content.strip()


def save_extracted_text(content, pdf_path, output_format='text', output_dir=None):
    """Save extracted content to file."""
    try:
        # Determine output path
        if output_dir:
            out_dir = Path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            base_name = Path(pdf_path).stem
        else:
            out_dir = Path(pdf_path).parent
            base_name = Path(pdf_path).stem
        
        # Determine extension
        extensions = {
            'text': '.txt',
            'markdown': '.md',
            'json': '.json'
        }
        ext = extensions.get(output_format, '.txt')
        
        output_path = out_dir / f"{base_name}{ext}"
        
        # Save content
        if output_format == 'json':
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(content, f, indent=2, ensure_ascii=False)
        else:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
        
        logger.info(f"Saved to: {output_path}")
        return output_path
        
    except Exception as e:
        logger.error(f"Failed to save: {e}")
        raise


def process_pdf_file(pdf_path, method='python', output_format='text', output_dir=None, 
                     force=False, azure_config=None):
    """Process a single PDF file."""
    try:
        logger.info(f"Processing: {pdf_path} (method={method}, format={output_format})")
        
        # Check if output exists
        extensions = {'text': '.txt', 'markdown': '.md', 'json': '.json'}
        ext = extensions.get(output_format, '.txt')
        
        if output_dir:
            output_path = Path(output_dir) / f"{Path(pdf_path).stem}{ext}"
        else:
            output_path = Path(pdf_path).with_suffix(ext)
        
        if output_path.exists() and not force:
            logger.info(f"Output exists, skipping: {output_path}")
            return {'success': True, 'skipped': True, 'output': output_path}
        
        # Extract text based on method
        if method == 'python':
            content = extract_text_python(pdf_path, output_format)
        elif method == 'ocr':
            content = extract_text_ocr(pdf_path, output_format)
        elif method == 'azure':
            content = extract_text_azure(pdf_path, output_format, azure_config)
        else:
            raise Exception(f"Unknown method: {method}")
        
        # Save content
        saved_path = save_extracted_text(content, pdf_path, output_format, output_dir)
        
        return {'success': True, 'skipped': False, 'output': saved_path}
        
    except Exception as e:
        logger.error(f"Failed to process {pdf_path}: {e}")
        return {'success': False, 'error': str(e)}


def find_pdf_files(input_path, recursive=False):
    """Find PDF files in the given path."""
    path_obj = Path(input_path)
    
    if path_obj.is_file():
        if path_obj.suffix.lower() == '.pdf':
            return [path_obj]
        else:
            return []
    elif path_obj.is_dir():
        if recursive:
            return list(path_obj.rglob('*.pdf'))
        else:
            return list(path_obj.glob('*.pdf'))
    else:
        return []


# ============================================================================
# GUI Application
# ============================================================================

class PDFTextExtractorApp:
    """Main GUI application for PDF Text Extractor."""
    
    def __init__(self):
        # Use TkinterDnD if available
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("PDF Text Extractor")
        self.root.geometry("1000x480")
        self.root.minsize(800, 450)
        self.root.resizable(True, True)
        
        # Position window
        self.position_window()
        
        self.root.configure(bg=UIColors.BG_SECONDARY)
        
        # Load Azure config if available
        self.azure_config = None
        if AZURE_CONFIG_AVAILABLE:
            try:
                self.azure_config = get_azure_config()
            except Exception as e:
                logger.warning(f"Could not load Azure config: {e}")
        
        self.setup_ui()
        
        if DND_AVAILABLE:
            self.setup_drag_drop()
    
    def position_window(self):
        """Position window using launcher environment variables."""
        try:
            if 'TOOL_WINDOW_X' in os.environ:
                x = int(os.environ.get('TOOL_WINDOW_X', 100))
                y = int(os.environ.get('TOOL_WINDOW_Y', 100))
                width = int(os.environ.get('TOOL_WINDOW_WIDTH', 1000))
                height = int(os.environ.get('TOOL_WINDOW_HEIGHT', 750))
                
                width = max(width, 800)
                height = max(height, 450)  # Reduced minimum height
                
                self.root.geometry(f"{width}x{height}+{x}+{y}")
                print(f"[INFO] Window positioned at {x},{y} with size {width}x{height}")
            else:
                self.root.geometry("1000x480")
                print("[INFO] Running standalone (not from launcher)")
        except (ValueError, TypeError) as e:
            print(f"[WARNING] Could not position window: {e}")
            self.root.geometry("1000x480")
    
    def setup_ui(self):
        """Setup the main UI."""
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        main_frame = tk.Frame(self.root, bg=UIColors.BG_SECONDARY)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=UISpacing.SM, pady=UISpacing.SM)
        main_frame.grid_rowconfigure(4, weight=1)
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
            text="📝  PDF Text Extractor",
            font=UIFonts.TITLE,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.PRIMARY
        )
        title_label.grid(row=0, column=0, pady=(UISpacing.XS, 0))
        
        desc_label = tk.Label(
            header_frame,
            text="Extract text from PDF files using Python or Azure AI",
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
        
        # Compact drop zone - single line with icon
        main_text = "📄 Drag and drop PDF files here" if DND_AVAILABLE else "📄 Click to select PDF files"
        if DND_AVAILABLE:
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
        options_frame.grid_columnconfigure(3, weight=1)
        
        # Extraction method
        tk.Label(
            options_frame,
            text="Method:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        ).grid(row=0, column=0, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        self.method_var = tk.StringVar(value="python")
        method_frame = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        method_frame.grid(row=0, column=1, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        methods = [
            ("Python (PyMuPDF)", "python", PYMUPDF_AVAILABLE),
            ("OCR (Scanned PDFs)", "ocr", OCRMYPDF_AVAILABLE),
            ("Azure AI", "azure", True)  # Always allow selection, check configuration when processing
        ]
        
        for text, value, available in methods:
            # Azure is always selectable, but may show as not configured
            if value == "azure":
                is_configured = AZURE_CONFIG_AVAILABLE and self.azure_config and self.azure_config.is_doc_intel_configured()
                state = "normal"
                fg_color = UIColors.TEXT_PRIMARY if is_configured else UIColors.TEXT_SECONDARY
            else:
                is_configured = available
                state = "normal" if available else "disabled"
                fg_color = UIColors.TEXT_PRIMARY if available else UIColors.TEXT_MUTED
            
            rb = tk.Radiobutton(
                method_frame,
                text=text,
                variable=self.method_var,
                value=value,
                font=UIFonts.BODY,
                bg=UIColors.BG_PRIMARY,
                fg=fg_color,
                state=state
            )
            rb.pack(side=tk.LEFT, padx=(0, UISpacing.MD))
        
        # Output format
        tk.Label(
            options_frame,
            text="Format:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        ).grid(row=0, column=2, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        self.format_var = tk.StringVar(value="text")
        format_frame = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        format_frame.grid(row=0, column=3, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        for text, value in [("Text (.txt)", "text"), ("Markdown (.md)", "markdown"), ("JSON", "json")]:
            rb = tk.Radiobutton(
                format_frame,
                text=text,
                variable=self.format_var,
                value=value,
                font=UIFonts.BODY,
                bg=UIColors.BG_PRIMARY,
                fg=UIColors.TEXT_PRIMARY
            )
            rb.pack(side=tk.LEFT, padx=(0, UISpacing.MD))
        
        # Force overwrite checkbox
        self.force_var = tk.BooleanVar(value=False)
        force_cb = tk.Checkbutton(
            options_frame,
            text="Overwrite existing files",
            variable=self.force_var,
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        force_cb.grid(row=1, column=0, columnspan=2, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        # Azure config button removed - configuration is now managed in the launcher
    
    def create_results_area(self, parent):
        """Create results text area."""
        result_frame = tk.LabelFrame(
            parent,
            text="  📋 Results  ",
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid"
        )
        result_frame.grid(row=4, column=0, sticky="nsew", pady=UISpacing.XS)
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
            height=4  # Further reduced to make window shorter
        )
        self.result_text.grid(row=0, column=0, sticky="nsew", padx=UISpacing.XS, pady=UISpacing.XS)
    
    def create_button_frame(self, parent):
        """Create button frame."""
        button_frame = tk.Frame(parent, bg=UIColors.BG_SECONDARY)
        button_frame.grid(row=5, column=0, pady=UISpacing.XS)
        
        select_files_btn = create_rounded_button(
            button_frame,
            "📂  Select PDF Files",
            self.select_files,
            style="primary",
            width=18
        )
        select_files_btn.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        select_folder_btn = create_rounded_button(
            button_frame,
            "📁  Select Folder",
            self.select_folder,
            style="secondary"
        )
        select_folder_btn.pack(side=tk.LEFT, padx=UISpacing.SM)
        
        clear_btn = create_rounded_button(
            button_frame,
            "Clear Results",
            self.clear_results,
            style="ghost"
        )
        clear_btn.pack(side=tk.LEFT, padx=UISpacing.SM)
    
    def create_status_bar(self, parent):
        """Create status bar."""
        status_frame = tk.Frame(parent, bg=UIColors.BG_TERTIARY, pady=UISpacing.XS)
        status_frame.grid(row=6, column=0, sticky="ew")
        
        # Status indicators
        status_items = [
            ("PyMuPDF", PYMUPDF_AVAILABLE),
            ("OCR", OCRMYPDF_AVAILABLE),
            ("Azure AI", AZURE_CONFIG_AVAILABLE and self.azure_config and self.azure_config.is_doc_intel_configured()),
            ("Drag&Drop", DND_AVAILABLE)
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
        
        return [f for f in files if f.lower().endswith('.pdf')]
    
    def select_files(self):
        """Open file dialog to select PDFs."""
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if files:
            self.process_files(list(files))
    
    def select_folder(self):
        """Select folder to process."""
        folder = filedialog.askdirectory(title="Select folder with PDF files")
        
        if folder:
            # Ask about recursive processing
            recursive = messagebox.askyesno(
                "Recursive Search",
                "Search subfolders recursively?",
                parent=self.root
            )
            
            files = find_pdf_files(folder, recursive)
            
            if files:
                self.process_files([str(f) for f in files])
            else:
                messagebox.showwarning("No PDFs", "No PDF files found in the selected folder.")
    
    def process_files(self, files):
        """Process a list of PDF files."""
        method = self.method_var.get()
        output_format = self.format_var.get()
        force = self.force_var.get()
        
        # Check if method is available
        if method == 'python' and not PYMUPDF_AVAILABLE:
            messagebox.showerror("Error", "PyMuPDF not available. Install with: pip install pymupdf")
            return
        elif method == 'ocr' and not OCRMYPDF_AVAILABLE:
            messagebox.showerror("Error", "OCRmyPDF not available. Install with: pip install ocrmypdf")
            return
        elif method == 'azure' and (not self.azure_config or not self.azure_config.is_doc_intel_configured()):
            messagebox.showinfo(
                "Azure AI Not Configured",
                "Azure AI is not configured.\n\n"
                "To configure Azure AI:\n"
                "1. Go to the PyPDF Toolbox launcher window\n"
                "2. Click the '⚙️ Azure' button\n"
                "3. Enter your Azure Document Intelligence credentials\n"
                "4. Click 'Save'\n\n"
                "All tools will automatically use the shared configuration."
            )
            return
        
        # Set wait cursor for entire window
        self.root.config(cursor="wait")
        self.root.update()  # Force cursor update
        
        try:
            self.result_text.insert(tk.END, f"\n{'='*60}\n")
            self.result_text.insert(tk.END, f"Processing {len(files)} PDF file(s)\n")
            self.result_text.insert(tk.END, f"Method: {method.upper()} | Format: {output_format.upper()}\n")
            self.result_text.insert(tk.END, f"{'='*60}\n\n")
            auto_scroll_text_widget(self.result_text)
            
            successful = 0
            skipped = 0
            failed = 0
            
            for i, pdf_path in enumerate(files, 1):
                self.result_text.insert(tk.END, f"[{i}/{len(files)}] {os.path.basename(pdf_path)}\n")
                auto_scroll_text_widget(self.result_text)
                
                try:
                    result = process_pdf_file(
                        pdf_path,
                        method=method,
                        output_format=output_format,
                        force=force,
                        azure_config=self.azure_config
                    )
                    
                    if result['success']:
                        if result.get('skipped'):
                            self.result_text.insert(tk.END, f"  ⏭ Skipped (output exists)\n")
                            skipped += 1
                        else:
                            self.result_text.insert(tk.END, f"  ✓ Saved: {result['output']}\n")
                            successful += 1
                    else:
                        self.result_text.insert(tk.END, f"  ✗ Error: {result.get('error', 'Unknown')}\n")
                        failed += 1
                        
                except Exception as e:
                    self.result_text.insert(tk.END, f"  ✗ Error: {str(e)}\n")
                    failed += 1
                
                auto_scroll_text_widget(self.result_text)
            
            # Summary
            self.result_text.insert(tk.END, f"\n{'='*60}\n")
            self.result_text.insert(tk.END, f"SUMMARY\n")
            self.result_text.insert(tk.END, f"{'='*60}\n")
            self.result_text.insert(tk.END, f"Total: {len(files)} | ")
            self.result_text.insert(tk.END, f"✓ Success: {successful} | ")
            self.result_text.insert(tk.END, f"⏭ Skipped: {skipped} | ")
            self.result_text.insert(tk.END, f"✗ Failed: {failed}\n")
            self.result_text.insert(tk.END, f"{'='*60}\n\n")
            auto_scroll_text_widget(self.result_text)
            
            # Show summary dialog
            if failed == 0:
                messagebox.showinfo(
                    "Complete",
                    f"Processed {len(files)} file(s).\n\n"
                    f"✓ Success: {successful}\n"
                    f"⏭ Skipped: {skipped}"
                )
            else:
                messagebox.showwarning(
                    "Complete with Errors",
                    f"Processed {len(files)} file(s).\n\n"
                    f"✓ Success: {successful}\n"
                    f"⏭ Skipped: {skipped}\n"
                    f"✗ Failed: {failed}"
                )
        finally:
            # Always restore cursor
            self.root.config(cursor="")
            self.root.update()
    
    def clear_results(self):
        """Clear results text."""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Results cleared.\n\n")
    
    def show_welcome_message(self):
        """Show welcome message."""
        self.result_text.insert(tk.END, "Welcome to PDF Text Extractor!\n")
        self.result_text.insert(tk.END, "━" * 45 + "\n")
        
        self.result_text.insert(tk.END, "📋 Methods: ")
        methods = []
        if PYMUPDF_AVAILABLE:
            methods.append("✓ Python")
        else:
            methods.append("✗ Python")
        if OCRMYPDF_AVAILABLE:
            methods.append("✓ OCR")
        else:
            methods.append("✗ OCR")
        if AZURE_CONFIG_AVAILABLE and self.azure_config and self.azure_config.is_doc_intel_configured():
            methods.append("✓ Azure AI")
        else:
            methods.append("✗ Azure AI (configure in launcher)")
        self.result_text.insert(tk.END, " | ".join(methods) + "\n")
        
        self.result_text.insert(tk.END, "📄 Formats: Text | Markdown | JSON\n")
        
        if DND_AVAILABLE:
            self.result_text.insert(tk.END, "💡 Drag and drop PDF files to begin.\n")
        else:
            self.result_text.insert(tk.END, "💡 Click 'Select PDF Files' to begin.\n")
    
    
    def update_status_bar(self):
        """Update the status bar."""
        status_items = [
            ("PyMuPDF", PYMUPDF_AVAILABLE),
            ("OCR", OCRMYPDF_AVAILABLE),
            ("Azure AI", AZURE_CONFIG_AVAILABLE and self.azure_config and self.azure_config.is_doc_intel_configured()),
            ("Drag&Drop", DND_AVAILABLE)
        ]
        
        status_parts = []
        for name, available in status_items:
            icon = "✓" if available else "✗"
            status_parts.append(f"{icon} {name}")
        
        status_text = "  •  ".join(status_parts)
        self.status_label.config(text=status_text)
    
    def run(self):
        """Run the application."""
        self.root.mainloop()


# ============================================================================
# Command Line Interface
# ============================================================================

def main():
    """Main entry point."""
    print("[INFO] PDF Text Extractor initialized")
    logger.info("Starting PDF Text Extractor")
    
    # GUI mode
    if args.gui or not args.input_path:
        print("[INFO] Starting GUI mode...")
        app = PDFTextExtractorApp()
        app.run()
        return
    
    # Command line mode
    input_path = args.input_path
    
    if not os.path.exists(input_path):
        print(f"❌ Path not found: {input_path}")
        sys.exit(1)
    
    # Find PDF files
    files = find_pdf_files(input_path, args.recursive)
    
    if not files:
        print(f"❌ No PDF files found in: {input_path}")
        sys.exit(1)
    
    print(f"📄 Found {len(files)} PDF file(s)")
    print(f"📋 Method: {args.method.upper()} | Format: {args.format.upper()}")
    print()
    
    # Load Azure config if needed
    azure_config = None
    if args.method == 'azure':
        if AZURE_CONFIG_AVAILABLE:
            azure_config = get_azure_config()
            if not azure_config.is_doc_intel_configured():
                print("❌ Azure Document Intelligence not configured")
                sys.exit(1)
        else:
            print("❌ Azure configuration not available")
            sys.exit(1)
    
    # Process files
    successful = 0
    failed = 0
    skipped = 0
    
    for i, pdf_file in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {pdf_file}")
        
        result = process_pdf_file(
            str(pdf_file),
            method=args.method,
            output_format=args.format,
            output_dir=args.output_dir,
            force=args.force,
            azure_config=azure_config
        )
        
        if result['success']:
            if result.get('skipped'):
                print(f"  ⏭ Skipped (output exists)")
                skipped += 1
            else:
                print(f"  ✓ Saved: {result['output']}")
                successful += 1
        else:
            print(f"  ✗ Error: {result.get('error', 'Unknown')}")
            failed += 1
    
    print()
    print(f"{'='*60}")
    print(f"SUMMARY: Total={len(files)} | Success={successful} | Skipped={skipped} | Failed={failed}")
    print(f"{'='*60}")
    
    if failed > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
