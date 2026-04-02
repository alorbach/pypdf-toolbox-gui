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
from typing import List, Optional, Tuple

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
    import pytesseract
    OCRMYPDF_AVAILABLE = True
except ImportError:
    ocrmypdf = None  # type: ignore
    pytesseract = None  # type: ignore
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

# PyMuPDF – used to write invisible text overlay for AI OCR
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    fitz = None  # type: ignore
    PYMUPDF_AVAILABLE = False
    print("[WARNING] PyMuPDF (fitz) not available. Install with: pip install pymupdf")

# HTTP requests – used for Azure Document Intelligence API calls
try:
    import requests as _requests_module
    import time as _time_module
    REQUESTS_AVAILABLE = True
except ImportError:
    _requests_module = None  # type: ignore
    _time_module = None  # type: ignore
    REQUESTS_AVAILABLE = False
    print("[WARNING] requests not available. Install with: pip install requests")

# Azure AI configuration
try:
    import sys as _sys_ref
    import os as _os_ref
    # Resolve src/utils relative to this file
    _utils_path = str(_os_ref.path.join(_os_ref.path.dirname(_os_ref.path.abspath(__file__)), "utils"))
    if _utils_path not in _sys_ref.path:
        _sys_ref.path.insert(0, _os_ref.path.dirname(_os_ref.path.abspath(__file__)))
    from utils.azure_config import AzureAIConfig, get_azure_config
    AZURE_CONFIG_AVAILABLE = True
except Exception as _azure_import_err:
    AzureAIConfig = None  # type: ignore
    get_azure_config = None  # type: ignore
    AZURE_CONFIG_AVAILABLE = False
    print(f"[WARNING] Azure config not available: {_azure_import_err}")


# ============================================================================
# Tesseract Availability Check
# ============================================================================

# Path to tesseract.exe when found in common install locations (not in PATH)
TESSERACT_PATH: Optional[str] = None


def _get_common_tesseract_paths() -> List[str]:
    """Return common Windows install paths for Tesseract OCR."""
    if sys.platform != "win32":
        return []
    drive = os.environ.get("SystemDrive", "C:")
    return [
        os.path.join(drive, r"Program Files\Tesseract-OCR\tesseract.exe"),
        os.path.join(drive, r"Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]


def find_tesseract_path() -> Optional[str]:
    """Find Tesseract executable - check PATH first, then common Windows paths.
    
    Returns:
        Full path to tesseract.exe if found, None otherwise.
    """
    # Try PATH first
    try:
        result = subprocess.run(
            ["tesseract", "--version"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return "tesseract"  # In PATH
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    # Try common Windows install paths (handles app running before PATH was updated)
    for path in _get_common_tesseract_paths():
        if os.path.isfile(path):
            return path
    return None


def check_tesseract_available() -> Tuple[bool, Optional[str]]:
    """Check if Tesseract OCR is installed (PATH or common Windows paths).
    
    Returns:
        Tuple of (is_available, path_or_none). path is set when found in common location.
    """
    global TESSERACT_PATH
    path = find_tesseract_path()
    if path is None:
        TESSERACT_PATH = None
        return False, None
    if path != "tesseract":
        TESSERACT_PATH = path
    else:
        TESSERACT_PATH = None  # In PATH, no need to set
    return True, path


def check_ghostscript_available() -> bool:
    """Check if Ghostscript (gswin64c) is installed - used for optional image optimization."""
    cmd = ["gswin64c", "--version"] if sys.platform == "win32" else ["gs", "--version"]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        return False


def get_tesseract_install_instructions() -> str:
    """Get platform-specific Tesseract installation instructions."""
    if sys.platform == "win32":
        return (
            "Tesseract OCR must be installed separately on Windows.\n\n"
            "Install using one of these methods:\n\n"
            "1. winget (Windows 10/11, recommended):\n"
            "   winget install --id UB-Mannheim.TesseractOCR -e\n\n"
            "2. Chocolatey:\n"
            "   choco install tesseract\n\n"
            "3. Manual download:\n"
            "   https://github.com/UB-Mannheim/tesseract/wiki\n"
            "   Download the installer, run it, and add the install folder to PATH.\n\n"
            "After installing, restart this application and try again."
        )
    elif sys.platform == "darwin":
        return "Install Tesseract: brew install tesseract"
    else:
        return "Install Tesseract: sudo apt install tesseract-ocr  (Debian/Ubuntu)"


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


def process_pdf_with_ocr(
    pdf_path: str,
    language: str = 'eng',
    output_callback=None,
    optimize_level: int = 0,
) -> bool:
    """Process a single PDF file with OCR.
    
    Args:
        pdf_path: Path to PDF file
        language: OCR language code
        output_callback: Optional callback function(line) to receive output lines
        optimize_level: 0=no optimization (no Ghostscript), 1=lossless (requires Ghostscript)
    
    Returns:
        True if successful, False otherwise
    """
    if not OCRMYPDF_AVAILABLE:
        return False

    # Configure pytesseract to use Tesseract when found in common path (not in PATH)
    if TESSERACT_PATH and pytesseract:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

    try:
        # In frozen (exe) mode, avoid spawning sys.executable which relaunches the app.
        if getattr(sys, 'frozen', False):
            try:
                ocrmypdf.ocr(
                    pdf_path,
                    pdf_path,
                    skip_text=True,
                    output_type='pdf',
                    language=language,
                    optimize=optimize_level,
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
            "--optimize", str(optimize_level),
            "--language", language,
            pdf_path,
            pdf_path  # Output to same file
        ]

        # When Tesseract is in common path but not in PATH, add it to subprocess env
        env = None
        if TESSERACT_PATH and TESSERACT_PATH != "tesseract":
            env = os.environ.copy()
            tesseract_dir = os.path.dirname(TESSERACT_PATH)
            env["PATH"] = tesseract_dir + os.pathsep + env.get("PATH", "")

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            env=env,
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
                language=language,
                optimize=optimize_level,
            )
            return True
        except Exception as e2:
            if output_callback:
                output_callback(f"Error: {str(e2)}")
            print(f"Error processing {pdf_path}: {error_msg}")
            return False


def process_directory_images(
    directory: str,
    language: str = 'eng',
    output_callback=None,
    optimize_level: int = 0,
    use_ai: bool = False,
    azure_config=None,
) -> bool:
    """Process all images in a directory and convert them to a single PDF with OCR.
    
    Args:
        directory: Directory containing image files
        language: OCR language code (used for local Tesseract OCR only)
        output_callback: Optional callback function(line) to receive output lines
        optimize_level: 0=no optimization, 1=lossless (requires Ghostscript)
        use_ai: If True, use Azure Document Intelligence instead of local Tesseract
        azure_config: AzureAIConfig instance for AI OCR (loaded automatically if None)
    
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
    if use_ai:
        if not process_pdf_with_ai_ocr(pdf_path, output_callback=output_callback, azure_config=azure_config):
            return False
    else:
        if not process_pdf_with_ocr(
            pdf_path, language, output_callback=output_callback, optimize_level=optimize_level
        ):
            return False
    
    return True


def combine_all_images_to_pdf(
    image_files: List[str],
    output_pdf: str,
    output_callback=None,
) -> bool:
    """Combine a flat list of image files (any directories) into a single PDF.

    Args:
        image_files: List of image file paths to merge (will be sorted).
        output_pdf: Destination PDF path.
        output_callback: Optional callback(line) for progress messages.

    Returns:
        True if successful, False otherwise.
    """
    if not image_files:
        return False

    sorted_images = sorted(image_files)

    if output_callback:
        output_callback(f"Combining {len(sorted_images)} image(s) into: {os.path.basename(output_pdf)}")

    if not convert_images_to_pdf(sorted_images, output_pdf):
        if output_callback:
            output_callback("Error: Failed to convert images to PDF.")
        return False

    if output_callback:
        output_callback(f"✓ Combined PDF created: {os.path.basename(output_pdf)}")

    return True


def process_pdf_with_ai_ocr(
    pdf_path: str,
    output_callback=None,
    azure_config=None,
) -> bool:
    """Process a PDF file with Azure Document Intelligence OCR.

    Submits the PDF to Azure Document Intelligence (prebuilt-read model),
    retrieves word-level bounding boxes from the response, and writes an
    invisible text overlay onto the original PDF using PyMuPDF so the file
    becomes searchable without visually altering it.

    Args:
        pdf_path: Path to the PDF file (modified in-place).
        output_callback: Optional callback function(line) for progress output.
        azure_config: AzureAIConfig instance; loaded automatically if None.

    Returns:
        True if successful, False otherwise.
    """
    if not PYMUPDF_AVAILABLE:
        if output_callback:
            output_callback("Error: PyMuPDF (fitz) is required for AI OCR. Install with: pip install pymupdf")
        return False

    if not REQUESTS_AVAILABLE:
        if output_callback:
            output_callback("Error: requests library is required for AI OCR. Install with: pip install requests")
        return False

    # Load config if not supplied
    if azure_config is None:
        if not AZURE_CONFIG_AVAILABLE:
            if output_callback:
                output_callback("Error: Azure configuration module not available.")
            return False
        try:
            azure_config = get_azure_config()
        except Exception as e:
            if output_callback:
                output_callback(f"Error loading Azure config: {e}")
            return False

    if not azure_config.is_doc_intel_configured():
        if output_callback:
            output_callback("Error: Azure Document Intelligence is not configured. Please set endpoint and API key.")
        return False

    import requests as req
    import time

    endpoint = azure_config.doc_intel_endpoint.rstrip('/')
    api_key = azure_config.doc_intel_api_key
    request_timeout = azure_config.timeout
    polling_timeout = azure_config.polling_timeout

    analyze_url = (
        f"{endpoint}/formrecognizer/documentModels/prebuilt-read:analyze"
        f"?api-version=2023-07-31"
    )
    headers = {
        'Content-Type': 'application/pdf',
        'Ocp-Apim-Subscription-Key': api_key,
    }

    if output_callback:
        output_callback(f"Submitting to Azure Document Intelligence: {os.path.basename(pdf_path)}")

    try:
        with open(pdf_path, 'rb') as f:
            response = req.post(analyze_url, headers=headers, data=f, timeout=request_timeout)
        response.raise_for_status()
    except Exception as e:
        if output_callback:
            output_callback(f"Error submitting PDF: {e}")
        return False

    operation_url = response.headers.get("Operation-Location")
    if not operation_url:
        if output_callback:
            output_callback("Error: No Operation-Location header in Azure response.")
        return False

    # Poll for results
    status_headers = {'Ocp-Apim-Subscription-Key': api_key}
    start_time = time.time()
    max_attempts = polling_timeout // 3
    result_json = None

    for attempt in range(int(max_attempts)):
        elapsed = time.time() - start_time
        if elapsed > polling_timeout:
            if output_callback:
                output_callback(f"Error: Azure AI processing timed out after {elapsed:.0f}s")
            return False
        try:
            poll_resp = req.get(operation_url, headers=status_headers, timeout=30)
            poll_resp.raise_for_status()
            result_json = poll_resp.json()
            status = result_json.get('status', '')

            if attempt % 10 == 0 or status != 'running':
                if output_callback:
                    output_callback(f"  Azure status: {status} (elapsed: {elapsed:.0f}s)")

            if status == 'succeeded':
                break
            elif status == 'failed':
                error = result_json.get('error', {})
                if output_callback:
                    output_callback(f"Error: Azure AI processing failed: {error.get('message', 'Unknown')}")
                return False
        except req.exceptions.Timeout:
            if output_callback:
                output_callback(f"  Poll attempt {attempt + 1} timed out, retrying...")
        except Exception as e:
            if output_callback:
                output_callback(f"  Poll attempt {attempt + 1} error: {e}, retrying...")
            time.sleep(2)

        time.sleep(3)
    else:
        if output_callback:
            output_callback(f"Error: Azure AI processing timed out.")
        return False

    # --- Write invisible text overlay via PyMuPDF ---
    analyze_result = result_json.get('analyzeResult', {})
    di_pages = analyze_result.get('pages', [])

    if not di_pages:
        if output_callback:
            output_callback("Warning: No pages returned by Azure AI.")
        return True  # Nothing to write but not a hard failure

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        if output_callback:
            output_callback(f"Error opening PDF with PyMuPDF: {e}")
        return False

    words_written = 0

    for page_idx, di_page in enumerate(di_pages):
        if page_idx >= len(doc):
            break

        fitz_page = doc[page_idx]
        fitz_rect = fitz_page.rect  # PDF points (width x height)

        # Azure DI returns page dimensions in the unit specified (usually 'inch' or 'pixel')
        # The polygon values are in the same unit as width/height.
        di_width = di_page.get('width', 0)
        di_height = di_page.get('height', 0)

        if di_width <= 0 or di_height <= 0:
            continue

        # Scale factors from DI coordinates to PDF points
        scale_x = fitz_rect.width / di_width
        scale_y = fitz_rect.height / di_height

        words = di_page.get('words', [])

        for word in words:
            content = word.get('content', '').strip()
            if not content:
                continue

            polygon = word.get('polygon', [])
            if len(polygon) < 8:
                continue

            # polygon: [x0,y0, x1,y1, x2,y2, x3,y3] (top-left clockwise)
            xs = [polygon[i] * scale_x for i in range(0, 8, 2)]
            ys = [polygon[i] * scale_y for i in range(1, 8, 2)]

            x0, y0 = min(xs), min(ys)
            x1, y1 = max(xs), max(ys)

            word_height = y1 - y0
            if word_height <= 0:
                word_height = 10  # fallback

            # Estimate a font size that roughly fits the bounding box
            font_size = max(1, word_height * 0.85)

            # Insert invisible text (render_mode 3 = invisible, searchable)
            try:
                fitz_page.insert_text(
                    fitz.Point(x0, y1),   # baseline at bottom-left of bbox
                    content,
                    fontsize=font_size,
                    render_mode=3,
                    overlay=True,
                )
                words_written += 1
            except Exception:
                pass  # Skip words that can't be placed

    if output_callback:
        output_callback(f"  Wrote {words_written} words as invisible text overlay.")

    try:
        doc.save(pdf_path, incremental=False, deflate=True)
        doc.close()
    except Exception as e:
        try:
            doc.close()
        except Exception:
            pass
        if output_callback:
            output_callback(f"Error saving PDF: {e}")
        return False

    if output_callback:
        output_callback(f"✓ AI OCR complete: {os.path.basename(pdf_path)}")

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
        
        # Check Tesseract availability (required for local OCR)
        self.tesseract_available, _ = check_tesseract_available()
        self.ghostscript_available = check_ghostscript_available()

        # Load Azure AI configuration (for AI OCR engine)
        self.azure_config = None
        self._azure_di_configured_flag = False
        if AZURE_CONFIG_AVAILABLE:
            try:
                self.azure_config = get_azure_config()
                self._azure_di_configured_flag = self.azure_config.is_doc_intel_configured()
            except Exception as e:
                print(f"[WARNING] Could not load Azure AI config: {e}")

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

        # ── Row 0: OCR Engine ─────────────────────────────────────────────
        tk.Label(
            options_frame,
            text="OCR Engine:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        ).grid(row=0, column=0, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")

        self.ocr_engine_var = tk.StringVar(value="local")
        engine_frame = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        engine_frame.grid(row=0, column=1, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")

        tk.Radiobutton(
            engine_frame,
            text="Local (Tesseract)",
            variable=self.ocr_engine_var,
            value="local",
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            command=self._on_engine_change,
        ).pack(side=tk.LEFT, padx=(0, UISpacing.MD))

        tk.Radiobutton(
            engine_frame,
            text="AI (Azure Document Intelligence)",
            variable=self.ocr_engine_var,
            value="azure",
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            command=self._on_engine_change,
        ).pack(side=tk.LEFT, padx=(0, UISpacing.SM))

        # Azure DI config status label (shown next to AI radio)
        di_status_text, di_status_color = self._azure_di_status_text()
        self.azure_di_status_label = tk.Label(
            engine_frame,
            text=di_status_text,
            font=UIFonts.SMALL_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=di_status_color,
            cursor="hand2",
        )
        self.azure_di_status_label.pack(side=tk.LEFT)
        self.azure_di_status_label.bind("<Button-1>", lambda e: self._on_azure_configure_click())

        # ── Row 1: Language (local OCR only) ─────────────────────────────
        self.language_label = tk.Label(
            options_frame,
            text="OCR Language:",
            font=UIFonts.BODY_BOLD,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY
        )
        self.language_label.grid(row=1, column=0, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        self.language_var = tk.StringVar(value="eng")
        self.language_frame = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        self.language_frame.grid(row=1, column=1, padx=UISpacing.SM, pady=UISpacing.SM, sticky="w")
        
        languages = [
            ("English", "eng"),
            ("German", "deu"),
            ("English + German", "eng+deu"),
            ("French", "fra"),
            ("Spanish", "spa"),
        ]
        
        self._language_radiobuttons = []
        for text, value in languages:
            rb = tk.Radiobutton(
                self.language_frame,
                text=text,
                variable=self.language_var,
                value=value,
                font=UIFonts.BODY,
                bg=UIColors.BG_PRIMARY,
                fg=UIColors.TEXT_PRIMARY,
            )
            rb.pack(side=tk.LEFT, padx=(0, UISpacing.MD))
            self._language_radiobuttons.append(rb)

        # Note shown when AI OCR selected (language auto-detected)
        self.ai_language_note = tk.Label(
            self.language_frame,
            text="(Azure AI auto-detects language)",
            font=UIFonts.SMALL,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_MUTED,
        )

        # ── Row 2: Ghostscript (local OCR only) ──────────────────────────
        self.gs_row = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        self.gs_row.grid(row=2, column=0, columnspan=2, padx=UISpacing.SM, pady=(0, UISpacing.SM), sticky="w")
        self._build_gs_row()

        # ── Row 3: Combine images into single PDF ─────────────────────────
        self.single_pdf_var = tk.BooleanVar(value=False)
        single_pdf_frame = tk.Frame(options_frame, bg=UIColors.BG_PRIMARY)
        single_pdf_frame.grid(row=3, column=0, columnspan=2, padx=UISpacing.SM, pady=(0, UISpacing.SM), sticky="w")

        tk.Checkbutton(
            single_pdf_frame,
            text="Combine selected images into a single PDF (asks for output path)",
            variable=self.single_pdf_var,
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            activebackground=UIColors.BG_PRIMARY,
        ).pack(side=tk.LEFT)

    def _azure_di_status_text(self):
        """Return (label_text, color) for Azure DI status."""
        if not AZURE_CONFIG_AVAILABLE:
            return "✗ Azure config not available", UIColors.ERROR
        if self._azure_di_configured_flag:
            return "✓ Configured", UIColors.SUCCESS
        return "✗ Not configured — click to configure", UIColors.ERROR

    def _on_engine_change(self):
        """Update UI visibility when OCR engine selection changes."""
        using_ai = self.ocr_engine_var.get() == "azure"

        # Grey out language controls when AI selected
        state = "disabled" if using_ai else "normal"
        for rb in self._language_radiobuttons:
            rb.config(state=state)

        if using_ai:
            # Hide language radio buttons, show note
            for rb in self._language_radiobuttons:
                rb.pack_forget()
            self.ai_language_note.pack(side=tk.LEFT)
        else:
            self.ai_language_note.pack_forget()
            for rb in self._language_radiobuttons:
                rb.pack(side=tk.LEFT, padx=(0, UISpacing.MD))

        # Show/hide GS row (only relevant for local OCR)
        if using_ai:
            self.gs_row.grid_remove()
        else:
            self.gs_row.grid()

    def _on_azure_configure_click(self):
        """Prompt the user to configure Azure Document Intelligence."""
        messagebox.showinfo(
            "Configure Azure AI",
            "Please configure Azure Document Intelligence in the launcher:\n\n"
            "1. Open the main launcher window\n"
            "2. Click the '⚙️ Azure' button\n"
            "3. Enter your Azure Document Intelligence endpoint and API key\n"
            "4. Click Save, then restart this tool.\n\n"
            "Alternatively, set environment variables:\n"
            "  AZURE_DOC_INTEL_ENDPOINT\n"
            "  AZURE_DOC_INTEL_API_KEY",
            parent=self.root,
        )

    def _build_gs_row(self):
        """Build or rebuild the Ghostscript options row."""
        for w in self.gs_row.winfo_children():
            w.destroy()
        tk.Label(
            self.gs_row,
            text="Optional: Ghostscript for image optimization",
            font=UIFonts.SMALL,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_SECONDARY,
        ).pack(side=tk.LEFT)
        if not self.ghostscript_available and sys.platform == "win32":
            create_rounded_button(
                self.gs_row, "Install Ghostscript", self._run_ghostscript_install, style="ghost"
            ).pack(side=tk.LEFT, padx=(UISpacing.MD, 0))
            create_rounded_button(
                self.gs_row, "Check again", self._check_ghostscript_again, style="ghost"
            ).pack(side=tk.LEFT, padx=(UISpacing.SM, 0))
        elif not self.ghostscript_available:
            tk.Label(
                self.gs_row,
                text="(install gs for optimization)",
                font=UIFonts.SMALL,
                bg=UIColors.BG_PRIMARY,
                fg=UIColors.TEXT_MUTED,
            ).pack(side=tk.LEFT, padx=(UISpacing.SM, 0))

    def _check_ghostscript_again(self):
        """Re-check Ghostscript availability after user installed it."""
        self.ghostscript_available = check_ghostscript_available()
        self._refresh_status_bar()
        if self.ghostscript_available:
            self._build_gs_row()
            messagebox.showinfo(
                "Ghostscript Found",
                "Ghostscript was detected. Image optimization is now enabled.",
                parent=self.root,
            )

    def _run_ghostscript_install(self):
        """Run winget to install Ghostscript in a new console window (Windows)."""
        if sys.platform != "win32":
            return
        try:
            subprocess.Popen(
                ["cmd", "/k", "winget", "install", "-e", "--id", "ArtifexSoftware.GhostScript"],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
            messagebox.showinfo(
                "Installation Started",
                "A command window has opened to install Ghostscript.\n\n"
                "Follow the prompts, then click 'Check again' when done.",
                parent=self.root,
            )
        except Exception as e:
            messagebox.showerror(
                "Could Not Start Install",
                f"Failed to run winget: {e}\n\n"
                "Please run manually in Command Prompt:\n"
                "winget install -e --id ArtifexSoftware.GhostScript",
                parent=self.root,
            )

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
        
        self._status_items = [
            ("Tesseract", lambda: self.tesseract_available),
            ("Ghostscript", lambda: self.ghostscript_available),
            ("OCRmyPDF", lambda: OCRMYPDF_AVAILABLE),
            ("Pillow", lambda: PIL_AVAILABLE),
            ("img2pdf", lambda: IMG2PDF_AVAILABLE),
            ("PyMuPDF", lambda: PYMUPDF_AVAILABLE),
            ("Azure DI", lambda: self._azure_di_configured_flag),
            ("Drag&Drop", lambda: HAS_DND),
        ]
        status_parts = []
        for name, get_available in self._status_items:
            icon = "✓" if get_available() else "✗"
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

    def _refresh_status_bar(self):
        """Update status bar text (e.g. after Tesseract is detected)."""
        status_parts = []
        for name, get_available in self._status_items:
            icon = "✓" if get_available() else "✗"
            status_parts.append(f"{icon} {name}")
        self.status_label.config(text="  •  ".join(status_parts))

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

        ocr_engine = self.ocr_engine_var.get()  # "local" or "azure"
        use_ai = ocr_engine == "azure"

        # ── Preflight checks ──────────────────────────────────────────────
        if use_ai:
            # AI engine: validate Azure DI configuration
            if not PYMUPDF_AVAILABLE:
                messagebox.showerror(
                    "Missing Dependency",
                    "AI OCR requires PyMuPDF.\n\nInstall with:\n  pip install pymupdf"
                )
                return
            if not REQUESTS_AVAILABLE:
                messagebox.showerror(
                    "Missing Dependency",
                    "AI OCR requires the requests library.\n\nInstall with:\n  pip install requests"
                )
                return
            if not self._azure_di_configured_flag:
                messagebox.showerror(
                    "Azure Not Configured",
                    "Azure Document Intelligence is not configured.\n\n"
                    "Please configure it in the launcher (⚙️ Azure button) and restart this tool.",
                    parent=self.root,
                )
                return
        else:
            # Local engine: validate OCRmyPDF + Tesseract
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
            if not self.tesseract_available:
                self._show_tesseract_install_dialog()
                return

        # ── Image-dependency check ────────────────────────────────────────
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
            engine_label = "Azure Document Intelligence" if use_ai else f"Local Tesseract ({self.language})"
            self.result_text.insert(tk.END, f"\n{'='*60}\n")
            self.result_text.insert(tk.END, f"Processing {len(files)} file(s)\n")
            self.result_text.insert(tk.END, f"Engine: {engine_label}\n")
            self.result_text.insert(tk.END, f"{'='*60}\n\n")
            self.result_text.see(tk.END)
            self.root.update()

            # Separate PDFs and images
            pdf_files = [f for f in files if f.lower().endswith('.pdf')]
            image_files = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.bmp'))]

            processed_count = 0
            skipped_count = 0
            error_count = 0

            # ── Helper: per-file log callback ─────────────────────────────
            def make_log_callback():
                def log_output(line):
                    self.result_text.insert(tk.END, f"  {line}\n")
                    self.result_text.see(tk.END)
                    self.root.update()
                return log_output

            # ── Process PDFs ──────────────────────────────────────────────
            for i, pdf_file in enumerate(pdf_files, 1):
                self.result_text.insert(
                    tk.END, f"[{i}/{len(pdf_files)}] Processing PDF: {os.path.basename(pdf_file)}\n"
                )
                self.result_text.see(tk.END)
                self.root.update()

                try:
                    log_cb = make_log_callback()
                    if use_ai:
                        success = process_pdf_with_ai_ocr(
                            pdf_file,
                            output_callback=log_cb,
                            azure_config=self.azure_config,
                        )
                    else:
                        optimize_lvl = 1 if self.ghostscript_available else 0
                        success = process_pdf_with_ocr(
                            pdf_file, self.language, output_callback=log_cb, optimize_level=optimize_lvl
                        )

                    if success:
                        self.result_text.insert(tk.END, "  ✓ Successfully processed\n")
                        processed_count += 1
                    else:
                        self.result_text.insert(tk.END, "  ⏭ Skipped (may already have OCR text)\n")
                        skipped_count += 1
                except Exception as e:
                    self.result_text.insert(tk.END, f"  ✗ Error: {str(e)}\n")
                    error_count += 1

                self.result_text.see(tk.END)
                self.root.update()

            # ── Process images ────────────────────────────────────────────
            if image_files:
                combine_single = self.single_pdf_var.get()

                if combine_single:
                    # Auto-save combined PDF in the folder of the first image
                    first_img = sorted(image_files)[0]
                    default_name = os.path.splitext(os.path.basename(first_img))[0] + "_combined.pdf"
                    output_pdf = os.path.join(os.path.dirname(first_img), default_name)
                    if not output_pdf:
                        self.result_text.insert(tk.END, "\nImage processing cancelled.\n")
                        self.result_text.see(tk.END)
                    else:
                        self.result_text.insert(
                            tk.END, f"\nCombining {len(image_files)} image(s) into single PDF…\n"
                        )
                        self.result_text.insert(tk.END, f"  Output: {output_pdf}\n")
                        self.result_text.see(tk.END)
                        self.root.update()

                        log_cb = make_log_callback()
                        try:
                            if combine_all_images_to_pdf(image_files, output_pdf, output_callback=log_cb):
                                # Now OCR the combined PDF
                                self.result_text.insert(tk.END, "  Running OCR on combined PDF…\n")
                                self.result_text.see(tk.END)
                                self.root.update()
                                if use_ai:
                                    ocr_ok = process_pdf_with_ai_ocr(
                                        output_pdf, output_callback=log_cb, azure_config=self.azure_config
                                    )
                                else:
                                    optimize_lvl = 1 if self.ghostscript_available else 0
                                    ocr_ok = process_pdf_with_ocr(
                                        output_pdf, self.language, output_callback=log_cb,
                                        optimize_level=optimize_lvl
                                    )
                                if ocr_ok:
                                    self.result_text.insert(
                                        tk.END, f"  ✓ Combined searchable PDF: {os.path.basename(output_pdf)}\n"
                                    )
                                    processed_count += 1
                                else:
                                    self.result_text.insert(tk.END, "  ⏭ OCR skipped (already has text?)\n")
                                    skipped_count += 1
                            else:
                                self.result_text.insert(tk.END, "  ✗ Failed to combine images.\n")
                                error_count += 1
                        except Exception as e:
                            self.result_text.insert(tk.END, f"  ✗ Error: {str(e)}\n")
                            error_count += 1

                        self.result_text.see(tk.END)
                        self.root.update()
                else:
                    # Per-image processing: each image → its own PDF with the same basename
                    for i, img in enumerate(image_files, 1):
                        img_dir = os.path.dirname(img)
                        basename = os.path.splitext(os.path.basename(img))[0]
                        output_pdf = os.path.join(img_dir, f"{basename}.pdf")

                        self.result_text.insert(
                            tk.END, f"\n[{i}/{len(image_files)}] Converting: {os.path.basename(img)} → {basename}.pdf\n"
                        )
                        self.result_text.see(tk.END)
                        self.root.update()

                        log_cb = make_log_callback()
                        try:
                            # Convert single image to PDF
                            if not convert_images_to_pdf([img], output_pdf):
                                self.result_text.insert(tk.END, "  ✗ Failed to convert image to PDF.\n")
                                error_count += 1
                                continue

                            # Run OCR on the resulting PDF
                            if use_ai:
                                ocr_ok = process_pdf_with_ai_ocr(
                                    output_pdf, output_callback=log_cb, azure_config=self.azure_config
                                )
                            else:
                                optimize_lvl = 1 if self.ghostscript_available else 0
                                ocr_ok = process_pdf_with_ocr(
                                    output_pdf, self.language, output_callback=log_cb,
                                    optimize_level=optimize_lvl
                                )

                            if ocr_ok:
                                self.result_text.insert(tk.END, f"  ✓ Saved: {output_pdf}\n")
                                processed_count += 1
                            else:
                                self.result_text.insert(tk.END, "  ⏭ OCR skipped (already has text?).\n")
                                skipped_count += 1
                        except Exception as e:
                            self.result_text.insert(tk.END, f"  ✗ Error: {str(e)}\n")
                            error_count += 1

                        self.result_text.see(tk.END)
                        self.root.update()

            # ── Summary ───────────────────────────────────────────────────
            self.result_text.insert(tk.END, f"\n{'='*60}\n")
            self.result_text.insert(tk.END, "SUMMARY\n")
            self.result_text.insert(tk.END, f"{'='*60}\n")
            self.result_text.insert(tk.END, f"Total: {len(files)} | ")
            self.result_text.insert(tk.END, f"✓ Success: {processed_count} | ")
            self.result_text.insert(tk.END, f"⏭ Skipped: {skipped_count} | ")
            self.result_text.insert(tk.END, f"✗ Failed: {error_count}\n")
            self.result_text.insert(tk.END, f"{'='*60}\n\n")
            self.result_text.see(tk.END)

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
    
    def _run_winget_install(self):
        """Run winget to install Tesseract in a new console window (Windows)."""
        if sys.platform != "win32":
            return
        try:
            subprocess.Popen(
                ["cmd", "/k", "winget", "install", "--id", "UB-Mannheim.TesseractOCR", "-e"],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
            messagebox.showinfo(
                "Installation Started",
                "A command window has opened to install Tesseract.\n\n"
                "Follow the prompts, then restart this application when done.",
                parent=self.root,
            )
        except Exception as e:
            messagebox.showerror(
                "Could Not Start Install",
                f"Failed to run winget: {e}\n\n"
                "Please run manually in Command Prompt:\n"
                "winget install --id UB-Mannheim.TesseractOCR -e",
                parent=self.root,
            )

    def _show_tesseract_install_dialog(self):
        """Show dialog with Tesseract install options and optional auto-run button."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Tesseract OCR Not Found")
        dialog.configure(bg=UIColors.BG_SECONDARY)
        dialog.transient(self.root)
        dialog.grab_set()

        width, height = 540, 420
        dialog.geometry(f"{width}x{height}")
        dialog.resizable(True, True)

        # Content
        msg_frame = tk.Frame(dialog, bg=UIColors.BG_SECONDARY, padx=UISpacing.LG, pady=UISpacing.LG)
        msg_frame.pack(fill="both", expand=True)

        tk.Label(
            msg_frame,
            text="Tesseract OCR is required for OCR processing but was not found.",
            font=UIFonts.BODY,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_PRIMARY,
            wraplength=480,
        ).pack(anchor="w", pady=(0, UISpacing.SM))

        tk.Label(
            msg_frame,
            text=get_tesseract_install_instructions(),
            font=UIFonts.SMALL,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_SECONDARY,
            justify="left",
            wraplength=480,
        ).pack(anchor="w", pady=(0, UISpacing.LG))

        # Buttons
        btn_frame = tk.Frame(dialog, bg=UIColors.BG_SECONDARY)
        btn_frame.pack(fill="x", padx=UISpacing.LG, pady=(0, UISpacing.LG))

        def run_install_and_close():
            self._run_winget_install()
            dialog.destroy()

        def check_again():
            self.tesseract_available, _ = check_tesseract_available()
            if self.tesseract_available:
                self._refresh_status_bar()
                dialog.destroy()
                messagebox.showinfo(
                    "Tesseract Found",
                    "Tesseract OCR was detected. You can now process PDFs.",
                    parent=self.root,
                )
            else:
                messagebox.showinfo(
                    "Not Found Yet",
                    "Tesseract still not found.\n\n"
                    "If you just installed it, try closing and reopening this application.",
                    parent=dialog,
                )

        if sys.platform == "win32":
            install_btn = create_rounded_button(
                btn_frame,
                "Install with winget (auto)",
                run_install_and_close,
                style="primary",
            )
            install_btn.pack(side="left", padx=(0, UISpacing.SM))

        check_btn = create_rounded_button(
            btn_frame, "Check again", check_again, style="secondary"
        )
        check_btn.pack(side="left", padx=(0, UISpacing.SM))

        ok_btn = create_rounded_button(btn_frame, "OK", dialog.destroy, style="secondary")
        ok_btn.pack(side="left")

        # Center dialog on screen
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - width) // 2
        y = (dialog.winfo_screenheight() - height) // 2
        dialog.geometry(f"+{x}+{y}")

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
        if OCRMYPDF_AVAILABLE and self.tesseract_available:
            features.append("✓ Local OCR (Tesseract)")
        elif OCRMYPDF_AVAILABLE and not self.tesseract_available:
            features.append("⚠ Local OCR (install Tesseract)")
        else:
            features.append("✗ Local OCR")
        if self._azure_di_configured_flag and PYMUPDF_AVAILABLE:
            features.append("✓ AI OCR (Azure DI)")
        elif PYMUPDF_AVAILABLE:
            features.append("⚠ AI OCR (configure Azure DI)")
        else:
            features.append("✗ AI OCR")
        if PIL_AVAILABLE and IMG2PDF_AVAILABLE:
            features.append("✓ Image to PDF")
        else:
            features.append("✗ Image to PDF")
        self.result_text.insert(tk.END, " | ".join(features) + "\n")
        
        self.result_text.insert(tk.END, "📄 Supported: PDF, JPG, PNG, TIFF, BMP\n")
        
        if not self.tesseract_available and OCRMYPDF_AVAILABLE:
            self.result_text.insert(tk.END, "\n⚠️ Tesseract OCR not found! Install it to use local OCR. See status bar.\n")
        if not self._azure_di_configured_flag:
            self.result_text.insert(tk.END, "💡 Azure AI OCR: configure via launcher ⚙️ Azure button, then restart.\n")
        
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
    try:
        from utils.i18n import init_tool_i18n
    except ImportError:
        from pathlib import Path
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from utils.i18n import init_tool_i18n
    init_tool_i18n(__file__)
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
