# PyPDF Toolbox - Agent Guidelines

This document provides guidelines for AI agents working on this project.

## Session Commands

### SUMMARIZE Command

When the user types **SUMMARIZE** (or requests a "squashed commit message" or "COPY READY" commit message), create a squashed commit message summarizing all changes made during the session. The message should be:
- **Copy-ready** (user can paste directly into git commit -m "")
- **Follow conventional commit format** (feat:, fix:, refactor:, docs:, etc.)
- **Include a brief summary line** (max 72 characters) followed by bullet points of changes
- **Group related changes together** logically
- **Be specific** about what was changed and in which files
- **Use present tense** ("Add feature" not "Added feature")
- **Focus on user-visible changes** and important technical changes

**Format:**
```
<type>: <brief summary (max 72 chars)>

- <Change 1>
- <Change 2>
- <Change 3>
```

**Conventional Commit Types:**
- `feat:` - New feature
- `fix:` - Bug fix
- `refactor:` - Code refactoring
- `docs:` - Documentation changes
- `style:` - UI/styling changes
- `perf:` - Performance improvements
- `chore:` - Maintenance tasks

**Example formats:**

```
feat: Add wait cursor feedback during PDF processing in Text Extractor

- Add wait cursor (hourglass) when PDF files are dropped or processing starts
- Implement cursor management in process_files() method with try/finally block
- Ensure cursor is restored to normal state after processing completes or on errors
- Works for all input methods: drag & drop, file selection, and folder selection
- Force immediate cursor update with root.update() for better UX feedback
```

```
feat: Add tool categorization and grouping in launcher GUI

- Add _get_tool_category() method to categorize tools by functionality
- Group tools by category: split_merge, extract_analyze, convert_transform, optimize, security, annotate, other
- Add visual category separators and labels in launcher toolbar
- Sort tools by category first, then alphabetically within category
- Add category metadata to launcher entries
- Update populate_tools() to display categorized groups with separators
```

## Project Overview

PyPDF Toolbox GUI is a collection of Python-based PDF utility tools with a unified graphical launcher interface. The launcher is a slim top-bar that stays at the top of the screen, and individual tools open in windows positioned below it.

## Architecture

### Launcher System

- **Main Launcher**: `src/launcher_gui.py` - Slim top-bar GUI that discovers and launches tools
- **Launcher Scripts**: `launcher.bat` (Windows) / `launcher.sh` (Linux/macOS) - Bootstrap scripts that create venv and launch the GUI
- **Tool Discovery**: Tools are discovered by scanning for `launch_*.bat` / `launch_*.sh` files in the root directory

### Tool Window Positioning

Tools receive positioning information via environment variables:
- `TOOL_WINDOW_X`: X position for tool window
- `TOOL_WINDOW_Y`: Y position for tool window
- `TOOL_WINDOW_WIDTH`: Available width for tool window
- `TOOL_WINDOW_HEIGHT`: Available height for tool window

Tools should use these to position themselves in the screen area below the launcher bar.

## Coding Standards

### GUI Framework

- Use **tkinter** with **ttk** widgets for all GUI components
- Apply modern styling using the UIColors, UIFonts, and UISpacing classes
- Use consistent fonts: "Segoe UI" on Windows, system default on other platforms

## Modern UI Design Guidelines

All tools in this project MUST follow these modern UI/UX best practices for a consistent, professional appearance.

### Color System (UIColors Class)

Use the predefined color palette in `UIColors` class for consistency:

```python
class UIColors:
    """Modern color palette for consistent UI styling."""
    # Primary colors (Blue)
    PRIMARY = "#2563eb"          # Main actions, links, active states
    PRIMARY_HOVER = "#1d4ed8"    # Hover state for primary
    PRIMARY_LIGHT = "#dbeafe"    # Backgrounds, badges
    
    # Secondary (Slate gray)
    SECONDARY = "#64748b"
    SECONDARY_HOVER = "#475569"
    
    # Status colors
    SUCCESS = "#16a34a"          # Green - success actions, confirmations
    SUCCESS_LIGHT = "#dcfce7"
    SUCCESS_HOVER = "#15803d"
    ERROR = "#dc2626"            # Red - errors, destructive actions
    ERROR_LIGHT = "#fee2e2"
    ERROR_HOVER = "#b91c1c"
    WARNING = "#f59e0b"          # Amber - warnings
    WARNING_LIGHT = "#fef3c7"
    
    # Neutral backgrounds
    BG_PRIMARY = "#ffffff"       # White - cards, dialogs
    BG_SECONDARY = "#f8fafc"     # Slate 50 - page background
    BG_TERTIARY = "#f1f5f9"      # Slate 100 - subtle sections
    
    # Borders and text
    BORDER = "#e2e8f0"           # Light border
    BORDER_DARK = "#cbd5e1"      # Emphasized border
    TEXT_PRIMARY = "#1e293b"     # Main text
    TEXT_SECONDARY = "#64748b"   # Secondary text
    TEXT_MUTED = "#94a3b8"       # Disabled/placeholder text
    
    # Special purpose
    SPLIT_ACTIVE = "#ef4444"     # Red for active split points
    THUMBNAIL_BG = "#ffffff"
    THUMBNAIL_HOVER = "#dbeafe"  # Blue highlight on hover
    
    # Drop zone
    DROP_ZONE_BG = "#f8fafc"
    DROP_ZONE_BORDER = "#94a3b8"
    DROP_ZONE_ACTIVE = "#dbeafe"
    DROP_ZONE_BORDER_ACTIVE = "#2563eb"
```

### Typography (UIFonts Class)

Use consistent fonts throughout the application:

```python
class UIFonts:
    """Font configurations for consistent typography."""
    TITLE = ("Segoe UI", 18, "bold")      # Main window titles
    SUBTITLE = ("Segoe UI", 14, "bold")   # Section titles
    HEADING = ("Segoe UI", 12, "bold")    # Card headers, labels
    BODY = ("Segoe UI", 10)               # Regular text
    BODY_BOLD = ("Segoe UI", 10, "bold")  # Emphasized body text
    SMALL = ("Segoe UI", 9)               # Secondary info, hints
    SMALL_BOLD = ("Segoe UI", 9, "bold")  # Small labels
    MONO = ("Consolas", 9)                # Code, file paths, logs
    BUTTON = ("Segoe UI", 10, "bold")     # Button text
    BUTTON_SMALL = ("Segoe UI", 9)        # Small buttons
```

### Spacing System (UISpacing Class)

Use consistent spacing values:

```python
class UISpacing:
    """Consistent spacing values."""
    XS = 2    # Tight spacing (between related elements)
    SM = 5    # Small spacing (padding, small gaps)
    MD = 10   # Medium spacing (standard padding)
    LG = 15   # Large spacing (section gaps)
    XL = 20   # Extra large (major sections)
    XXL = 30  # Maximum spacing
```

### Button Styling

Create modern flat buttons with hover effects:

```python
def create_rounded_button(parent, text, command, style="primary", width=None):
    """Create a styled button with hover effect.
    
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
```

### Card/Panel Design

Create card-like frames with proper borders:

```python
def create_card_frame(parent, title=None):
    """Create a card-like frame with optional title."""
    if title:
        frame = tk.LabelFrame(
            parent,
            text=f"  {title}  ",  # Add spacing around title
            font=UIFonts.HEADING,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            bd=1,
            relief="solid",
            padx=UISpacing.MD,
            pady=UISpacing.SM,
        )
    else:
        frame = tk.Frame(
            parent,
            bg=UIColors.BG_PRIMARY,
            bd=1,
            relief="solid",
        )
    return frame
```

### Header Design Pattern

Every tool should have a consistent header:

```python
def create_header(parent, title, subtitle=None):
    """Create a styled header section."""
    header_frame = tk.Frame(parent, bg=UIColors.BG_PRIMARY, pady=UISpacing.MD)
    header_frame.grid(row=0, column=0, sticky="ew")
    header_frame.grid_columnconfigure(0, weight=1)
    
    # Title with icon
    title_label = tk.Label(
        header_frame,
        text=f"✂️  {title}",  # Add relevant emoji
        font=UIFonts.TITLE,
        bg=UIColors.BG_PRIMARY,
        fg=UIColors.PRIMARY
    )
    title_label.grid(row=0, column=0, pady=(UISpacing.SM, 0))
    
    if subtitle:
        subtitle_label = tk.Label(
            header_frame,
            text=subtitle,
            font=UIFonts.BODY,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_SECONDARY
        )
        subtitle_label.grid(row=1, column=0, pady=(0, UISpacing.SM))
    
    return header_frame
```

### Drop Zone Design

Modern drop zone with visual feedback:

```python
def create_drop_zone(parent):
    """Create a modern drop zone."""
    drop_frame = tk.Frame(
        parent,
        bg=UIColors.DROP_ZONE_BG,
        highlightbackground=UIColors.DROP_ZONE_BORDER,
        highlightthickness=2,
        padx=UISpacing.XL,
        pady=UISpacing.XL
    )
    
    # Large icon
    icon_label = tk.Label(
        drop_frame,
        text="📄",
        font=("Segoe UI", 32),
        bg=UIColors.DROP_ZONE_BG
    )
    icon_label.pack(pady=(UISpacing.SM, UISpacing.XS))
    
    # Main text
    main_label = tk.Label(
        drop_frame,
        text="Drag and drop PDF file here",
        font=UIFonts.SUBTITLE,
        bg=UIColors.DROP_ZONE_BG,
        fg=UIColors.TEXT_PRIMARY,
        cursor="hand2"
    )
    main_label.pack()
    
    # Sub text
    sub_label = tk.Label(
        drop_frame,
        text="or click to browse",
        font=UIFonts.SMALL,
        bg=UIColors.DROP_ZONE_BG,
        fg=UIColors.TEXT_MUTED
    )
    sub_label.pack(pady=(UISpacing.XS, UISpacing.SM))
    
    return drop_frame
```

### Interactive Elements

Add hover effects to interactive elements:

```python
def add_hover_effect(widget, normal_bg, hover_bg):
    """Add hover effect to a widget."""
    def on_enter(e):
        widget.config(bg=hover_bg)
        # Also update child labels
        for child in widget.winfo_children():
            if isinstance(child, tk.Label):
                child.config(bg=hover_bg)
    
    def on_leave(e):
        widget.config(bg=normal_bg)
        for child in widget.winfo_children():
            if isinstance(child, tk.Label):
                child.config(bg=normal_bg)
    
    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)
```

### Status Indicators

Use badges for status information:

```python
def create_status_badge(parent, text, status="info"):
    """Create a colored status badge.
    
    Status: info, success, warning, error
    """
    colors = {
        "info": (UIColors.PRIMARY_LIGHT, UIColors.PRIMARY),
        "success": (UIColors.SUCCESS_LIGHT, UIColors.SUCCESS),
        "warning": (UIColors.WARNING_LIGHT, UIColors.TEXT_PRIMARY),
        "error": (UIColors.ERROR_LIGHT, UIColors.ERROR),
    }
    
    bg, fg = colors.get(status, colors["info"])
    
    badge = tk.Label(
        parent,
        text=f" {text} ",
        font=UIFonts.SMALL_BOLD,
        bg=bg,
        fg=fg,
        padx=UISpacing.SM,
        pady=UISpacing.XS
    )
    return badge
```

### Window/Dialog Design

Modal dialogs should be on top of parent:

```python
def create_modal_dialog(parent, title, width=500, height=400):
    """Create a modal dialog window."""
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.geometry(f"{width}x{height}")
    dialog.configure(bg=UIColors.BG_SECONDARY)
    
    # Make modal and on top
    dialog.transient(parent)
    dialog.grab_set()
    dialog.lift()
    dialog.focus_force()
    
    # Center on parent
    dialog.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() - width) // 2
    y = parent.winfo_y() + (parent.winfo_height() - height) // 2
    dialog.geometry(f"+{x}+{y}")
    
    return dialog
```

### UI Best Practices Summary

1. **Consistent Colors**: Always use `UIColors` constants, never hardcode colors
2. **Proper Spacing**: Use `UISpacing` constants for all padding/margins
3. **Typography Hierarchy**: Use appropriate font sizes (Title > Subtitle > Heading > Body > Small)
4. **Interactive Feedback**: Add hover effects to all clickable elements
5. **Visual Grouping**: Use cards/panels to group related content
6. **Status Communication**: Use color-coded badges and icons for status
7. **Flat Design**: Use `relief="flat"` for buttons, subtle borders for containers
8. **Cursor Feedback**: Set `cursor="hand2"` for clickable elements
9. **Modal Dialogs**: Always make dialogs modal with `transient()` and `grab_set()`
10. **Icons/Emojis**: Use emojis sparingly for visual cues (📄, ✂️, ✓, ⚠️)

### File Structure

```
pypdf-toolbox-gui/
├── launcher.bat / launcher.sh    # Main launcher scripts
├── launch_*.bat / launch_*.sh    # Individual tool launchers
├── requirements.txt              # Python dependencies
├── src/
│   ├── launcher_gui.py          # Main launcher GUI
│   ├── pdf_*.py                 # Individual tool scripts
│   └── utils/                   # Shared utilities
└── venv/                        # Virtual environment (auto-created)
```

### Naming Conventions

- Tool scripts: `src/pdf_<toolname>.py` (e.g., `pdf_split.py`, `pdf_merge.py`)
- Launcher scripts: `launch_<toolname>.bat` / `launch_<toolname>.sh`
- Class names: PascalCase (e.g., `PDFSplitTool`)
- Functions/methods: snake_case (e.g., `split_pdf()`)
- Constants: UPPER_SNAKE_CASE (e.g., `MAX_FILE_SIZE`)

## CRITICAL: Drag and Drop Support

**All PDF tools MUST support drag and drop for input files.**

### Implementation Pattern

Every tool GUI must implement drag and drop using tkinter's DnD (drag and drop) capabilities. Use the `tkinterdnd2` library for cross-platform drag and drop support.

#### Required Setup

1. Add `tkinterdnd2` to requirements.txt
2. Use `TkinterDnD.Tk()` instead of `tk.Tk()` as the root window
3. Register drop targets on file input areas

#### Standard Drag and Drop Implementation

```python
from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFToolBase:
    def __init__(self):
        # Use TkinterDnD instead of tk.Tk
        self.root = TkinterDnD.Tk()
        
        # ... setup UI ...
        
        # Register drop target on the main frame or specific widget
        self.setup_drag_drop()
    
    def setup_drag_drop(self, widget=None):
        """Setup drag and drop on a widget (defaults to root)"""
        target = widget or self.root
        target.drop_target_register(DND_FILES)
        target.dnd_bind('<<Drop>>', self.on_drop)
        target.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        target.dnd_bind('<<DragLeave>>', self.on_drag_leave)
    
    def on_drop(self, event):
        """Handle dropped files"""
        # event.data contains the dropped file path(s)
        # On Windows, paths with spaces are wrapped in {}
        files = self.parse_dropped_files(event.data)
        self.process_dropped_files(files)
    
    def on_drag_enter(self, event):
        """Visual feedback when dragging over"""
        # Change background color or show indicator
        event.widget.configure(style='DropTarget.TFrame')
    
    def on_drag_leave(self, event):
        """Reset visual feedback"""
        event.widget.configure(style='TFrame')
    
    def parse_dropped_files(self, data):
        """Parse dropped file paths (handles Windows {} wrapping)"""
        files = []
        # Handle Windows path format with {} for spaces
        if '{' in data:
            import re
            files = re.findall(r'\{([^}]+)\}', data)
            # Also get non-bracketed paths
            remaining = re.sub(r'\{[^}]+\}', '', data).strip()
            if remaining:
                files.extend(remaining.split())
        else:
            files = data.split()
        
        # Filter for PDF files only
        return [f for f in files if f.lower().endswith('.pdf')]
    
    def process_dropped_files(self, files):
        """Override in subclass to handle the dropped files"""
        raise NotImplementedError("Subclass must implement process_dropped_files")
```

#### Visual Drop Zone

Create a visible drop zone area in the UI:

```python
def create_drop_zone(self, parent):
    """Create a visual drop zone for files"""
    drop_frame = ttk.LabelFrame(parent, text="Drop PDF Files Here", padding=20)
    
    drop_label = ttk.Label(
        drop_frame,
        text="📄 Drag and drop PDF files here\nor click to browse",
        font=("Segoe UI", 11),
        anchor='center',
        justify='center'
    )
    drop_label.pack(expand=True, fill='both', pady=20)
    
    # Setup drag and drop on this frame
    self.setup_drag_drop(drop_frame)
    
    # Also allow click to open file dialog
    drop_frame.bind('<Button-1>', self.browse_files)
    drop_label.bind('<Button-1>', self.browse_files)
    
    return drop_frame
```

#### Fallback for Systems Without tkinterdnd2

If `tkinterdnd2` is not available, gracefully fall back to file dialog only:

```python
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

class PDFTool:
    def __init__(self):
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
            print("[WARNING] tkinterdnd2 not available, drag and drop disabled")
```

## Tool Template

When creating a new PDF tool, use this template structure:

```python
"""
PDF <ToolName> Tool

<Description of what this tool does>

Copyright 2025-2026 Andre Lorbach
Licensed under Apache License 2.0
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


class PDF<ToolName>Tool:
    def __init__(self):
        # Initialize root window with DnD support
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("PDF <ToolName>")
        
        # Position window using environment variables from launcher
        self.position_window()
        
        # Setup UI
        self.setup_ui()
        
        # Setup drag and drop
        if HAS_DND:
            self.setup_drag_drop()
    
    def position_window(self):
        """Position window in the area below the launcher"""
        x = int(os.environ.get('TOOL_WINDOW_X', 100))
        y = int(os.environ.get('TOOL_WINDOW_Y', 100))
        width = int(os.environ.get('TOOL_WINDOW_WIDTH', 800))
        height = int(os.environ.get('TOOL_WINDOW_HEIGHT', 600))
        
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        """Setup the tool UI"""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill='both', expand=True)
        
        # Drop zone for files
        self.drop_zone = self.create_drop_zone(main_frame)
        self.drop_zone.pack(fill='x', pady=(0, 10))
        
        # Tool-specific UI elements go here
        # ...
        
        # Action buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        
        self.process_btn = ttk.Button(
            btn_frame,
            text="Process",
            command=self.process_files
        )
        self.process_btn.pack(side='right')
    
    def create_drop_zone(self, parent):
        """Create visual drop zone"""
        # Implementation as shown above
        pass
    
    def setup_drag_drop(self):
        """Setup drag and drop handlers"""
        # Implementation as shown above
        pass
    
    def process_files(self):
        """Main processing logic"""
        # Tool-specific implementation
        pass
    
    def run(self):
        self.root.mainloop()


def main():
    app = PDF<ToolName>Tool()
    app.run()


if __name__ == "__main__":
    main()
```

## Dependencies

Core dependencies (in requirements.txt):
- `PyPDF2` / `pypdf` - PDF manipulation
- `pdf2image` - PDF to image conversion
- `reportlab` - PDF creation
- `pytesseract` - OCR support
- `pillow` - Image processing
- `tkinterdnd2` - Drag and drop support
- `tqdm` - Progress bars
- `openai` - Azure OpenAI API client (for AI-powered tools)

## CRITICAL: Azure AI API Configuration

**All tools that use Azure AI APIs MUST share a common configuration file.**

**Azure configuration is now centrally managed in the launcher GUI** - users can configure Azure AI settings once, and all tools will use the same configuration.

### Configuration File Location

Azure AI configuration is stored in: `config/azure_ai.yaml`

This file is shared across ALL tools that require Azure AI services (OCR, document analysis, AI-assisted splitting, etc.).

### Configuration File Format

```yaml
# Azure AI Configuration
# This file is shared by all tools that use Azure AI services

azure_openai:
  # Azure OpenAI endpoint URL
  endpoint: "https://your-resource-name.openai.azure.com/"
  # API key (consider using environment variable AZURE_OPENAI_API_KEY instead)
  api_key: ""
  # API version
  api_version: "2024-02-15-preview"
  # Deployment name for the model
  deployment_name: "gpt-4"

azure_document_intelligence:
  # Azure Document Intelligence (Form Recognizer) endpoint
  endpoint: "https://your-resource-name.cognitiveservices.azure.com/"
  # API key (consider using environment variable AZURE_DOC_INTEL_API_KEY instead)
  api_key: ""

# Settings
settings:
  # Prefer environment variables over config file for API keys
  prefer_env_vars: true
  # Timeout for API calls in seconds
  timeout: 60
  # Maximum retries for failed API calls
  max_retries: 3
```

### Loading Configuration

Use this standard pattern to load Azure AI configuration in tools:

```python
import os
import yaml
from pathlib import Path

class AzureAIConfig:
    """Shared Azure AI configuration loader."""
    
    CONFIG_FILE = "config/azure_ai.yaml"
    
    def __init__(self):
        self.config = self._load_config()
    
    def _load_config(self):
        """Load configuration from file and environment variables."""
        config = {
            'azure_openai': {
                'endpoint': '',
                'api_key': '',
                'api_version': '2024-02-15-preview',
                'deployment_name': 'gpt-4'
            },
            'azure_document_intelligence': {
                'endpoint': '',
                'api_key': ''
            },
            'settings': {
                'prefer_env_vars': True,
                'timeout': 60,
                'max_retries': 3
            }
        }
        
        # Try to load from config file
        config_path = Path(__file__).parent.parent / self.CONFIG_FILE
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    file_config = yaml.safe_load(f)
                    if file_config:
                        self._merge_config(config, file_config)
            except Exception as e:
                print(f"[WARNING] Could not load config file: {e}")
        
        # Override with environment variables if preferred
        if config['settings'].get('prefer_env_vars', True):
            self._load_env_vars(config)
        
        return config
    
    def _merge_config(self, base, override):
        """Recursively merge override into base config."""
        for key, value in override.items():
            if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                self._merge_config(base[key], value)
            else:
                base[key] = value
    
    def _load_env_vars(self, config):
        """Load API keys from environment variables."""
        # Azure OpenAI
        if os.environ.get('AZURE_OPENAI_ENDPOINT'):
            config['azure_openai']['endpoint'] = os.environ['AZURE_OPENAI_ENDPOINT']
        if os.environ.get('AZURE_OPENAI_API_KEY'):
            config['azure_openai']['api_key'] = os.environ['AZURE_OPENAI_API_KEY']
        if os.environ.get('AZURE_OPENAI_DEPLOYMENT'):
            config['azure_openai']['deployment_name'] = os.environ['AZURE_OPENAI_DEPLOYMENT']
        
        # Azure Document Intelligence
        if os.environ.get('AZURE_DOC_INTEL_ENDPOINT'):
            config['azure_document_intelligence']['endpoint'] = os.environ['AZURE_DOC_INTEL_ENDPOINT']
        if os.environ.get('AZURE_DOC_INTEL_API_KEY'):
            config['azure_document_intelligence']['api_key'] = os.environ['AZURE_DOC_INTEL_API_KEY']
    
    @property
    def openai_endpoint(self):
        return self.config['azure_openai']['endpoint']
    
    @property
    def openai_api_key(self):
        return self.config['azure_openai']['api_key']
    
    @property
    def openai_deployment(self):
        return self.config['azure_openai']['deployment_name']
    
    @property
    def openai_api_version(self):
        return self.config['azure_openai']['api_version']
    
    @property
    def doc_intel_endpoint(self):
        return self.config['azure_document_intelligence']['endpoint']
    
    @property
    def doc_intel_api_key(self):
        return self.config['azure_document_intelligence']['api_key']
    
    def is_openai_configured(self):
        """Check if Azure OpenAI is properly configured."""
        return bool(self.openai_endpoint and self.openai_api_key)
    
    def is_doc_intel_configured(self):
        """Check if Azure Document Intelligence is properly configured."""
        return bool(self.doc_intel_endpoint and self.doc_intel_api_key)


# Usage in tools:
# from utils.azure_config import AzureAIConfig
# 
# config = AzureAIConfig()
# if config.is_openai_configured():
#     client = AzureOpenAI(
#         azure_endpoint=config.openai_endpoint,
#         api_key=config.openai_api_key,
#         api_version=config.openai_api_version
#     )
```

### Environment Variables

For security, prefer environment variables over storing API keys in config files:

| Variable | Description |
|----------|-------------|
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API key |
| `AZURE_OPENAI_DEPLOYMENT` | Azure OpenAI deployment/model name |
| `AZURE_DOC_INTEL_ENDPOINT` | Azure Document Intelligence endpoint |
| `AZURE_DOC_INTEL_API_KEY` | Azure Document Intelligence API key |

### Launcher Integration

The launcher GUI (`src/launcher_gui.py`) includes a **"⚙️ Azure"** button that opens a global Azure AI configuration dialog. This allows users to:

1. Configure Azure OpenAI settings (endpoint, API key, deployment, API version)
2. Configure Azure Document Intelligence settings (endpoint, API key)
3. View configuration status for both services
4. Test connections
5. Save configuration to `config/azure_ai.yaml`

**All tools automatically use this shared configuration** - no need for individual tool configuration dialogs.

### Configuration UI (For Tools)

If a tool needs to show Azure configuration status or provide a quick link to the launcher's config:

1. Show current configuration status (configured/not configured) using `AzureAIConfig.get_status_text()`
2. Provide a button/link that opens the launcher's Azure config dialog (if accessible)
3. Display helpful messages if Azure features are not configured

Example configuration status display:

```python
def show_azure_config_dialog(self):
    """Show dialog to configure Azure AI settings."""
    dialog = tk.Toplevel(self.root)
    dialog.title("Azure AI Configuration")
    dialog.geometry("500x400")
    dialog.transient(self.root)
    dialog.grab_set()
    
    # Azure OpenAI section
    openai_frame = ttk.LabelFrame(dialog, text="Azure OpenAI", padding=10)
    openai_frame.pack(fill='x', padx=10, pady=5)
    
    ttk.Label(openai_frame, text="Endpoint:").grid(row=0, column=0, sticky='w')
    endpoint_entry = ttk.Entry(openai_frame, width=50)
    endpoint_entry.grid(row=0, column=1, padx=5, pady=2)
    endpoint_entry.insert(0, self.azure_config.openai_endpoint)
    
    ttk.Label(openai_frame, text="API Key:").grid(row=1, column=0, sticky='w')
    apikey_entry = ttk.Entry(openai_frame, width=50, show='*')
    apikey_entry.grid(row=1, column=1, padx=5, pady=2)
    
    # ... more fields ...
    
    # Save button
    def save_config():
        # Save to config/azure_ai.yaml
        pass
    
    ttk.Button(dialog, text="Save", command=save_config).pack(pady=10)
```

### Window Positioning and Modal Dialogs

**CRITICAL**: When creating child windows or dialogs:

1. **Always pass parent window**: When creating dialogs from a tool, pass the parent window reference
2. **Use Toplevel for child windows**: Use `tk.Toplevel(parent)` instead of `tk.Tk()` when there's a parent
3. **Make dialogs modal**: Use `transient(parent)` and `grab_set()` to keep dialogs on top and modal
4. **Ensure visibility**: Call `lift()` and `focus_force()` to ensure dialogs appear on top

Example:
```python
def show_dialog(self):
    """Show a modal dialog that stays on top of parent."""
    dialog = tk.Toplevel(self.root)  # Use Toplevel, not Tk
    dialog.transient(self.root)      # Associate with parent
    dialog.grab_set()                # Make modal
    dialog.lift()                    # Bring to front
    dialog.focus_force()             # Focus
    # ... dialog content ...
```

### Important Notes

1. **Never commit API keys** - Add `config/azure_ai.yaml` to `.gitignore`
2. **Create template file** - Provide `config/azure_ai.yaml.template` without actual keys
3. **Validate on startup** - Check if configuration exists and is valid
4. **Show clear errors** - If AI features are used but not configured, show helpful message
5. **Graceful degradation** - Tools should work without AI features if not configured
6. **Centralized configuration** - Use the launcher's Azure config button for global configuration

## Error Handling

- Always wrap PDF operations in try/except blocks
- Show user-friendly error messages via `messagebox.showerror()`
- Log detailed errors to console for debugging
- Handle common issues: file not found, corrupted PDF, permission denied, password protected

## Testing

- Test drag and drop with single and multiple files
- Test with files containing spaces in paths
- Test with various PDF types (scanned, digital, encrypted)
- Test on Windows, macOS, and Linux if possible

## Documentation

**All tools MUST have corresponding documentation in the `doc/` folder.**

### Documentation Structure

```
doc/
├── README.md                      # Main documentation index
├── <tool-name>/
│   ├── README.md                  # Tool documentation
│   └── screenshots/
│       └── .gitkeep               # Placeholder for screenshots
```

### When Creating a New Tool

1. Create a new folder under `doc/` matching the tool name (e.g., `doc/pdf-splitter/`)
2. Add a `README.md` with:
   - Tool description and features
   - Screenshots table (placeholders for images to be added later)
   - Usage instructions
   - Technical details (source file, framework, dependencies)
   - Keyboard shortcuts (if applicable)
3. Create a `screenshots/` subfolder with a `.gitkeep` file listing expected screenshots
4. Update `doc/README.md` to include the new tool in the index table
5. **Update main `README.md`**: When a screenshot becomes available (`01-main-window.png`), add the tool to the 3-column image grid in the "Available Tools" section

### Screenshot Naming Convention

Use numbered, descriptive names:
- `01-main-window.png` - Main interface
- `02-file-loaded.png` - After loading a file
- `03-feature-name.png` - Specific feature demonstration

### Documentation Template

When creating documentation for a new tool, include these sections:

```markdown
# PDF <ToolName> Documentation

<Brief description>

## Features

- Feature 1
- Feature 2

## Screenshots

*Screenshots will be added to the `screenshots/` folder.*

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main interface |

## Usage

### Opening Files
...

## Technical Details

- **Source File**: `src/pdf_<toolname>.py`
- **Launch Script**: `launch_<toolname>.bat` / `launch_<toolname>.sh`

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Open file |
```

### Main README.md Image Grid

The main `README.md` displays tools in a **3-column image grid** format:

```markdown
<div align="center">

| | | |
|---|---|---|
| [![Tool 1](doc/tool1/screenshots/01-main-window.png)](doc/tool1/) | [![Tool 2](doc/tool2/screenshots/01-main-window.png)](doc/tool2/) | [![Tool 3](doc/tool3/screenshots/01-main-window.png)](doc/tool3/) |
| **Tool 1**<br/>Description | **Tool 2**<br/>Description | **Tool 3**<br/>Description |
| [`launch_tool1`](doc/tool1/) | [`launch_tool2`](doc/tool2/) | [`launch_tool3`](doc/tool3/) |

</div>
```

**Guidelines:**
- Tools are displayed in a 3-column grid (3 tools per row)
- Each tool shows: screenshot thumbnail, name, brief description, and launch script link
- When adding a new tool, append it to the grid (start a new row if needed)
- Only include tools that have a `01-main-window.png` screenshot available
- If a tool doesn't have a screenshot yet, add it to the grid once the screenshot is created

### Important Notes

- Keep documentation in sync with code changes
- Update screenshots when UI changes significantly
- Document all keyboard shortcuts and features
- Include troubleshooting tips for common issues
- Update main README.md image grid when new tool screenshots become available
