# PyPDF Toolbox GUI

A collection of Python-based PDF utility tools with a unified graphical launcher interface.

## Features

- **Slim Launcher Bar**: A compact, always-on-top toolbar that stays at the top of your screen
- **Individual Tool Windows**: Each PDF tool opens in its own window positioned below the launcher
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Virtual Environment**: Automatically creates and manages a Python virtual environment
- **Easy Extension**: Add new tools by simply creating `launch_*.bat`/`launch_*.sh` scripts

## Available Tools

### Implemented Tools

- **Manual Splitter** (`launch_manual_splitter`): Manually split PDFs by selecting split points with visual thumbnail preview
- **Text Extractor** (`launch_text_extractor`): Extract text from PDFs using Python (PyMuPDF), OCR (OCRmyPDF), or Azure AI Document Intelligence. Supports Text, Markdown, and JSON output formats

### Planned Tools

- **Split PDF**: Split a PDF into individual pages or page ranges
- **Merge PDF**: Combine multiple PDFs into one document
- **Reorder Pages**: Rearrange pages within a PDF
- **Compress**: Reduce PDF file size
- **Rotate**: Rotate pages within a PDF
- **Extract Pages**: Extract specific pages from a PDF
- **Convert**: Convert between PDF and other formats
- **Watermark**: Add watermarks to PDF documents
- **Encrypt/Decrypt**: Add or remove password protection
- **Metadata**: View and edit PDF metadata

## Quick Start

### Windows

1. Double-click `launcher.bat`
2. The launcher will automatically create a virtual environment on first run
3. Click on any tool button to launch that tool

### Linux / macOS

1. Make the script executable: `chmod +x launcher.sh`
2. Run: `./launcher.sh`
3. The launcher will automatically create a virtual environment on first run
4. Click on any tool button to launch that tool

## Requirements

- Python 3.8 or higher
- Tkinter (usually included with Python)
- For OCR functionality: Tesseract OCR installed on your system

## Adding New Tools

To add a new PDF tool:

1. Create your tool script in the `src/` directory (e.g., `src/pdf_split.py`)
2. Create a launcher script in the root directory:
   - Windows: `launch_split.bat`
   - Linux/macOS: `launch_split.sh`
3. The tool will automatically appear in the launcher GUI

### Tool Window Positioning

Tools can read these environment variables to position themselves in the available screen area below the launcher:

- `TOOL_WINDOW_X`: X position for tool window
- `TOOL_WINDOW_Y`: Y position for tool window  
- `TOOL_WINDOW_WIDTH`: Available width for tool window
- `TOOL_WINDOW_HEIGHT`: Available height for tool window

### Example Tool Launcher (Windows)

```batch
@echo off
setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%\venv
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe
"%PYTHON_EXE%" "%SCRIPT_DIR%\src\pdf_split.py"
endlocal
```

### Example Tool Launcher (Linux/macOS)

```bash
#!/bin/bash
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
VENV_DIR="$SCRIPT_DIR/venv"
"$VENV_DIR/bin/python" "$SCRIPT_DIR/src/pdf_split.py"
```

## Project Structure

```
pypdf-toolbox-gui/
├── launcher.bat          # Windows launcher script
├── launcher.sh           # Linux/macOS launcher script
├── requirements.txt      # Python dependencies
├── README.md            # This file
├── LICENSE              # Apache 2.0 License
├── launch_*.bat         # Individual tool launchers (Windows)
├── launch_*.sh          # Individual tool launchers (Linux/macOS)
├── src/
│   ├── launcher_gui.py  # Main launcher GUI
│   └── (tool scripts)   # Individual PDF tool scripts
└── venv/                # Virtual environment (auto-created)
```

## License

Copyright 2025-2026 Andre Lorbach

Licensed under the Apache License, Version 2.0. See [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bug reports and feature requests.
