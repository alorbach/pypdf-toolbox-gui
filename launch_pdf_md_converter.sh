#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"
PYTHON_EXE="$VENV_DIR/bin/python"
REQUIREMENTS="$SCRIPT_DIR/requirements.txt"

# Check if venv exists
if [ ! -f "$PYTHON_EXE" ]; then
    echo "[ERROR] Virtual environment not found."
    echo "[INFO] Please run launcher.sh first to create the virtual environment."
    exit 1
fi

# Check and install dependencies if needed
if [ -f "$REQUIREMENTS" ]; then
    echo "[INFO] Checking dependencies for Markdown Converter..."
    "$PYTHON_EXE" -m pip install --quiet --upgrade -r "$REQUIREMENTS" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "[WARNING] Some dependencies may have failed to install."
        echo "[INFO] You can manually install with: pip install -r requirements.txt"
    fi
fi

# Run the script
"$PYTHON_EXE" "$SCRIPT_DIR/src/pdf_md_converter.py" "$@"
