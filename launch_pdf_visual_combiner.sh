#!/bin/bash

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
VENV_DIR="$SCRIPT_DIR/.venv"
PYTHON_EXE="$VENV_DIR/bin/python"

# Check if venv exists
if [ ! -f "$PYTHON_EXE" ]; then
    echo "[ERROR] Virtual environment not found."
    echo "[INFO] Please run launcher.sh first to create the virtual environment."
    exit 1
fi

# Launch the PDF Visual Combiner tool
"$PYTHON_EXE" "$SCRIPT_DIR/src/pdf_combiner.py" "$@"
