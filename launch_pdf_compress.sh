#!/bin/bash
# ============================================
# PDF / Image Recompress Tool Launcher
#
# Copyright 2025-2026 Andre Lorbach
# Licensed under the Apache License, Version 2.0
# ============================================

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
VENV_DIR="$SCRIPT_DIR/.venv"
PYTHON_EXE="$VENV_DIR/bin/python"

if [ ! -f "$PYTHON_EXE" ]; then
    echo "[ERROR] Virtual environment not found."
    echo "[INFO] Please run launcher.sh first to create the virtual environment."
    exit 1
fi

"$PYTHON_EXE" "$SCRIPT_DIR/src/pdf_compress.py" "$@"
