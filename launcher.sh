#!/bin/bash
# ============================================
# PyPDF Toolbox - Main Launcher
#
# Copyright 2025-2026 Andre Lorbach
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# ============================================

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR="$SCRIPT_DIR"
VENV_DIR="$ROOT_DIR/.venv"
PYTHON_EXE="$VENV_DIR/bin/python"

echo "============================================"
echo "PyPDF Toolbox - Main Launcher"
echo "============================================"
echo ""

# Check if virtual environment exists
if [ ! -d "$VENV_DIR" ]; then
    echo "[INFO] Virtual environment not found. Creating one..."
    python3 -m venv "$VENV_DIR"
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to create virtual environment."
        echo "[INFO] Make sure Python 3.7+ is installed and in PATH."
        exit 1
    fi
    echo "[SUCCESS] Virtual environment created."
    echo ""
fi

# Check if Python executable exists
if [ ! -f "$PYTHON_EXE" ]; then
    echo "[ERROR] Python executable not found in venv."
    echo "[INFO] Expected location: $PYTHON_EXE"
    exit 1
fi

# Install/update dependencies
echo "[INFO] Checking dependencies..."
"$PYTHON_EXE" -m pip --version >/dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "[ERROR] pip not available in virtual environment."
    exit 1
fi

if [ -f "$ROOT_DIR/requirements.txt" ]; then
    echo "[INFO] Installing/updating requirements from requirements.txt..."
    "$PYTHON_EXE" -m pip install --upgrade -r "$ROOT_DIR/requirements.txt"
    if [ $? -eq 0 ]; then
        echo "[SUCCESS] All dependencies installed/updated."
    else
        echo "[WARNING] Some dependencies may have failed to install."
        echo "[INFO] You can manually install with: pip install -r requirements.txt"
    fi
else
    echo "[WARNING] requirements.txt not found."
fi

echo ""
echo "[INFO] Launching PyPDF Toolbox GUI..."
echo ""

# Launch the GUI launcher
"$PYTHON_EXE" "$ROOT_DIR/src/launcher_gui.py"

if [ $? -ne 0 ]; then
    echo "[ERROR] Failed to launch GUI launcher."
    exit 1
fi
