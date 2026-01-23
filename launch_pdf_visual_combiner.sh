#!/bin/bash
# ============================================
# PDF Visual Combiner Tool Launcher
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
