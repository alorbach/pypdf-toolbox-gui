@echo off
REM ============================================
REM PDF Markdown Converter Tool Launcher
REM
REM Copyright 2025-2026 Andre Lorbach
REM
REM Licensed under the Apache License, Version 2.0 (the "License");
REM you may not use this file except in compliance with the License.
REM You may obtain a copy of the License at
REM
REM     http://www.apache.org/licenses/LICENSE-2.0
REM
REM Unless required by applicable law or agreed to in writing, software
REM distributed under the License is distributed on an "AS IS" BASIS,
REM WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
REM See the License for the specific language governing permissions and
REM limitations under the License.
REM ============================================

setlocal

set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe
set PYTHONW_EXE=%VENV_DIR%\Scripts\pythonw.exe
set REQUIREMENTS=%SCRIPT_DIR%requirements.txt

REM Check if venv exists
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment not found.
    echo [INFO] Please run launcher.bat first to create the virtual environment.
    pause
    exit /b 1
)

REM Check and install dependencies if needed
if exist "%REQUIREMENTS%" (
    echo [INFO] Checking dependencies for Markdown Converter...
    "%PYTHON_EXE%" -m pip install --quiet --upgrade -r "%REQUIREMENTS%" 2>nul
    if errorlevel 1 (
        echo [WARNING] Some dependencies may have failed to install.
        echo [INFO] You can manually install with: pip install -r requirements.txt
    )
)

REM Use pythonw.exe if available (no console window), otherwise fall back to python.exe
if exist "%PYTHONW_EXE%" (
    "%PYTHONW_EXE%" "%SCRIPT_DIR%src\pdf_md_converter.py" %*
) else (
    "%PYTHON_EXE%" "%SCRIPT_DIR%src\pdf_md_converter.py" %*
)

endlocal
