@echo off
REM ============================================
REM PyPDF Toolbox - Main Launcher
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
set ROOT_DIR=%SCRIPT_DIR%
set VENV_DIR=%ROOT_DIR%\.venv
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe

echo ============================================
echo PyPDF Toolbox - Main Launcher
echo ============================================
echo.

REM Check if virtual environment exists
if not exist "%VENV_DIR%\" (
    echo [INFO] Virtual environment not found. Creating one...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        echo [INFO] Make sure Python is installed and in PATH.
        pause
        exit /b 1
    )
    echo [SUCCESS] Virtual environment created.
    echo.
)

REM Check if Python executable exists
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python executable not found in venv.
    echo [INFO] Expected location: %PYTHON_EXE%
    pause
    exit /b 1
)

REM Install/update dependencies
echo [INFO] Checking dependencies...
"%PYTHON_EXE%" -m pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip not available in virtual environment.
    pause
    exit /b 1
)

if exist "%ROOT_DIR%\requirements.txt" (
    echo [INFO] Installing/updating requirements from requirements.txt...
    "%PYTHON_EXE%" -m pip install --upgrade -r "%ROOT_DIR%\requirements.txt"
    if errorlevel 1 (
        echo [WARNING] Some dependencies may have failed to install.
        echo [INFO] You can manually install with: pip install -r requirements.txt
    ) else (
        echo [SUCCESS] All dependencies installed/updated.
    )
) else (
    echo [WARNING] requirements.txt not found.
)

echo.
echo [INFO] Launching PyPDF Toolbox GUI...
echo.

REM Launch the GUI launcher
"%PYTHON_EXE%" "%ROOT_DIR%\src\launcher_gui.py"

if errorlevel 1 (
    echo [ERROR] Failed to launch GUI launcher.
    pause
    exit /b 1
)

endlocal
