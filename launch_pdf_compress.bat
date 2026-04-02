@echo off
REM ============================================
REM PDF / Image Recompress Tool Launcher
REM
REM Copyright 2025-2026 Andre Lorbach
REM Licensed under the Apache License, Version 2.0
REM ============================================

setlocal

set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe
set PYTHONW_EXE=%VENV_DIR%\Scripts\pythonw.exe

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment not found.
    echo [INFO] Please run launcher.bat first to create the virtual environment.
    pause
    exit /b 1
)

if exist "%PYTHONW_EXE%" (
    "%PYTHONW_EXE%" "%SCRIPT_DIR%src\pdf_compress.py" %*
) else (
    "%PYTHON_EXE%" "%SCRIPT_DIR%src\pdf_compress.py" %*
)

endlocal
