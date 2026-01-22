@echo off
setlocal

set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%venv
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe
set PYTHONW_EXE=%VENV_DIR%\Scripts\pythonw.exe

REM Check if venv exists
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment not found.
    echo [INFO] Please run launcher.bat first to create the virtual environment.
    pause
    exit /b 1
)

REM Use pythonw.exe if available (no console window), otherwise fall back to python.exe
if exist "%PYTHONW_EXE%" (
    "%PYTHONW_EXE%" "%SCRIPT_DIR%src\pdf_text_extractor.py" %*
) else (
    "%PYTHON_EXE%" "%SCRIPT_DIR%src\pdf_text_extractor.py" %*
)

endlocal
