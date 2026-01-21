@echo off
:: PyPDF Toolbox - Silent Launcher
:: Double-click this file to start PyPDF Toolbox without a console window

set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%venv
set PYTHONW_EXE=%VENV_DIR%\Scripts\pythonw.exe
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe
set LAUNCHER_SCRIPT=%SCRIPT_DIR%src\launcher_gui.py

:: Check if venv exists, if not run setup first
if not exist "%PYTHON_EXE%" (
    echo First run - setting up environment...
    call "%SCRIPT_DIR%launcher.bat"
    if errorlevel 1 exit /b 1
)

:: Launch silently with pythonw.exe (no console)
if exist "%PYTHONW_EXE%" (
    start "" "%PYTHONW_EXE%" "%LAUNCHER_SCRIPT%"
) else (
    start "" "%PYTHON_EXE%" "%LAUNCHER_SCRIPT%"
)
