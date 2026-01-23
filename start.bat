@echo off
REM ============================================
REM PyPDF Toolbox - Silent Launcher
REM Double-click this file to start PyPDF Toolbox without a console window
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

set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv
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
