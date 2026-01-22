"""
PyPDF Toolbox - Silent Launcher

Double-click this file to start PyPDF Toolbox without a console window.
Windows automatically runs .pyw files with pythonw.exe.

Copyright 2025-2026 Andre Lorbach
Licensed under Apache License 2.0
"""

import sys
import os
import subprocess
from pathlib import Path


def show_error(title, message):
    """Show error dialog using tkinter."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except:
        pass


def show_info(title, message):
    """Show info dialog using tkinter."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(title, message)
        root.destroy()
    except:
        pass


def main():
    # Get the directory where this script is located
    script_dir = Path(__file__).parent.resolve()
    
    # Paths
    venv_dir = script_dir / "venv"
    launcher_script = script_dir / "src" / "launcher_gui.py"
    requirements_file = script_dir / "requirements.txt"
    
    # Determine Python paths based on OS
    if sys.platform == "win32":
        python_exe = venv_dir / "Scripts" / "python.exe"
        pythonw_exe = venv_dir / "Scripts" / "pythonw.exe"
    else:
        python_exe = venv_dir / "bin" / "python"
        pythonw_exe = python_exe  # No pythonw on Linux/Mac
    
    # Check if venv exists
    if not python_exe.exists():
        # Need to create venv - show progress
        show_info(
            "PyPDF Toolbox - First Run",
            "Setting up virtual environment...\n\n"
            "This only happens once. Please wait."
        )
        
        try:
            # Create venv
            creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            
            subprocess.run(
                [sys.executable, "-m", "venv", str(venv_dir)],
                check=True,
                creationflags=creation_flags
            )
            
            # Install requirements
            if requirements_file.exists():
                subprocess.run(
                    [str(python_exe), "-m", "pip", "install", "--upgrade", "-r", str(requirements_file)],
                    check=True,
                    creationflags=creation_flags
                )
            
        except subprocess.CalledProcessError as e:
            show_error(
                "PyPDF Toolbox - Setup Error",
                f"Failed to create virtual environment.\n\n"
                f"Error: {e}\n\n"
                f"Make sure Python is installed correctly."
            )
            return
        except Exception as e:
            show_error(
                "PyPDF Toolbox - Setup Error", 
                f"Unexpected error during setup:\n\n{e}"
            )
            return
    else:
        # Venv exists - check and update requirements
        try:
            if requirements_file.exists():
                creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                subprocess.run(
                    [str(python_exe), "-m", "pip", "install", "--upgrade", "-r", str(requirements_file)],
                    check=False,  # Don't fail if some packages can't be installed
                    creationflags=creation_flags,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
        except Exception:
            pass  # Silently fail - launcher will check dependencies on startup
    
    # Verify everything exists
    if not launcher_script.exists():
        show_error(
            "PyPDF Toolbox - Error",
            f"Launcher script not found:\n{launcher_script}"
        )
        return
    
    # Use pythonw.exe on Windows to avoid any console flash
    exe_to_use = pythonw_exe if pythonw_exe.exists() else python_exe
    
    if not exe_to_use.exists():
        show_error(
            "PyPDF Toolbox - Error",
            f"Python executable not found:\n{exe_to_use}\n\n"
            f"Please run launcher.bat to set up the environment."
        )
        return
    
    # Launch the GUI
    try:
        creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        subprocess.Popen(
            [str(exe_to_use), str(launcher_script)],
            cwd=str(script_dir),
            creationflags=creation_flags
        )
    except Exception as e:
        show_error(
            "PyPDF Toolbox - Launch Error",
            f"Failed to launch application:\n\n{e}"
        )


if __name__ == "__main__":
    main()
