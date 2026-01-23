"""
PyPDF Toolbox - Unified Entry Point

This script serves as the unified entry point for both the launcher and all tools.
When packaged as an executable, this allows a single EXE to run as:
- Launcher (default, no arguments)
- Any tool (with --tool <toolname> argument)

Copyright 2025-2026 Andre Lorbach

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import sys
import os
from pathlib import Path

# Handle both script mode and executable mode (PyInstaller)
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    _script_dir = Path(sys.executable).parent
    # In executable mode, modules are bundled by PyInstaller
    # They should be directly importable (PyInstaller adds them to sys.path)
    _src_dir = _script_dir
    print(f"[DEBUG] Frozen mode - executable dir: {_script_dir}")
    print(f"[DEBUG] sys.path: {sys.path[:5]}...")  # Show first 5 entries
else:
    # Running as script
    _script_dir = Path(__file__).parent
    _src_dir = _script_dir / "src"
    if _src_dir.exists():
        sys.path.insert(0, str(_src_dir))
        print(f"[DEBUG] Script mode - added to path: {_src_dir}")

# Tool name to module mapping
TOOL_MODULES = {
    "pdf_ocr": "pdf_ocr",
    "pdf_text_extractor": "pdf_text_extractor",
    "pdf_combiner": "pdf_combiner",
    "pdf_visual_combiner": "pdf_combiner",  # Alias
    "pdf_splitter": "pdf_manual_splitter",
    "pdf_manual_splitter": "pdf_manual_splitter",
    "pdf_md_converter": "pdf_md_converter",
}

# Map launcher names to tool names
LAUNCHER_TO_TOOL = {
    "pdf_ocr": "pdf_ocr",
    "pdf_text_extractor": "pdf_text_extractor",
    "pdf_visual_combiner": "pdf_combiner",
    "pdf_combiner": "pdf_combiner",
    "pdf_splitter": "pdf_manual_splitter",
    "pdf_md_converter": "pdf_md_converter",
}


def run_launcher():
    """Run the main launcher GUI"""
    try:
        print("[DEBUG] Importing launcher_gui...")
        print(f"[DEBUG] Current sys.path entries: {len(sys.path)}")
        print(f"[DEBUG] Checking if launcher_gui is importable...")
        
        # Try to import
        import importlib.util
        if getattr(sys, 'frozen', False):
            # In frozen mode, try direct import
            try:
                from launcher_gui import main as launcher_main
            except ImportError:
                # Try importing from src package
                import src.launcher_gui as launcher_gui
                launcher_main = launcher_gui.main
        else:
            from launcher_gui import main as launcher_main
        
        print("[DEBUG] Successfully imported launcher_gui")
        print("[DEBUG] Starting launcher GUI...")
        launcher_main()
    except ImportError as e:
        print(f"[ERROR] Failed to import launcher: {e}")
        print(f"[INFO] Looking for launcher in: {_src_dir}")
        print(f"[DEBUG] sys.path: {sys.path}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Failed to run launcher: {e}")
        import traceback
        traceback.print_exc()
        raise


def run_tool(tool_name):
    """Run a specific tool by name"""
    try:
        print(f"[DEBUG] Running tool: {tool_name}")
        
        # CRITICAL: Remove --tool arguments from sys.argv before importing tools
        # Tools parse argparse at module level, so they need clean argv
        original_argv = sys.argv[:]
        try:
            # Remove '--tool' and tool_name from sys.argv
            if len(sys.argv) >= 3 and sys.argv[1] == "--tool":
                # Remove both '--tool' and the tool name
                sys.argv = [sys.argv[0]] + sys.argv[3:]
                print(f"[DEBUG] Cleaned sys.argv: {sys.argv}")
            elif len(sys.argv) >= 2 and sys.argv[1] == "--tool":
                # Only --tool, no tool name (shouldn't happen, but handle it)
                sys.argv = [sys.argv[0]] + sys.argv[2:]
            
            # Normalize tool name
            tool_name = tool_name.lower().replace("-", "_")
            print(f"[DEBUG] Normalized tool name: {tool_name}")
            
            # Map launcher name to tool name if needed
            if tool_name in LAUNCHER_TO_TOOL:
                tool_name = LAUNCHER_TO_TOOL[tool_name]
                print(f"[DEBUG] Mapped to tool: {tool_name}")
            
            # Get module name
            module_name = TOOL_MODULES.get(tool_name, tool_name)
            print(f"[DEBUG] Module name: {module_name}")
            
            if module_name not in TOOL_MODULES.values():
                print(f"[ERROR] Unknown tool: {tool_name}")
                print(f"[INFO] Available tools: {', '.join(sorted(set(TOOL_MODULES.values())))}")
                sys.exit(1)
            
            # Import and run the tool (now with clean sys.argv)
            print(f"[DEBUG] Importing module: {module_name}")
            if module_name == "pdf_ocr":
                from pdf_ocr import main as tool_main
            elif module_name == "pdf_text_extractor":
                from pdf_text_extractor import main as tool_main
            elif module_name == "pdf_combiner":
                from pdf_combiner import main as tool_main
            elif module_name == "pdf_manual_splitter":
                from pdf_manual_splitter import main as tool_main
            elif module_name == "pdf_md_converter":
                from pdf_md_converter import main as tool_main
            else:
                print(f"[ERROR] Tool module '{module_name}' not implemented in entry point")
                sys.exit(1)
            
            print(f"[DEBUG] Starting tool: {module_name}")
            print(f"[DEBUG] Calling tool_main() for {module_name}...")
            try:
                tool_main()
                print(f"[DEBUG] tool_main() completed for {module_name}")
            except Exception as tool_error:
                print(f"[ERROR] Exception in tool_main() for {module_name}: {tool_error}")
                import traceback
                traceback.print_exc()
                raise
        finally:
            # Restore original argv (though we're in a subprocess, so this might not matter)
            sys.argv = original_argv
        
    except ImportError as e:
        print(f"[ERROR] Failed to import tool '{tool_name}': {e}")
        print(f"[INFO] Looking for module: {module_name}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Failed to run tool '{tool_name}': {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def main():
    """Main entry point - route to launcher or tool based on arguments"""
    try:
        # Check if running as executable (PyInstaller)
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            exe_dir = Path(sys.executable).parent
            print(f"[DEBUG] Running as executable from: {exe_dir}")
            # Update paths for executable mode
            os.chdir(exe_dir)
            print(f"[DEBUG] Changed working directory to: {os.getcwd()}")
        else:
            print(f"[DEBUG] Running as script from: {Path(__file__).parent}")
        
        # Parse command line arguments
        if len(sys.argv) > 1:
            if sys.argv[1] == "--tool" and len(sys.argv) > 2:
                # Run specific tool
                tool_name = sys.argv[2]
                print(f"[DEBUG] main() called with --tool {tool_name}")
                print(f"[DEBUG] Full sys.argv: {sys.argv}")
                run_tool(tool_name)
                print(f"[DEBUG] run_tool() completed, exiting...")
                sys.exit(0)  # Explicitly exit after tool completes
            elif sys.argv[1] in ["-h", "--help"]:
                # Show help
                print("PyPDF Toolbox - Unified Entry Point")
                print()
                print("Usage:")
                print("  PyPDF_Toolbox.exe                    # Run launcher (default)")
                print("  PyPDF_Toolbox.exe --tool <toolname>  # Run specific tool")
                print()
                print("Available tools:")
                for tool in sorted(set(TOOL_MODULES.values())):
                    print(f"  - {tool}")
                print()
                print("Examples:")
                print("  PyPDF_Toolbox.exe --tool pdf_ocr")
                print("  PyPDF_Toolbox.exe --tool pdf_text_extractor")
            else:
                # Unknown argument, show help
                print(f"[WARNING] Unknown argument: {sys.argv[1]}")
                print("[INFO] Use --help for usage information")
                print("[INFO] Running launcher by default...")
                run_launcher()
        else:
            # No arguments - run launcher
            print("[DEBUG] No arguments provided, running launcher...")
            run_launcher()
    
    except Exception as e:
        # Catch any errors and display them
        import traceback
        error_msg = f"""
[ERROR] Failed to start PyPDF Toolbox

Error: {str(e)}

Traceback:
{traceback.format_exc()}

Please report this error with the above information.
"""
        print(error_msg)
        
        # Try to show error dialog if possible
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("PyPDF Toolbox - Startup Error", 
                               f"Failed to start application:\n\n{str(e)}\n\n"
                               "Check console output for details.")
            root.destroy()
        except:
            pass
        
        # Wait for user input before closing (if console is available)
        if getattr(sys, 'frozen', False):
            try:
                input("\nPress Enter to exit...")
            except:
                pass
        
        sys.exit(1)


if __name__ == "__main__":
    main()
