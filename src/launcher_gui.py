"""
PyPDF Toolbox - Launcher GUI

A slim top-bar launcher that opens PDF tools in separate windows
positioned below the launcher bar. Captures stdout/stderr from
launched tools in an expandable log panel.

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

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import sys
import subprocess
import platform
import threading
import queue
from pathlib import Path
from datetime import datetime


class PDFToolLauncher:
    """Main launcher window - slim bar at top of screen with expandable log panel"""
    
    # Configuration for launcher bar
    LAUNCHER_HEIGHT = 70  # Height of the launcher bar
    LAUNCHER_PADDING = 10  # Padding from screen edges
    LOG_PANEL_WIDTH = 500  # Width of the log panel when expanded
    
    def __init__(self, root):
        self.root = root
        self.root.title("PyPDF Toolbox")
        
        # Get root directory
        self.script_dir = Path(__file__).parent
        self.root_dir = self.script_dir.parent
        
        # Determine OS and launcher extension
        self.is_windows = platform.system() == "Windows"
        self.launcher_ext = ".bat" if self.is_windows else ".sh"
        
        # Store launcher files and running processes
        self.launchers = []
        self.running_tools = {}  # Track running tool processes
        self.output_threads = {}  # Track output reading threads
        
        # Log panel state
        self.log_panel_visible = False
        self.log_queue = queue.Queue()
        
        # Calculate screen dimensions
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        # Position launcher at top of screen
        self.position_launcher()
        
        # Setup UI
        self.setup_ui()
        
        # Scan for tools
        self.scan_launchers()
        self.populate_tools()
        
        # Make window stay on top
        self.root.attributes('-topmost', True)
        
        # Bind close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Bind window move/resize events to update tool area
        self.root.bind('<Configure>', self._on_window_configure)
        
        # Track last known position to detect actual moves
        self._last_x = self.x_pos
        self._last_y = self.y_pos
        self._last_height = self.LAUNCHER_HEIGHT
        
        # Start log queue processor
        self.process_log_queue()
    
    def position_launcher(self):
        """Position the launcher bar at the top of the screen"""
        # Calculate launcher dimensions (full width minus padding, minus log panel if visible)
        self.launcher_width = self.screen_width - (2 * self.LAUNCHER_PADDING)
        
        # Position at top of screen with padding
        self.x_pos = self.LAUNCHER_PADDING
        self.y_pos = self.LAUNCHER_PADDING
        
        # Set geometry
        self.root.geometry(f"{self.launcher_width}x{self.LAUNCHER_HEIGHT}+{self.x_pos}+{self.y_pos}")
        
        # Calculate available area for tool windows (below launcher)
        self.update_tool_area()
        
        # Prevent vertical resizing
        self.root.resizable(True, False)
    
    def update_tool_area(self):
        """Update the available area for tool windows based on current launcher position"""
        # Get current window geometry
        try:
            # Get actual window position (may differ from initial position if moved)
            current_x = self.root.winfo_x()
            current_y = self.root.winfo_y()
            current_height = self.root.winfo_height()
            current_width = self.root.winfo_width()
            
            # Update stored position
            self.x_pos = current_x
            self.y_pos = current_y
        except:
            current_height = self.LAUNCHER_HEIGHT
            current_width = self.launcher_width
        
        # Tool area starts below the launcher window with extra padding
        # Add extra offset to ensure tools don't overlap with launcher (account for window decorations)
        extra_offset = 50  # Additional spacing below launcher for window title bar clearance
        self.tool_area_y = self.y_pos + current_height + extra_offset
        self.tool_area_height = self.screen_height - self.tool_area_y - 60  # 60 for taskbar and bottom margin
        
        # Tool area X position follows launcher
        self.tool_area_x = self.x_pos
        
        # Tool area width matches launcher width (or screen width if launcher is wider)
        if self.log_panel_visible:
            # When log panel is visible, tools use left portion of screen
            self.tool_area_width = min(current_width, self.screen_width - self.x_pos - self.LAUNCHER_PADDING)
        else:
            self.tool_area_width = min(current_width, self.screen_width - self.x_pos - self.LAUNCHER_PADDING)
        
        # Ensure minimum dimensions
        self.tool_area_width = max(self.tool_area_width, 600)
        self.tool_area_height = max(self.tool_area_height, 400)
    
    def _on_window_configure(self, event):
        """Handle window move/resize events"""
        # Only respond to root window events, not child widgets
        if event.widget != self.root:
            return
        
        # Get current position
        try:
            current_x = self.root.winfo_x()
            current_y = self.root.winfo_y()
            current_height = self.root.winfo_height()
        except:
            return
        
        # Check if position or size actually changed
        if (current_x != self._last_x or 
            current_y != self._last_y or 
            current_height != self._last_height):
            
            # Update tracking variables
            self._last_x = current_x
            self._last_y = current_y
            self._last_height = current_height
            
            # Update tool area dimensions
            self.update_tool_area()
    
    def setup_ui(self):
        """Setup the launcher UI with expandable log panel"""
        # Main container
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill='both', expand=True)
        
        # Launcher bar frame
        self.launcher_frame = ttk.Frame(self.main_container, padding=5)
        self.launcher_frame.pack(side='left', fill='both', expand=True)
        
        # Left side: Title and logo area
        left_frame = ttk.Frame(self.launcher_frame)
        left_frame.pack(side='left', fill='y')
        
        title_label = ttk.Label(
            left_frame,
            text="📄 PyPDF Toolbox",
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(side='left', padx=(5, 20))
        
        # Separator
        sep = ttk.Separator(self.launcher_frame, orient='vertical')
        sep.pack(side='left', fill='y', padx=5)
        
        # Center: Tool buttons (scrollable)
        self.tools_frame = ttk.Frame(self.launcher_frame)
        self.tools_frame.pack(side='left', fill='both', expand=True)
        
        # Canvas for horizontal scrolling of tools
        self.tools_canvas = tk.Canvas(
            self.tools_frame, 
            height=50,
            highlightthickness=0
        )
        self.tools_canvas.pack(side='top', fill='x', expand=True)
        
        # Frame inside canvas for tool buttons
        self.buttons_frame = ttk.Frame(self.tools_canvas)
        self.canvas_window = self.tools_canvas.create_window(
            (0, 0), 
            window=self.buttons_frame, 
            anchor='nw'
        )
        
        # Configure scrolling
        self.buttons_frame.bind(
            '<Configure>',
            lambda e: self.tools_canvas.configure(scrollregion=self.tools_canvas.bbox('all'))
        )
        
        # Mouse wheel scrolling
        self.tools_canvas.bind('<MouseWheel>', self._on_mousewheel)
        self.buttons_frame.bind('<MouseWheel>', self._on_mousewheel)
        
        # Right side: Control buttons
        right_frame = ttk.Frame(self.launcher_frame)
        right_frame.pack(side='right', fill='y')
        
        # Separator
        sep2 = ttk.Separator(self.launcher_frame, orient='vertical')
        sep2.pack(side='right', fill='y', padx=5)
        
        # Refresh button
        refresh_btn = ttk.Button(
            right_frame,
            text="🔄",
            width=3,
            command=self.refresh_tools
        )
        refresh_btn.pack(side='left', padx=2)
        
        # Close all tools button
        close_all_btn = ttk.Button(
            right_frame,
            text="Close All",
            command=self.close_all_tools
        )
        close_all_btn.pack(side='left', padx=2)
        
        # Log panel toggle button
        self.log_toggle_btn = ttk.Button(
            right_frame,
            text="📋 Log",
            command=self.toggle_log_panel,
            width=8
        )
        self.log_toggle_btn.pack(side='left', padx=2)
        
        # Exit button
        exit_btn = ttk.Button(
            right_frame,
            text="Exit",
            command=self.on_close
        )
        exit_btn.pack(side='left', padx=2)
        
        # Status indicator (small label)
        self.status_label = ttk.Label(
            right_frame,
            text="",
            font=("Segoe UI", 8)
        )
        self.status_label.pack(side='left', padx=5)
        
        # Create log panel (initially hidden)
        self.create_log_panel()
    
    def create_log_panel(self):
        """Create the expandable log panel"""
        # Log panel frame (will be shown/hidden)
        self.log_panel = ttk.Frame(self.main_container, padding=5)
        
        # Log panel header
        log_header = ttk.Frame(self.log_panel)
        log_header.pack(fill='x', pady=(0, 5))
        
        ttk.Label(log_header, text="📋 Tool Output Log", font=("Segoe UI", 10, "bold")).pack(side='left')
        
        # Clear log button
        clear_btn = ttk.Button(log_header, text="Clear", command=self.clear_log, width=6)
        clear_btn.pack(side='right', padx=2)
        
        # Log text widget
        self.log_text = scrolledtext.ScrolledText(
            self.log_panel,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#d4d4d4",
            width=60,
            height=3
        )
        self.log_text.pack(fill='both', expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # Configure log text tags for different output types
        self.log_text.tag_configure("timestamp", foreground="#6a9955")
        self.log_text.tag_configure("tool_name", foreground="#4ec9b0", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("error", foreground="#f14c4c")
        self.log_text.tag_configure("info", foreground="#3794ff")
        self.log_text.tag_configure("separator", foreground="#808080")
    
    def toggle_log_panel(self):
        """Toggle the log panel visibility"""
        # Get current position before resize
        current_x = self.root.winfo_x()
        current_y = self.root.winfo_y()
        current_width = self.root.winfo_width()
        
        if self.log_panel_visible:
            # Hide log panel
            self.log_panel.pack_forget()
            self.log_panel_visible = False
            self.log_toggle_btn.config(text="📋 Log")
            
            # Resize launcher to slim height, keep current position and width
            self.root.geometry(f"{current_width}x{self.LAUNCHER_HEIGHT}+{current_x}+{current_y}")
        else:
            # Show log panel
            self.log_panel.pack(side='right', fill='y')
            self.log_panel_visible = True
            self.log_toggle_btn.config(text="📋 Hide")
            
            # Resize launcher to accommodate log panel, keep current position and width
            new_height = 300  # Taller when log is visible
            self.root.geometry(f"{current_width}x{new_height}+{current_x}+{current_y}")
        
        # Update stored position
        self.x_pos = current_x
        self.y_pos = current_y
        
        # Update tool area dimensions
        self.update_tool_area()
    
    def append_log(self, text, tool_name=None, is_error=False):
        """Append text to the log panel (thread-safe via queue)"""
        self.log_queue.put((text, tool_name, is_error))
    
    def process_log_queue(self):
        """Process log messages from the queue"""
        try:
            while True:
                text, tool_name, is_error = self.log_queue.get_nowait()
                self._append_log_direct(text, tool_name, is_error)
        except queue.Empty:
            pass
        
        # Schedule next check
        self.root.after(100, self.process_log_queue)
    
    def _append_log_direct(self, text, tool_name=None, is_error=False):
        """Actually append text to the log panel"""
        self.log_text.config(state=tk.NORMAL)
        
        # Add timestamp
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        
        # Add tool name if provided
        if tool_name:
            self.log_text.insert(tk.END, f"[{tool_name}] ", "tool_name")
        
        # Add text with appropriate tag
        tag = "error" if is_error else None
        self.log_text.insert(tk.END, text + "\n", tag)
        
        # Auto-scroll to bottom
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def clear_log(self):
        """Clear the log panel"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        self.tools_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def scan_launchers(self):
        """Scan for PDF tool launcher files"""
        self.launchers = []
        
        if not self.root_dir.exists():
            return
        
        # Scan for launch_*.bat or launch_*.sh files
        for launcher_file in self.root_dir.glob(f"launch_*{self.launcher_ext}"):
            name = launcher_file.stem.replace("launch_", "")
            self.launchers.append({
                "name": name,
                "display_name": self._format_tool_name(name),
                "path": launcher_file,
                "icon": self._get_tool_icon(name)
            })
        
        # Sort alphabetically
        self.launchers.sort(key=lambda x: x["name"])
    
    def _format_tool_name(self, name):
        """Format tool name for display"""
        return name.replace("_", " ").title()
    
    def _get_tool_icon(self, name):
        """Get an appropriate icon for a tool based on its name"""
        icons = {
            "split": "✂️",
            "merge": "🔗",
            "compress": "📦",
            "ocr": "👁️",
            "rotate": "🔄",
            "extract": "📤",
            "convert": "🔀",
            "watermark": "💧",
            "encrypt": "🔒",
            "decrypt": "🔓",
            "metadata": "📋",
            "preview": "👀",
            "reorder": "📑",
            "remove": "🗑️",
            "add": "➕",
            "info": "ℹ️",
        }
        
        name_lower = name.lower()
        for keyword, icon in icons.items():
            if keyword in name_lower:
                return icon
        
        return "📄"
    
    def populate_tools(self):
        """Populate the launcher with tool buttons"""
        for widget in self.buttons_frame.winfo_children():
            widget.destroy()
        
        if not self.launchers:
            placeholder = ttk.Label(
                self.buttons_frame,
                text="No PDF tools found. Add launch_*.bat/.sh files to the root directory.",
                font=("Segoe UI", 9)
            )
            placeholder.pack(padx=10, pady=15)
            return
        
        for launcher in self.launchers:
            self._create_tool_button(launcher)
        
        self.status_label.config(text=f"{len(self.launchers)} tools")
    
    def _create_tool_button(self, launcher):
        """Create a tool button"""
        btn_frame = ttk.Frame(self.buttons_frame)
        btn_frame.pack(side='left', padx=3, pady=5)
        
        btn_text = f"{launcher['icon']} {launcher['display_name']}"
        btn = ttk.Button(
            btn_frame,
            text=btn_text,
            command=lambda l=launcher: self.launch_tool(l),
            width=max(12, len(launcher['display_name']) + 4)
        )
        btn.pack()
        
        launcher['button'] = btn
    
    def _get_tool_python_script(self, launcher_name):
        """Get the Python script path for a tool based on its launcher name."""
        # Tool scripts follow the naming convention: src/pdf_<name>.py
        script_name = f"pdf_{launcher_name}.py"
        script_path = self.root_dir / "src" / script_name
        
        if script_path.exists():
            return script_path
        
        # Try alternative naming patterns
        alt_names = [
            f"pdf_{launcher_name.replace('_', '')}.py",
            f"{launcher_name}.py",
        ]
        for alt_name in alt_names:
            alt_path = self.root_dir / "src" / alt_name
            if alt_path.exists():
                return alt_path
        
        return None
    
    def launch_tool(self, launcher):
        """Launch a PDF tool without opening a console window, capturing output"""
        launcher_path = launcher["path"]
        tool_name = launcher["name"]
        display_name = launcher["display_name"]
        
        if not launcher_path.exists():
            messagebox.showerror("Error", f"Tool launcher not found:\n{launcher_path}")
            return
        
        # Check if tool is already running
        if tool_name in self.running_tools:
            process = self.running_tools[tool_name]
            if process.poll() is None:
                response = messagebox.askyesno(
                    "Tool Running",
                    f"{display_name} is already running.\n\n"
                    "Do you want to launch another instance?"
                )
                if not response:
                    return
        
        try:
            # Prepare environment
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONIOENCODING'] = 'utf-8'
            
            # Pass tool window position via environment
            env['TOOL_WINDOW_X'] = str(self.tool_area_x)
            env['TOOL_WINDOW_Y'] = str(self.tool_area_y)
            env['TOOL_WINDOW_WIDTH'] = str(self.tool_area_width)
            env['TOOL_WINDOW_HEIGHT'] = str(self.tool_area_height)
            
            # Log launch
            self.append_log(f"{'='*50}", tool_name)
            self.append_log(f"Launching {display_name}...", tool_name)
            
            # Try to launch directly with Python (no console window)
            if self.is_windows:
                # Find the Python script for this tool
                tool_script = self._get_tool_python_script(tool_name)
                
                # Get Python executable from venv
                python_exe = self.root_dir / "venv" / "Scripts" / "python.exe"
                
                if tool_script and python_exe.exists():
                    # Launch Python script directly - no console window
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    process = subprocess.Popen(
                        [str(python_exe), str(tool_script)],
                        cwd=str(self.root_dir),
                        env=env,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        stdin=subprocess.PIPE,
                        startupinfo=startupinfo,
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        text=True,
                        bufsize=1
                    )
                    self.append_log(f"Started: {tool_script.name}", tool_name)
                else:
                    # Fallback: run through batch file
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    process = subprocess.Popen(
                        ["cmd", "/c", str(launcher_path)],
                        cwd=str(self.root_dir),
                        env=env,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        stdin=subprocess.PIPE,
                        startupinfo=startupinfo,
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        text=True,
                        bufsize=1
                    )
                    self.append_log(f"Started via batch: {launcher_path.name}", tool_name)
            else:
                # Linux/Mac: try direct Python launch first
                tool_script = self._get_tool_python_script(tool_name)
                python_exe = self.root_dir / "venv" / "bin" / "python"
                
                if tool_script and python_exe.exists():
                    process = subprocess.Popen(
                        [str(python_exe), str(tool_script)],
                        cwd=str(self.root_dir),
                        env=env,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        stdin=subprocess.PIPE,
                        text=True,
                        bufsize=1
                    )
                    self.append_log(f"Started: {tool_script.name}", tool_name)
                else:
                    # Fallback: run through shell script
                    process = subprocess.Popen(
                        ["bash", str(launcher_path)],
                        cwd=str(self.root_dir),
                        env=env,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        stdin=subprocess.PIPE,
                        text=True,
                        bufsize=1
                    )
                    self.append_log(f"Started via shell: {launcher_path.name}", tool_name)
            
            # Track running process
            self.running_tools[tool_name] = process
            
            # Start thread to read output
            output_thread = threading.Thread(
                target=self._read_process_output,
                args=(process, tool_name, display_name),
                daemon=True
            )
            output_thread.start()
            self.output_threads[tool_name] = output_thread
            
            # Visual feedback
            self._flash_button(launcher)
            
        except Exception as e:
            self.append_log(f"Failed to launch: {str(e)}", tool_name, is_error=True)
            messagebox.showerror("Error", f"Failed to launch {display_name}:\n{str(e)}")
    
    def _read_process_output(self, process, tool_name, display_name):
        """Read process output in a background thread"""
        try:
            while True:
                line = process.stdout.readline()
                if not line:
                    if process.poll() is not None:
                        break
                    continue
                
                # Send output to log
                line = line.rstrip()
                if line:
                    self.append_log(line, tool_name)
            
            # Process finished
            return_code = process.poll()
            self.append_log(f"Process exited with code: {return_code}", tool_name, 
                          is_error=(return_code != 0))
            
        except Exception as e:
            self.append_log(f"Error reading output: {str(e)}", tool_name, is_error=True)
    
    def _flash_button(self, launcher):
        """Briefly highlight button to show tool was launched"""
        if 'button' in launcher:
            self.status_label.config(text=f"Launched: {launcher['display_name']}")
            self.root.after(2000, lambda: self.status_label.config(text=f"{len(self.launchers)} tools"))
    
    def refresh_tools(self):
        """Refresh the tool list"""
        self.scan_launchers()
        self.populate_tools()
    
    def close_all_tools(self):
        """Close all running tool windows"""
        closed_count = 0
        for tool_name, process in list(self.running_tools.items()):
            if process.poll() is None:
                try:
                    process.terminate()
                    self.append_log(f"Terminated", tool_name)
                    closed_count += 1
                except Exception as e:
                    self.append_log(f"Failed to terminate: {str(e)}", tool_name, is_error=True)
        
        self.running_tools.clear()
        self.status_label.config(text=f"Closed {closed_count} tools")
        self.root.after(2000, lambda: self.status_label.config(text=f"{len(self.launchers)} tools"))
    
    def on_close(self):
        """Handle window close - offer to close all tools"""
        running_count = sum(1 for p in self.running_tools.values() if p.poll() is None)
        
        if running_count > 0:
            response = messagebox.askyesnocancel(
                "Exit PyPDF Toolbox",
                f"There are {running_count} tool(s) still running.\n\n"
                "Yes = Close all tools and exit\n"
                "No = Exit launcher only (keep tools open)\n"
                "Cancel = Don't exit"
            )
            
            if response is None:  # Cancel
                return
            elif response:  # Yes - close all
                self.close_all_tools()
        
        self.root.destroy()


def main():
    root = tk.Tk()
    
    # Apply a modern theme if available
    try:
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
    except Exception:
        pass
    
    app = PDFToolLauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()
