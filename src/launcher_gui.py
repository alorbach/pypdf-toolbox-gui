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
import time
from pathlib import Path
from datetime import datetime

# HTTP requests for testing connections
try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# Azure AI configuration
try:
    # Try both import paths (when run as module or directly)
    try:
        from src.utils.azure_config import get_azure_config
        print("[INFO] Azure config loaded from src.utils.azure_config")
    except ImportError:
        try:
            from utils.azure_config import get_azure_config
            print("[INFO] Azure config loaded from utils.azure_config")
        except ImportError:
            # Try relative import
            import sys
            from pathlib import Path
            utils_path = Path(__file__).parent / "utils"
            if utils_path.exists():
                sys.path.insert(0, str(Path(__file__).parent))
                from utils.azure_config import get_azure_config
                print("[INFO] Azure config loaded via relative import")
            else:
                raise ImportError("Could not find utils.azure_config module")
    AZURE_CONFIG_AVAILABLE = True
except ImportError as e:
    print(f"[WARNING] Azure config not available: {e}")
    AZURE_CONFIG_AVAILABLE = False


class PDFToolLauncher:
    """Main launcher window - slim bar at top of screen with expandable log panel"""
    
    # Configuration for launcher bar
    LAUNCHER_HEIGHT = 70  # Height of the launcher bar
    LAUNCHER_PADDING = 10  # Padding from screen edges
    LOG_PANEL_WIDTH = 800  # Width of the log panel when expanded (increased for wider log output)
    
    def __init__(self, root):
        self.root = root
        self.root.title("PyPDF Toolbox")
        
        # Get root directory
        self.script_dir = Path(__file__).parent
        self.root_dir = self.script_dir.parent
        
        # Check if running as executable (PyInstaller)
        self.is_executable = getattr(sys, 'frozen', False)
        if self.is_executable:
            # When running as EXE, the executable is the entry point
            self.exe_path = Path(sys.executable)
            self.root_dir = self.exe_path.parent
        
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
        
        # Check and install missing critical dependencies
        self.check_dependencies()
    
    def position_launcher(self):
        """Position the launcher bar at the top of the screen"""
        # Minimum launcher width
        MIN_LAUNCHER_WIDTH = 1280
        
        # Calculate launcher dimensions (full width minus padding, minus log panel if visible)
        self.launcher_width = self.screen_width - (2 * self.LAUNCHER_PADDING)
        
        # Ensure minimum width
        if self.launcher_width < MIN_LAUNCHER_WIDTH:
            self.launcher_width = MIN_LAUNCHER_WIDTH
        
        # Position at top of screen with padding
        self.x_pos = self.LAUNCHER_PADDING
        self.y_pos = self.LAUNCHER_PADDING
        
        # Set geometry
        self.root.geometry(f"{self.launcher_width}x{self.LAUNCHER_HEIGHT}+{self.x_pos}+{self.y_pos}")
        
        # Calculate available area for tool windows (below launcher)
        self.update_tool_area()
        
        # Allow horizontal resizing, prevent vertical resizing
        self.root.resizable(True, False)
        # Set minimum width to match MIN_LAUNCHER_WIDTH
        self.root.minsize(MIN_LAUNCHER_WIDTH, self.LAUNCHER_HEIGHT)
    
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
        
        # Tool area starts below the launcher window with small padding
        # Add small offset to ensure tools don't overlap with launcher (account for window decorations)
        extra_offset = 40  # Small spacing below launcher for visual separation
        self.tool_area_y = self.y_pos + current_height + extra_offset
        # Reduce height by the gap amount so total space usage remains the same
        self.tool_area_height = self.screen_height - self.tool_area_y - 80 - extra_offset  # 80 for taskbar and bottom margin, minus gap
        
        # Tool area X position follows launcher
        self.tool_area_x = self.x_pos
        
        # Tool area width matches launcher width (or screen width if launcher is wider)
        if self.log_panel_visible:
            # When log panel is visible, tools use left portion of screen
            self.tool_area_width = min(current_width, self.screen_width - self.x_pos - self.LAUNCHER_PADDING)
        else:
            self.tool_area_width = min(current_width, self.screen_width - self.x_pos - self.LAUNCHER_PADDING)
        
        # Ensure minimum dimensions
        self.tool_area_width = max(self.tool_area_width, 1280)  # Match launcher minimum width
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
        # Main container - vertical layout
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill='both', expand=True)
        
        # Launcher bar frame - always at top
        self.launcher_frame = ttk.Frame(self.main_container, padding=5)
        self.launcher_frame.pack(side='top', fill='x', expand=False)
        
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
        
        # Horizontal scrollbar for tools
        self.tools_scrollbar = ttk.Scrollbar(
            self.tools_frame,
            orient='horizontal',
            command=self.tools_canvas.xview
        )
        self.tools_canvas.configure(xscrollcommand=self.tools_scrollbar.set)
        
        # Pack canvas and scrollbar
        self.tools_canvas.pack(side='top', fill='both', expand=True)
        self.tools_scrollbar.pack(side='bottom', fill='x')
        
        # Frame inside canvas for tool buttons
        self.buttons_frame = ttk.Frame(self.tools_canvas)
        self.canvas_window = self.tools_canvas.create_window(
            (0, 0), 
            window=self.buttons_frame, 
            anchor='nw'
        )
        
        # Configure scrolling - update scrollregion
        def configure_scroll_region(event=None):
            # Update scrollregion to include all content
            bbox = self.tools_canvas.bbox('all')
            if bbox:
                self.tools_canvas.configure(scrollregion=bbox)
            # DO NOT constrain canvas window width - let it be as wide as content for scrolling
        
        def configure_canvas_window(event=None):
            # Only update canvas window height to match canvas, but NOT width
            # Width should be unconstrained to allow horizontal scrolling
            canvas_height = self.tools_canvas.winfo_height()
            if canvas_height > 1:
                self.tools_canvas.itemconfig(self.canvas_window, height=canvas_height)
        
        self.buttons_frame.bind('<Configure>', configure_scroll_region)
        self.tools_canvas.bind('<Configure>', configure_canvas_window)
        
        # Mouse wheel scrolling (horizontal)
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
        
        # Azure AI Configuration button (always show, will handle error in dialog if needed)
        azure_btn = ttk.Button(
            right_frame,
            text="⚙️ AI Config",
            command=self.show_azure_config,
            width=8
        )
        azure_btn.pack(side='left', padx=2)
        
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
        """Create the expandable log panel with buttons above the log output"""
        # Log panel frame (will be shown/hidden) - positioned below launcher bar
        self.log_panel = ttk.Frame(self.main_container, padding=5)
        
        # Log panel header with buttons ABOVE the log output
        log_header = ttk.Frame(self.log_panel)
        log_header.pack(fill='x', pady=(0, 5))
        
        # Title label
        ttk.Label(log_header, text="📋 Tool Output Log", font=("Segoe UI", 10, "bold")).pack(side='left')
        
        # Clear log button (on the right side of header)
        clear_btn = ttk.Button(log_header, text="Clear", command=self.clear_log, width=6)
        clear_btn.pack(side='right', padx=2)
        
        # Log text widget (below the header/buttons) - wider for better readability
        self.log_text = scrolledtext.ScrolledText(
            self.log_panel,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#d4d4d4",
            width=120,  # Increased for wider log output (full width available)
            height=10   # Increased height for better visibility
        )
        # Pack log text below the header, taking up remaining space
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
            
            # Resize launcher to slim height, restore original width
            # When log is hidden, window is just the launcher bar height
            self.root.geometry(f"{current_width}x{self.LAUNCHER_HEIGHT}+{current_x}+{current_y}")
        else:
            # Show log panel - pack BELOW the launcher bar (not to the right)
            # Buttons are already above the log output in the layout
            # Pack after launcher_frame (which is already packed to 'top')
            self.log_panel.pack(side='top', fill='both', expand=True)
            self.log_panel_visible = True
            self.log_toggle_btn.config(text="📋 Hide")
            
            # Resize launcher to accommodate log panel below
            # Use full width for log panel (wider log output), increase height for log panel
            # Keep current width or use screen width minus padding
            new_width = max(current_width, self.screen_width - (2 * self.LAUNCHER_PADDING))
            new_height = 400  # Taller when log is visible (launcher bar + log panel)
            self.root.geometry(f"{new_width}x{new_height}+{current_x}+{current_y}")
            
            # Update to ensure proper sizing
            self.root.update_idletasks()
        
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
        """Handle mouse wheel scrolling - horizontal scroll for tool buttons"""
        # On Windows, delta is positive for scrolling up/right, negative for down/left
        # We want horizontal scrolling, so use Shift+Wheel or just wheel
        # Check if Shift is held for horizontal scroll, otherwise use horizontal
        if event.state & 0x1:  # Shift key held
            # Shift+Wheel = horizontal scroll
            delta = int(-1 * (event.delta / 120))
        else:
            # Regular wheel = horizontal scroll (for tool buttons area)
            delta = int(-1 * (event.delta / 120))
        
        self.tools_canvas.xview_scroll(delta, "units")
    
    def _get_tool_category(self, name):
        """Get the category for a tool based on its name."""
        name_lower = name.lower()
        
        # Define categories
        if any(keyword in name_lower for keyword in ["split", "combine", "merge", "join"]):
            return "split_merge"
        elif any(keyword in name_lower for keyword in ["extract", "text", "ocr", "read"]):
            return "extract_analyze"
        elif any(keyword in name_lower for keyword in ["convert", "transform", "rotate", "resize"]):
            return "convert_transform"
        elif any(keyword in name_lower for keyword in ["compress", "optimize", "reduce"]):
            return "optimize"
        elif any(keyword in name_lower for keyword in ["encrypt", "decrypt", "protect", "password"]):
            return "security"
        elif any(keyword in name_lower for keyword in ["watermark", "stamp", "annotate"]):
            return "annotate"
        else:
            return "other"
    
    def scan_launchers(self):
        """Scan for PDF tool launcher files"""
        self.launchers = []
        
        if not self.root_dir.exists():
            return
        
        if self.is_executable:
            # In executable mode, scan for known tools and use the executable itself
            # Map of tool names to their launcher names
            # Map of launcher names to their tool names for --tool argument
            # Key: launcher name (as it appears in launch_*.bat files)
            # Value: tool name to pass to --tool argument
            tool_mapping = {
                "pdf_ocr": "pdf_ocr",
                "pdf_text_extractor": "pdf_text_extractor",
                "pdf_visual_combiner": "pdf_visual_combiner",  # Maps to pdf_combiner module
                "pdf_splitter": "pdf_splitter",  # Maps to pdf_manual_splitter module
                "pdf_md_converter": "pdf_md_converter",
            }
            
            for launcher_name, tool_name in tool_mapping.items():
                category = self._get_tool_category(launcher_name)
                self.launchers.append({
                    "name": launcher_name,
                    "display_name": self._format_tool_name(launcher_name),
                    "path": self.exe_path,  # Use executable path
                    "icon": self._get_tool_icon(launcher_name),
                    "category": category,
                    "tool_name": tool_name  # Store actual tool name for --tool argument
                })
        else:
            # Normal mode: scan for launch_*.bat or launch_*.sh files
            for launcher_file in self.root_dir.glob(f"launch_*{self.launcher_ext}"):
                name = launcher_file.stem.replace("launch_", "")
                category = self._get_tool_category(name)
                self.launchers.append({
                    "name": name,
                    "display_name": self._format_tool_name(name),
                    "path": launcher_file,
                    "icon": self._get_tool_icon(name),
                    "category": category
                })
        
        # Sort by category, then alphabetically within category
        category_order = ["split_merge", "extract_analyze", "convert_transform", "optimize", "security", "annotate", "other"]
        self.launchers.sort(key=lambda x: (
            category_order.index(x["category"]) if x["category"] in category_order else len(category_order),
            x["name"]
        ))
    
    def _format_tool_name(self, name):
        """Format tool name for display"""
        return name.replace("_", " ").title()
    
    def _get_tool_icon(self, name):
        """Get an appropriate icon for a tool based on its name"""
        icons = {
            "split": "✂️",
            "merge": "🔗",
            "combine": "🔗",
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
        """Populate the launcher with tool buttons, grouped by category with separators"""
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
        
        # Group tools by category
        current_category = None
        category_names = {
            "split_merge": "Split & Merge",
            "extract_analyze": "Extract & Analyze",
            "convert_transform": "Convert & Transform",
            "optimize": "Optimize",
            "security": "Security",
            "annotate": "Annotate",
            "other": "Other"
        }
        
        for launcher in self.launchers:
            # Add separator and category label if category changed
            if launcher["category"] != current_category:
                # Add separator before new category (but not before first category)
                if current_category is not None:
                    self._create_category_separator()
                
                # Add category label for new category
                category_label = ttk.Label(
                    self.buttons_frame,
                    text=category_names.get(launcher["category"], "Other"),
                    font=("Segoe UI", 8, "bold"),
                    foreground="#64748b"
                )
                category_label.pack(side='left', padx=(10, 5), pady=5)
                current_category = launcher["category"]
            
            self._create_tool_button(launcher)
        
        # Update scrollregion after all buttons are added
        self.root.after_idle(lambda: self.tools_canvas.configure(scrollregion=self.tools_canvas.bbox('all')))
        
        self.status_label.config(text=f"{len(self.launchers)} tools")
    
    def _create_category_separator(self):
        """Create a visual separator between tool categories"""
        # Create a frame to hold the separator with proper sizing
        sep_frame = ttk.Frame(self.buttons_frame)
        sep_frame.pack(side='left', padx=5, pady=5)
        
        separator = ttk.Separator(sep_frame, orient='vertical')
        separator.pack(fill='y', expand=True)
        
        # Set minimum height for separator visibility
        sep_frame.config(height=30)
    
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
            
            # Check if running in executable mode
            if self.is_executable:
                # Launch the executable itself with --tool argument
                tool_name_arg = launcher.get("tool_name", tool_name)
                
                # Debug: Log what we're launching
                self.append_log(f"[DEBUG] Launching executable: {self.exe_path}", tool_name)
                self.append_log(f"[DEBUG] Tool name from launcher: {tool_name}", tool_name)
                self.append_log(f"[DEBUG] Tool name arg for --tool: {tool_name_arg}", tool_name)
                
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                cmd = [str(self.exe_path), "--tool", tool_name_arg]
                self.append_log(f"[DEBUG] Command: {' '.join(cmd)}", tool_name)
                
                process = subprocess.Popen(
                    cmd,
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
                self.append_log(f"Started: {display_name} (via executable)", tool_name)
            # Try to launch directly with Python (no console window)
            elif self.is_windows:
                # Find the Python script for this tool
                tool_script = self._get_tool_python_script(tool_name)
                
                # Get Python executable from venv
                python_exe = self.root_dir / ".venv" / "Scripts" / "python.exe"
                
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
                python_exe = self.root_dir / ".venv" / "bin" / "python"
                
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
    
    def _kill_process_tree_windows(self, pid):
        """Kill a process and all its children on Windows using taskkill"""
        try:
            # Use /T flag to kill process tree (parent and all children)
            result = subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                timeout=3,
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
            )
            return result.returncode == 0
        except Exception as e:
            print(f"taskkill failed for PID {pid}: {e}")
            return False
    
    def close_all_tools(self):
        """Close all running tool windows - forcefully terminate processes and their children"""
        closed_count = 0
        processes_to_kill = []
        
        # Collect all processes to kill
        for tool_name, process in list(self.running_tools.items()):
            if process.poll() is None:  # Process is still running
                processes_to_kill.append((tool_name, process, process.pid))
        
        if not processes_to_kill:
            return
        
        # First, try graceful termination
        for tool_name, process, pid in processes_to_kill:
            try:
                process.terminate()
                self.append_log(f"Terminating {tool_name} (PID {pid})...", tool_name)
            except Exception as e:
                self.append_log(f"Failed to terminate: {str(e)}", tool_name, is_error=True)
        
        # Wait for graceful shutdown
        time.sleep(0.8)
        self.root.update()
        
        # Force kill any processes that are still running
        for tool_name, process, pid in processes_to_kill:
            try:
                if process.poll() is None:  # Still running
                    if self.is_windows:
                        # On Windows, kill the process tree using taskkill
                        self.append_log(f"Force killing {tool_name} and children (PID {pid})...", tool_name)
                        if self._kill_process_tree_windows(pid):
                            self.append_log(f"Killed {tool_name} via taskkill /T", tool_name)
                            closed_count += 1
                        else:
                            # Fallback: try process.kill()
                            try:
                                process.kill()
                                time.sleep(0.2)
                                if process.poll() is None:
                                    # Still running, try taskkill again without /T
                                    subprocess.run(
                                        ["taskkill", "/F", "/PID", str(pid)],
                                        capture_output=True,
                                        timeout=2,
                                        creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                                    )
                                self.append_log(f"Killed {tool_name}", tool_name)
                                closed_count += 1
                            except Exception as e:
                                self.append_log(f"Failed to kill {tool_name}: {str(e)}", tool_name, is_error=True)
                    else:
                        # On Unix, kill() sends SIGKILL
                        process.kill()
                        time.sleep(0.2)
                        self.append_log(f"Force killed {tool_name}", tool_name)
                        closed_count += 1
                else:
                    # Process already terminated
                    closed_count += 1
            except (ProcessLookupError, ValueError):
                # Process already gone
                closed_count += 1
            except Exception as e:
                self.append_log(f"Error killing {tool_name}: {str(e)}", tool_name, is_error=True)
                # Still try Windows taskkill as last resort
                if self.is_windows:
                    try:
                        self._kill_process_tree_windows(pid)
                        closed_count += 1
                    except Exception:
                        pass
        
        # Clear the tracking dictionaries
        self.running_tools.clear()
        if hasattr(self, 'output_threads'):
            self.output_threads.clear()
        
        self.status_label.config(text=f"Closed {closed_count} tools")
        self.root.after(2000, lambda: self.status_label.config(text=f"{len(self.launchers)} tools"))
    
    def check_dependencies(self):
        """Check for missing critical dependencies and offer to install them."""
        missing_deps = []
        
        # Check for requests (needed for Azure config testing)
        try:
            import requests
        except ImportError:
            missing_deps.append("requests")
        
        # Check for yaml (needed for Azure config)
        try:
            import yaml
        except ImportError:
            missing_deps.append("pyyaml")
        
        if missing_deps:
            response = messagebox.askyesno(
                "Missing Dependencies",
                f"The following required packages are missing:\n\n"
                f"{', '.join(missing_deps)}\n\n"
                f"Would you like to install them now?\n\n"
                f"(This will run: pip install {' '.join(missing_deps)})"
            )
            
            if response:
                try:
                    python_exe = self.root_dir / ".venv" / "Scripts" / "python.exe" if self.is_windows else self.root_dir / ".venv" / "bin" / "python"
                    
                    if not python_exe.exists():
                        python_exe = sys.executable
                    
                    # Install missing packages
                    for dep in missing_deps:
                        self.append_log(f"Installing {dep}...", "Launcher")
                        result = subprocess.run(
                            [str(python_exe), "-m", "pip", "install", dep],
                            cwd=str(self.root_dir),
                            capture_output=True,
                            text=True,
                            timeout=60
                        )
                        if result.returncode == 0:
                            self.append_log(f"✓ {dep} installed successfully", "Launcher")
                        else:
                            self.append_log(f"✗ Failed to install {dep}: {result.stderr}", "Launcher", is_error=True)
                    
                    messagebox.showinfo(
                        "Installation Complete",
                        "Dependencies installed. Please restart the application for changes to take effect."
                    )
                except Exception as e:
                    messagebox.showerror(
                        "Installation Failed",
                        f"Failed to install dependencies:\n\n{str(e)}\n\n"
                        f"Please run manually:\n"
                        f"pip install {' '.join(missing_deps)}"
                    )
    
    def show_azure_config(self):
        """Show Azure AI configuration dialog"""
        if not AZURE_CONFIG_AVAILABLE:
            messagebox.showerror(
                "Error",
                "Azure configuration module not available.\n\n"
                "Please ensure:\n"
                "1. PyYAML is installed: pip install pyyaml\n"
                "2. The utils/azure_config.py file exists"
            )
            return
        
        try:
            # Get config instance
            config = get_azure_config(root_dir=self.root_dir)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to load Azure configuration:\n{str(e)}\n\n"
                "Please check that the configuration module is available."
            )
            return
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Azure AI Configuration")
        dialog.geometry("600x650")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main container
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill='both', expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="Azure AI Configuration",
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(
            main_frame,
            text="Configure Azure AI services for all PDF tools.\nSettings are shared across all tools.",
            font=("Segoe UI", 9),
            justify='center'
        )
        desc_label.pack(pady=(0, 15))
        
        # Azure OpenAI section
        openai_frame = ttk.LabelFrame(main_frame, text="Azure OpenAI", padding=10)
        openai_frame.pack(fill='x', pady=5)
        
        ttk.Label(openai_frame, text="Endpoint:").grid(row=0, column=0, sticky='w', pady=2)
        openai_endpoint_var = tk.StringVar(value=config.openai_endpoint)
        openai_endpoint_entry = ttk.Entry(openai_frame, textvariable=openai_endpoint_var, width=50)
        openai_endpoint_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        
        ttk.Label(openai_frame, text="API Key:").grid(row=1, column=0, sticky='w', pady=2)
        openai_key_var = tk.StringVar(value=config.openai_api_key)
        openai_key_entry = ttk.Entry(openai_frame, textvariable=openai_key_var, width=50, show='*')
        openai_key_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        
        ttk.Label(openai_frame, text="Deployment:").grid(row=2, column=0, sticky='w', pady=2)
        openai_deploy_var = tk.StringVar(value=config.openai_deployment)
        openai_deploy_entry = ttk.Entry(openai_frame, textvariable=openai_deploy_var, width=50)
        openai_deploy_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        
        ttk.Label(openai_frame, text="API Version:").grid(row=3, column=0, sticky='w', pady=2)
        openai_version_var = tk.StringVar(value=config.openai_api_version)
        openai_version_entry = ttk.Entry(openai_frame, textvariable=openai_version_var, width=50)
        openai_version_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        
        openai_frame.columnconfigure(1, weight=1)
        
        # Azure Document Intelligence section
        doc_intel_frame = ttk.LabelFrame(main_frame, text="Azure Document Intelligence", padding=10)
        doc_intel_frame.pack(fill='x', pady=5)
        
        ttk.Label(doc_intel_frame, text="Endpoint:").grid(row=0, column=0, sticky='w', pady=2)
        doc_intel_endpoint_var = tk.StringVar(value=config.doc_intel_endpoint)
        doc_intel_endpoint_entry = ttk.Entry(doc_intel_frame, textvariable=doc_intel_endpoint_var, width=50)
        doc_intel_endpoint_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        
        ttk.Label(doc_intel_frame, text="API Key:").grid(row=1, column=0, sticky='w', pady=2)
        doc_intel_key_var = tk.StringVar(value=config.doc_intel_api_key)
        doc_intel_key_entry = ttk.Entry(doc_intel_frame, textvariable=doc_intel_key_var, width=50, show='*')
        doc_intel_key_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        
        doc_intel_frame.columnconfigure(1, weight=1)
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Configuration Status", padding=10)
        status_frame.pack(fill='x', pady=5)
        
        status_text = config.get_status_text()
        status_label = ttk.Label(
            status_frame,
            text=status_text,
            font=("Consolas", 9),
            justify='left'
        )
        status_label.pack(anchor='w')
        
        # Environment variables hint
        env_hint_frame = ttk.Frame(main_frame)
        env_hint_frame.pack(fill='x', pady=5)
        
        hint_text = (
            "💡 You can also set environment variables:\n"
            "   AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, AZURE_OPENAI_DEPLOYMENT\n"
            "   AZURE_DOC_INTEL_ENDPOINT, AZURE_DOC_INTEL_API_KEY"
        )
        hint_label = ttk.Label(
            env_hint_frame,
            text=hint_text,
            font=("Segoe UI", 8),
            foreground='gray',
            justify='left'
        )
        hint_label.pack(anchor='w', padx=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        
        def test_connection():
            """Test Azure connections"""
            if not REQUESTS_AVAILABLE:
                messagebox.showerror(
                    "Error",
                    "The 'requests' library is not installed.\n\n"
                    "Install it with: pip install requests"
                )
                return
            
            # Update config with current values
            config.openai_endpoint = openai_endpoint_var.get().strip()
            config.openai_api_key = openai_key_var.get().strip()
            config.openai_deployment = openai_deploy_var.get().strip()
            config.openai_api_version = openai_version_var.get().strip()
            config.doc_intel_endpoint = doc_intel_endpoint_var.get().strip()
            config.doc_intel_api_key = doc_intel_key_var.get().strip()
            
            results = []
            
            # Test Azure OpenAI
            if config.openai_endpoint and config.openai_api_key:
                try:
                    # Test OpenAI endpoint
                    test_url = config.openai_endpoint.rstrip('/')
                    if not test_url.endswith('/openai'):
                        test_url = f"{test_url}/openai"
                    test_url = f"{test_url}/models?api-version={config.openai_api_version}"
                    
                    headers = {
                        "api-key": config.openai_api_key,
                        "Content-Type": "application/json"
                    }
                    
                    response = requests.get(test_url, headers=headers, timeout=10)
                    if response.status_code == 200:
                        results.append("✓ Azure OpenAI: Connection successful")
                    else:
                        results.append(f"✗ Azure OpenAI: Connection failed (HTTP {response.status_code})")
                except requests.exceptions.RequestException as e:
                    results.append(f"✗ Azure OpenAI: Connection failed ({str(e)})")
                except Exception as e:
                    results.append(f"✗ Azure OpenAI: Error ({str(e)})")
            else:
                results.append("○ Azure OpenAI: Not configured (missing endpoint or API key)")
            
            # Test Azure Document Intelligence
            if config.doc_intel_endpoint and config.doc_intel_api_key:
                try:
                    # Test Document Intelligence endpoint
                    test_url = config.doc_intel_endpoint.rstrip('/')
                    if not test_url.endswith('/formrecognizer'):
                        test_url = f"{test_url}/formrecognizer"
                    test_url = f"{test_url}/info?api-version=2023-07-31"
                    
                    headers = {
                        "Ocp-Apim-Subscription-Key": config.doc_intel_api_key
                    }
                    
                    response = requests.get(test_url, headers=headers, timeout=10)
                    if response.status_code == 200:
                        results.append("✓ Document Intelligence: Connection successful")
                    else:
                        results.append(f"✗ Document Intelligence: Connection failed (HTTP {response.status_code})")
                except requests.exceptions.RequestException as e:
                    results.append(f"✗ Document Intelligence: Connection failed ({str(e)})")
                except Exception as e:
                    results.append(f"✗ Document Intelligence: Error ({str(e)})")
            else:
                results.append("○ Document Intelligence: Not configured (missing endpoint or API key)")
            
            # Update status display
            status_text = config.get_status_text()
            status_label.config(text=status_text)
            
            # Show test results
            result_message = "Connection Test Results:\n\n" + "\n".join(results)
            messagebox.showinfo("Test Complete", result_message)
        
        def save_config():
            """Save configuration"""
            # Update config with current values
            config.openai_endpoint = openai_endpoint_var.get().strip()
            config.openai_api_key = openai_key_var.get().strip()
            config.openai_deployment = openai_deploy_var.get().strip()
            config.openai_api_version = openai_version_var.get().strip()
            config.doc_intel_endpoint = doc_intel_endpoint_var.get().strip()
            config.doc_intel_api_key = doc_intel_key_var.get().strip()
            
            # Check if user wants to save API keys
            has_keys = bool(config.openai_api_key or config.doc_intel_api_key)
            
            if has_keys:
                # Ask user if they want to save API keys to file
                response = messagebox.askyesno(
                    "Save API Keys?",
                    "Do you want to save API keys to the configuration file?\n\n"
                    "⚠️ Security Note: API keys will be stored in plain text.\n"
                    "The config file is in .gitignore, but keep it secure.\n\n"
                    "Yes = Save keys to file\n"
                    "No = Save endpoints only (use environment variables for keys)"
                )
                save_keys = response
            else:
                save_keys = False
            
            # Save to file (with or without keys based on user choice)
            if config.save_config(save_api_keys=save_keys):
                # Reload config to ensure it's up to date (in case of env var overrides)
                # But preserve the keys we just saved if they're not in env vars
                saved_openai_key = config.openai_api_key if not os.environ.get('AZURE_OPENAI_API_KEY') else None
                saved_doc_intel_key = config.doc_intel_api_key if not os.environ.get('AZURE_DOC_INTEL_API_KEY') else None
                
                # Reload from file
                config.reload_config()
                
                # Restore saved keys if they weren't overridden by env vars
                if saved_openai_key and not os.environ.get('AZURE_OPENAI_API_KEY'):
                    config.openai_api_key = saved_openai_key
                if saved_doc_intel_key and not os.environ.get('AZURE_DOC_INTEL_API_KEY'):
                    config.doc_intel_api_key = saved_doc_intel_key
                
                # Update status
                status_text = config.get_status_text()
                status_label.config(text=status_text)
                
                if save_keys:
                    messagebox.showinfo(
                        "Saved",
                        "Azure AI configuration saved successfully!\n\n"
                        "⚠️ API keys are stored in the config file.\n"
                        "Keep config/azure_ai.yaml secure!\n\n"
                        "All tools will now use this configuration."
                    )
                else:
                    messagebox.showinfo(
                        "Saved",
                        "Azure AI configuration saved successfully!\n\n"
                        "Note: API keys were not saved to file.\n"
                        "Use environment variables for API keys:\n"
                        "• AZURE_OPENAI_API_KEY\n"
                        "• AZURE_DOC_INTEL_API_KEY\n\n"
                        "Or click Save again and choose to save keys."
                    )
            else:
                messagebox.showerror("Error", "Failed to save configuration.")
        
        test_btn = ttk.Button(button_frame, text="Test Connection", command=test_connection)
        test_btn.pack(side='left', padx=5)
        
        save_btn = ttk.Button(button_frame, text="Save", command=save_config)
        save_btn.pack(side='left', padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=dialog.destroy)
        cancel_btn.pack(side='right', padx=5)
    
    def on_close(self):
        """Handle window close - offer to close all tools"""
        try:
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
                    # Get list of processes with PIDs before closing (since close_all_tools clears the dict)
                    processes_to_kill = [(name, proc, proc.pid) for name, proc in self.running_tools.items() if proc.poll() is None]
                    
                    # Close all tools (this will try to kill them)
                    self.close_all_tools()
                    
                    # Wait a bit for processes to actually terminate
                    time.sleep(0.5)
                    self.root.update()
                    
                    # Double-check and force kill any remaining processes by PID
                    for tool_name, process, pid in processes_to_kill:
                        try:
                            # Check if process is still running by PID
                            if process.poll() is None:
                                # On Windows, kill the entire process tree
                                if self.is_windows:
                                    self._kill_process_tree_windows(pid)
                                    time.sleep(0.2)
                                    # Try one more time if still running
                                    if process.poll() is None:
                                        subprocess.run(
                                            ["taskkill", "/F", "/T", "/PID", str(pid)],
                                            capture_output=True,
                                            timeout=2,
                                            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                                        )
                                else:
                                    process.kill()
                        except (ProcessLookupError, ValueError):
                            # Process already terminated, ignore
                            pass
                        except Exception as e:
                            # Log but don't block exit
                            print(f"Warning: Could not kill process {tool_name} (PID {pid}): {e}")
                    
                    # Final wait to ensure processes are gone
                    time.sleep(0.3)
            
            # Always destroy the window, even if there were errors
            self.root.quit()  # Stop the mainloop first
            self.root.destroy()  # Then destroy the window
            
        except Exception as e:
            # If anything goes wrong, still try to close
            print(f"Error in on_close: {e}")
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass


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
