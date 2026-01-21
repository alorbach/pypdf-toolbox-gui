# Launcher GUI Documentation

The Launcher is a slim top-bar GUI that stays at the top of the screen, providing quick access to all PyPDF Toolbox utilities.

## Features

- **Slim Design**: Minimal screen footprint with a top-bar interface
- **Tool Discovery**: Automatically discovers available tools from `launch_*.bat` / `launch_*.sh` files
- **Window Positioning**: Positions tool windows in the available screen space below the launcher

## Screenshots

*Screenshots will be added to the `screenshots/` folder.*

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main launcher bar |
| `02-tool-buttons.png` | Available tool buttons |
| `03-tool-launched.png` | Launcher with a tool window open |

## Usage

### Starting the Launcher

**Windows:**
```batch
launcher.bat
```

**Linux/macOS:**
```bash
./launcher.sh
```

### Silent Launch (No Console Window)

**Windows:**
- Double-click `PyPDF_Toolbox.pyw`
- Or use `start.bat`

### Launching Tools

Click on any tool button in the launcher bar to open that utility. The tool window will automatically position itself below the launcher.

## Technical Details

- **Source File**: `src/launcher_gui.py`
- **Framework**: tkinter with ttk widgets
- **Position**: Top of screen (fullwidth or centered)

## Environment Variables

The launcher sets these environment variables for child tools:

| Variable | Description |
|----------|-------------|
| `TOOL_WINDOW_X` | X position for tool window |
| `TOOL_WINDOW_Y` | Y position for tool window |
| `TOOL_WINDOW_WIDTH` | Available width for tool window |
| `TOOL_WINDOW_HEIGHT` | Available height for tool window |
