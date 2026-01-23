# Installation Guide - Windows Executable

This guide explains how to build and use PyPDF Toolbox as a single Windows executable that doesn't require Python to be installed.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Building the Executable](#building-the-executable)
3. [Using the Executable](#using-the-executable)
4. [Troubleshooting](#troubleshooting)
5. [Distribution](#distribution)

## Prerequisites

Before building the executable, ensure you have:

### Required

- **Python 3.8 or higher** installed on your system
  - **For Windows 10/11**: Install from [Microsoft Store](https://apps.microsoft.com/store/detail/python-311/9NRWMJP3717K) (recommended)
    - Search for "Python 3.11" or "Python 3.12" in Microsoft Store
    - Click "Get" or "Install"
    - Automatically added to PATH, no manual configuration needed
  - **Alternative**: Download from [python.org](https://www.python.org/downloads/)
    - During installation, check "Add Python to PATH"
  - Verify installation: Open Command Prompt and run `python --version`

- **Windows 10 or later**

- **Internet connection** (for downloading dependencies during build)

- **At least 1 GB free disk space** (for build process and final executable)

### Optional (for specific features)

- **Tesseract OCR** (for OCR functionality)
  - Download from [GitHub releases](https://github.com/UB-Mannheim/tesseract/wiki)
  - Install and add to PATH
  - Required for: PDF OCR tool, OCR-based text extraction

- **Poppler** (for PDF to image conversion)
  - Download from [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)
  - Extract and add `bin` folder to PATH
  - Required for: PDF thumbnail generation

## Building the Executable

### Step 1: Prepare the Project

1. **Open Command Prompt**
   - Press `Win + R`, type `cmd`, press Enter
   - Or search for "Command Prompt" in Start menu

2. **Navigate to the project directory**
   ```batch
   cd "C:\Projects\pypdf-toolbox-gui"
   ```
   *(Replace with your actual project path - typically where you cloned or extracted the project)*

3. **Verify you're in the right location**
   ```batch
   dir build_executable.bat
   ```
   You should see the file listed. If not, check your path.

### Step 2: Run the Build Script

1. **Execute the build script**
   
   The batch file is a thin starter that calls the PowerShell build script.
   
   **Production build (no console window):**
   ```batch
   build_executable.bat
   ```
   or directly:
   ```powershell
   .\build_executable.ps1
   ```
   
   **Debug build (with console window for troubleshooting):**
   ```batch
   build_executable.bat --debug
   ```
   or directly:
   ```powershell
   .\build_executable.ps1 --debug
   ```
   
   **When to use `--debug`:**
   - When troubleshooting startup issues
   - When you need to see error messages and debug output
   - During development and testing
   
   **Production builds (default):**
   - No console window appears when running the executable
   - Cleaner user experience
   - Recommended for distribution

2. **Wait for the build process**
   - The script will:
     - Create a virtual environment (`.venv/`) if it doesn't exist
     - Install PyInstaller and all dependencies
     - Build the executable (this takes 5-10 minutes)
   - You'll see progress messages in the console
   - **Don't close the window** during the build

3. **Build completion**
   - When finished, you'll see: `[SUCCESS] Executable created successfully!`
   - The executable will be in: `dist\PyPDF_Toolbox.exe`

### Step 3: Verify the Build

1. **Check the output**
   ```batch
   dir dist\PyPDF_Toolbox.exe
   ```
   You should see the executable file (typically 100-200 MB)

2. **Test the executable** (optional)
   - The build script will ask if you want to test it
   - Type `Y` and press Enter to launch it
   - Or manually double-click `dist\PyPDF_Toolbox.exe`

## Using the Executable

### Running the Launcher

**Method 1: Double-click**
- Navigate to the `dist` folder
- Double-click `PyPDF_Toolbox.exe`
- The launcher GUI will appear at the top of your screen

**Method 2: Command line**
```batch
cd dist
PyPDF_Toolbox.exe
```

**Method 3: Create a shortcut**
1. Right-click `PyPDF_Toolbox.exe`
2. Select "Create shortcut"
3. Move the shortcut to Desktop or Start Menu
4. Double-click the shortcut to launch

### Using the Launcher

1. **Launch a tool**
   - Click any tool button in the launcher
   - The tool window will open below the launcher

2. **Configure Azure AI** (optional)
   - Click the **"⚙️ Azure"** button
   - Enter your Azure credentials
   - Click "Save"
   - All tools will use these settings

3. **View tool output**
   - Click the **"📋 Log"** button to expand the log panel
   - See real-time output from running tools
   - Click again to collapse

### Running Tools Directly

You can run individual tools without the launcher:

```batch
# PDF OCR Tool
PyPDF_Toolbox.exe --tool pdf_ocr

# PDF Text Extractor
PyPDF_Toolbox.exe --tool pdf_text_extractor

# PDF Combiner
PyPDF_Toolbox.exe --tool pdf_combiner

# PDF Splitter
PyPDF_Toolbox.exe --tool pdf_splitter

# Markdown Converter
PyPDF_Toolbox.exe --tool pdf_md_converter
```

### Getting Help

```batch
PyPDF_Toolbox.exe --help
```

## Troubleshooting

### Build Issues

#### "Python is not installed or not in PATH"

**Solution:**
1. **For Windows 10/11** (Recommended):
   - Install Python from [Microsoft Store](https://apps.microsoft.com/store/detail/python-311/9NRWMJP3717K)
   - Search for "Python 3.11" or "Python 3.12" in Microsoft Store
   - Click "Get" or "Install"
   - Python is automatically added to PATH
   - Restart Command Prompt after installation
   
2. **Alternative**:
   - Install Python from [python.org](https://www.python.org/downloads/)
   - During installation, check **"Add Python to PATH"**
   - Restart Command Prompt after installation
   
3. **Verify installation**:
   ```batch
   python --version
   ```
   Should show Python version (e.g., "Python 3.11.5")

#### "Failed to create virtual environment"

**Solution:**
1. Check disk space (need at least 500 MB free)
2. Check if antivirus is blocking file creation
3. Try running Command Prompt as Administrator
4. Check Python installation: `python -m venv --help`

#### "Failed to install dependencies"

**Solution:**
1. Check internet connection
2. Try updating pip: `python -m pip install --upgrade pip`
3. Check if firewall is blocking pip
4. Try installing manually: `python -m pip install pyinstaller`

#### Build takes too long or hangs

**Solution:**
1. This is normal - first build can take 10-15 minutes
2. Ensure stable internet connection
3. Disable antivirus temporarily during build
4. Close other applications to free up resources

### Runtime Issues

#### Executable doesn't start

**Solution:**
1. **Check Windows Defender**
   - Open Windows Security
   - Go to "Virus & threat protection"
   - Click "Manage settings"
   - Add exception for `dist\PyPDF_Toolbox.exe`

2. **Check for error messages**
   - Run from Command Prompt to see errors:
     ```batch
     cd dist
     PyPDF_Toolbox.exe
     ```

3. **Check dependencies**
   - Some tools need external programs:
     - OCR: Install Tesseract OCR
     - PDF to Image: Install Poppler

#### "Module not found" errors

**Solution:**
1. Rebuild the executable
2. Ensure all dependencies in `requirements.txt` installed correctly
3. Check `PyPDF_Toolbox.spec` (in project root) for missing hidden imports

#### Tools don't launch from launcher

**Solution:**
1. Check if executable is in `dist` folder
2. Ensure you're running the executable (not the Python script)
3. Try running tool directly: `PyPDF_Toolbox.exe --tool pdf_ocr`

#### Large file size

**This is normal!** The executable includes:
- Python interpreter (~30 MB)
- All Python libraries (~100-150 MB)
- All tool modules
- Tkinter GUI framework

**To reduce size:**
- The build already uses UPX compression
- File size is typically 100-200 MB
- This is acceptable for a standalone executable

## Distribution

### Sharing the Executable

The `PyPDF_Toolbox.exe` file is **completely standalone** and can be:

1. **Copied to any Windows computer**
   - No Python installation required
   - No dependencies to install
   - Just copy and run

2. **Distributed via:**
   - USB drive
   - Network share
   - Cloud storage (Google Drive, Dropbox, etc.)
   - Email (if file size allows)

3. **Requirements on target computer:**
   - Windows 10 or later
   - For OCR features: Tesseract OCR (if using OCR tools)
   - For PDF to Image: Poppler (if using thumbnail features)

### Creating a Distribution Package

1. **Create a folder** (e.g., `PyPDF_Toolbox_v1.0`)

2. **Copy the executable**
   ```
   PyPDF_Toolbox_v1.0/
   └── PyPDF_Toolbox.exe
   ```

3. **Optional: Add README**
   - Create a `README.txt` with usage instructions
   - Include information about optional dependencies

4. **Optional: Create installer**
   - Use tools like Inno Setup or NSIS
   - Create a proper Windows installer
   - Add Start Menu shortcuts
   - Add desktop shortcut

### Updating the Executable

When you update the source code:

1. **Rebuild the executable**
   ```batch
   build_executable.bat
   ```

2. **Replace the old executable**
   - Delete old `dist\PyPDF_Toolbox.exe`
   - New executable will be in same location

3. **Test before distributing**
   - Always test the new executable
   - Verify all tools work correctly

## Next Steps

After building and testing:

1. **Create shortcuts** for easy access
2. **Configure Azure AI** if you plan to use AI features
3. **Install optional dependencies** (Tesseract, Poppler) if needed
4. **Read tool documentation** in the `doc/` folder

## Advanced Configuration

### Build Process Details

The build script (`build_executable.bat`) automatically:

1. **Uses or creates a virtual environment** (`.venv/`)
   - Uses the existing `.venv/` directory if it exists
   - Creates it if it doesn't exist
   - Same environment used for development and building

2. **Installs all dependencies**
   - PyInstaller (for creating executables)
   - All runtime dependencies from `requirements.txt`

3. **Runs PyInstaller**
   - Uses the spec file: `PyPDF_Toolbox.spec` (in project root)
   - Bundles all Python modules and dependencies
   - Creates a single-file executable

4. **Outputs the executable**
   - Location: `dist/PyPDF_Toolbox.exe`
   - Size: Typically 100-200 MB (includes all dependencies)

### Architecture

#### Unified Entry Point

The `PyPDF_Toolbox.py` script serves as the unified entry point:
- **No arguments**: Runs the launcher GUI
- **`--tool <name>`**: Runs the specified tool directly

#### Executable Mode Detection

When running as an executable:
- The launcher detects executable mode via `sys.frozen`
- Tools are launched by running the executable itself with `--tool` arguments
- All modules are bundled and available for import

### File Structure

```
pypdf-toolbox-gui/
├── PyPDF_Toolbox.py          # Unified entry point
├── PyPDF_Toolbox.spec       # PyInstaller spec file
├── build_executable.bat      # Build script
├── requirements.txt          # Python dependencies (includes PyInstaller)
└── dist/
    └── PyPDF_Toolbox.exe     # Output executable (after build)
```

### Custom Icon

1. Create or obtain an `.ico` file
2. Edit `PyPDF_Toolbox.spec` (in project root)
3. Update the `icon` parameter:
   ```python
   icon='path/to/icon.ico',
   ```
4. Rebuild

### Debug Mode

To see console output for debugging, use the `--debug` flag when building:

```batch
build_executable.bat --debug
```

Or run the PowerShell script directly:

```powershell
.\build_executable.ps1 --debug
```

This will:
- Enable console window in the executable
- Show all debug messages and error output
- Help troubleshoot startup issues and errors

**Note**: The debug build shows a console window when running the executable. For production builds, omit the `--debug` flag to create a clean GUI-only executable.

### Excluding Modules

To reduce file size by excluding unused modules:

1. Edit `PyPDF_Toolbox.spec` (in project root)
2. Add to the `excludes` list:
   ```python
   excludes=[
       'matplotlib',
       'numpy',
       # ... add more
   ],
   ```

### Build Artifacts

The build process creates:
- `build/PyPDF_Toolbox/` - Temporary build files (can be deleted)
- `dist/PyPDF_Toolbox.exe` - **Final executable (keep this)**

**Note**: The `.venv/` directory is shared between development and building. Do not delete it unless you want to recreate it.

You can safely delete `build/PyPDF_Toolbox/` after a successful build.

### Updating the Build

When adding new tools or dependencies:

1. **Add tool to entry point**: Update `PyPDF_Toolbox.py` with new tool mapping
2. **Update launcher**: Ensure launcher knows about the new tool
3. **Update spec file**: Add any new hidden imports if needed
4. **Rebuild**: Run `build_executable.bat` again

## Additional Resources

- **Tool Documentation**: See [doc/](doc/) for individual tool guides
- **Main README**: See [README.md](README.md) for project overview

## Support

If you encounter issues:

1. Check this troubleshooting section
2. Review the advanced configuration section above for technical details
3. Check the project's issue tracker (if available)
4. Verify all prerequisites are met

---

**Note**: The executable is built specifically for Windows. For Linux or macOS, you'll need to use the Python scripts directly or create platform-specific builds.
