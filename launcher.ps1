# PyPDF Toolbox - Silent Launcher (PowerShell)
# Run this script to start PyPDF Toolbox without a console window
#
# To run: Right-click -> Run with PowerShell
# Or: powershell -ExecutionPolicy Bypass -File launcher.ps1

$ErrorActionPreference = "Stop"

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Paths
$VenvDir = Join-Path $ScriptDir "venv"
$PythonExe = Join-Path $VenvDir "Scripts\python.exe"
$PythonwExe = Join-Path $VenvDir "Scripts\pythonw.exe"
$LauncherScript = Join-Path $ScriptDir "src\launcher_gui.py"
$RequirementsFile = Join-Path $ScriptDir "requirements.txt"

# Check if venv exists
if (-not (Test-Path $PythonExe)) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    
    # Create venv
    python -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to create virtual environment." -ForegroundColor Red
        Write-Host "Make sure Python is installed and in PATH." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Install requirements
    if (Test-Path $RequirementsFile) {
        Write-Host "Installing dependencies..." -ForegroundColor Yellow
        & $PythonExe -m pip install -q -r $RequirementsFile
    }
    
    Write-Host "Setup complete!" -ForegroundColor Green
}

# Launch the GUI silently using pythonw.exe (no console window)
if (Test-Path $PythonwExe) {
    Start-Process -FilePath $PythonwExe -ArgumentList "`"$LauncherScript`"" -WorkingDirectory $ScriptDir -WindowStyle Hidden
} else {
    # Fallback to python.exe if pythonw.exe doesn't exist
    Start-Process -FilePath $PythonExe -ArgumentList "`"$LauncherScript`"" -WorkingDirectory $ScriptDir -WindowStyle Hidden
}
