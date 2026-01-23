<#
PyPDF Toolbox - Build Executable Script (PowerShell)
Creates a single Windows executable containing the launcher and all PDF tools.

Usage:
  .\build_executable.ps1           # Production build (no console window)
  .\build_executable.ps1 --debug   # Debug build (console window enabled)
  .\build_executable.ps1 -CI       # CI build (non-interactive)
#>

param(
    [switch]$debug,
    [switch]$ci
)

$rootDir = $PSScriptRoot
$venvDir = Join-Path $rootDir ".venv"
$pythonExe = Join-Path $venvDir "Scripts\python.exe"
$buildDir = Join-Path $rootDir "build"
$distDir = Join-Path $rootDir "dist"
$specFile = Join-Path $rootDir "PyPDF_Toolbox.spec"
$tempSpec = $null
$isCI = $ci -or ($env:CI -eq "true")

function Prompt-Exit($message) {
    if (-not $isCI) {
        Read-Host $message
    }
}

if ($debug) {
    Write-Host ""
    Write-Host "============================================"
    Write-Host "PyPDF Toolbox - Build Executable (DEBUG MODE)"
    Write-Host "============================================"
    Write-Host "[INFO] Console window will be enabled in executable"
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "============================================"
    Write-Host "PyPDF Toolbox - Build Executable"
    Write-Host "============================================"
    Write-Host ""
}

Push-Location $rootDir
try {
    & python --version > $null 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Python is not installed or not in PATH."
        Write-Host "[INFO] Please install Python 3.8+ and try again."
        Prompt-Exit "Press Enter to exit"
        exit 1
    }

    Write-Host "[INFO] Python found."
    Write-Host ""

    if (-not (Test-Path $venvDir)) {
        Write-Host "[INFO] Creating virtual environment..."
        & python -m venv $venvDir
        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ERROR] Failed to create virtual environment."
            Prompt-Exit "Press Enter to exit"
            exit 1
        }
        Write-Host "[SUCCESS] Virtual environment created."
        Write-Host ""
    }

    if (-not (Test-Path $pythonExe)) {
        Write-Host "[ERROR] Python executable not found in virtual environment."
        Prompt-Exit "Press Enter to exit"
        exit 1
    }

    Write-Host "[INFO] Upgrading pip..."
    & $pythonExe -m pip install --upgrade pip --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[WARNING] Failed to upgrade pip, continuing anyway..."
    }

    Write-Host "[INFO] Installing requirements..."
    if (Test-Path "requirements.txt") {
        & $pythonExe -m pip install --upgrade -r "requirements.txt"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ERROR] Failed to install requirements."
            Prompt-Exit "Press Enter to exit"
            exit 1
        }
    } else {
        Write-Host "[ERROR] requirements.txt not found."
        Prompt-Exit "Press Enter to exit"
        exit 1
    }

    Write-Host ""
    Write-Host "[INFO] Starting PyInstaller build..."
    Write-Host ""

    if (-not (Test-Path $buildDir)) {
        New-Item -ItemType Directory -Path $buildDir | Out-Null
    }

    if (-not (Test-Path $specFile)) {
        Write-Host "[ERROR] Spec file not found: PyPDF_Toolbox.spec"
        Write-Host "[INFO] Please ensure PyPDF_Toolbox.spec exists in the project root directory."
        Write-Host "[INFO] Current directory: $rootDir"
        Prompt-Exit "Press Enter to exit"
        exit 1
    }

    if (Test-Path $distDir) {
        Write-Host "[INFO] Cleaning previous build..."
        Remove-Item -Recurse -Force $distDir
    }
    if (Test-Path (Join-Path $buildDir "PyPDF_Toolbox")) {
        Remove-Item -Recurse -Force (Join-Path $buildDir "PyPDF_Toolbox")
    }

    Write-Host "[INFO] Building executable (this may take several minutes)..."

    $consoleValue = if ($debug) { "True" } else { "False" }
    $tempSpec = Join-Path $rootDir "PyPDF_Toolbox.temp.spec"
    $specContent = Get-Content $specFile -Raw
    $specContent = $specContent -replace "console\s*=\s*(True|False)", "console=$consoleValue"
    Set-Content -Path $tempSpec -Value $specContent -Encoding UTF8

    $pyinstallerArgs = @(
        $tempSpec,
        "--clean",
        "--noconfirm",
        "--workpath", $buildDir,
        "--distpath", $distDir
    )

    if ($debug) {
        Write-Host "[INFO] Building with console window enabled (debug mode)..."
        $pyinstallerArgs += @("--log-level=DEBUG")
    } else {
        Write-Host "[INFO] Building without console window (production mode)..."
    }

    & $pythonExe -m PyInstaller @pyinstallerArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "[ERROR] Build failed!"
        Write-Host "[INFO] Check the output above for error messages."
        Prompt-Exit "Press Enter to exit"
        exit 1
    }

    Write-Host ""
    Write-Host "============================================"
    Write-Host "Build Complete!"
    Write-Host "============================================"
    Write-Host ""
    Write-Host "[SUCCESS] Executable created successfully!"
    Write-Host ""
    Write-Host "Output location: $distDir\PyPDF_Toolbox.exe"
    Write-Host ""

    if (Test-Path (Join-Path $distDir "PyPDF_Toolbox.exe")) {
        Write-Host "[INFO] Executable size:"
        Get-Item (Join-Path $distDir "PyPDF_Toolbox.exe") | Select-Object Name, Length
        Write-Host ""
        Write-Host "[INFO] You can now distribute the executable."
        Write-Host "[INFO] The executable is standalone and includes all dependencies."
        Write-Host ""
    } else {
        Write-Host "[WARNING] Executable file not found in expected location."
        Write-Host "[INFO] Check the dist directory for output files."
        Write-Host ""
    }

    if (-not $isCI) {
        $testExe = Read-Host "Do you want to test the executable now? (Y/N)"
        if ($testExe -match "^[Yy]") {
            Write-Host ""
            Write-Host "[INFO] Launching executable..."
            Start-Process -FilePath (Join-Path $distDir "PyPDF_Toolbox.exe")
        }
    }

    Write-Host ""
    Write-Host "[INFO] Build process completed."
    Write-Host ""
    Prompt-Exit "Press Enter to exit"
} finally {
    if ($tempSpec -and (Test-Path $tempSpec)) {
        Remove-Item -Force $tempSpec
    }
    Pop-Location
}
