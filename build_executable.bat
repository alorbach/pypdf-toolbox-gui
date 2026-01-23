@echo off
REM ============================================
REM PyPDF Toolbox - Build Executable (Batch Starter)
REM ============================================
REM This batch file calls the PowerShell build script.
REM ============================================

setlocal

REM Resolve script directory without breaking on "!" in paths
for /f "delims=" %%I in ('powershell -Command "(Get-Item '%~f0').DirectoryName"') do set "ROOT_DIR=%%I"

REM Call PowerShell build script, passing all arguments through
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%\build_executable.ps1" %*

endlocal
