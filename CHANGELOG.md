# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2026-01-23

### Added

#### Markdown Converter Enhancements

* **Encoding Detection**: Automatic detection of file encoding (UTF-16, UTF-8-BOM, and fallback encodings) for better file compatibility
* **WeasyPrint PDF Engine**: Added WeasyPrint as a PDF generation option with automatic GTK DLL detection for Windows
* **Enhanced PDF Styling**: Improved PDF typography with page break controls, orphans/widows handling, and better formatting
* **Settings System**: New global settings dialog (⚙️ Settings button) with persistent configuration saved to `config/md_converter_settings.json`
* **Optional Features** (configurable via settings):
  - **Page Break Support**: HTML page breaks, `<!-- PAGEBREAK -->` comments, and automatic page breaks before numbered headings
  - **Content Preprocessing**: Automatic removal of horizontal rules and line break normalization
  - **Advanced DOCX Features**: Bookmarks, internal/external hyperlinks, inline formatting (bold, italic, code), narrow margins option, and language settings
  - **Font Size Options**: Configurable font sizes (9pt, 10pt, 11pt, 12pt, 14pt) with proportional heading sizes
* **Auto-Open Feature**: Automatically open generated PDF/DOCX files after export (enabled by default, configurable in settings)
* **UTF-8 Icon Support**: Full support for UTF-8 icons and emojis (✅, ⭐, 🔧, 📋, etc.) in PDF exports
* **Smart Filename Detection**: Export dialogs automatically suggest filename based on loaded markdown file

### Changed

* **Default PDF Engine**: Browser engine is now the default (was WeasyPrint/ReportLab)
* **Default Font Size**: Changed from 11pt to 10pt
* **Export Behavior**: Removed confirmation dialogs when auto-open is enabled for smoother workflow

### Fixed

* **UTF-8 Character Support**: Fixed PDF export to properly preserve and display UTF-8 icons and special characters
* **Font Registration**: Enhanced font registration with Unicode support for better character rendering

### Technical

* Added `SettingsManager` class for centralized configuration management
* Enhanced `MarkdownConverter` class with preprocessing and page break handling
* Added helper methods for advanced DOCX features (bookmarks, hyperlinks, formatting)
* Improved WeasyPrint integration with automatic GTK DLL detection
* Updated all CSS presets with enhanced PDF styling rules

## [1.0.0] - 2026-01-23

### 🎉 First Release

This is the initial release of PyPDF Toolbox GUI - a collection of PDF utility tools with a unified graphical launcher interface.

### ✨ Features

* **Slim Launcher Bar**: Compact, always-on-top toolbar that stays at the top of your screen
* **Individual Tool Windows**: Each PDF tool opens in its own window positioned below the launcher
* **Global Azure AI Configuration**: Configure Azure AI services once in the launcher, all tools use the same settings
* **No Python Required**: Standalone Windows executable - just download and run!

### 📦 Included Tools

* **PDF Splitter**: Manually split PDFs by selecting split points with visual thumbnail preview
* **PDF Visual Combiner**: Combine multiple PDFs by visually selecting individual pages from thumbnails
* **PDF Text Extractor**: Extract text from PDFs using Python, OCR, or Azure AI
* **PDF OCR**: Add OCR text recognition to PDFs and convert images to searchable PDFs
* **Markdown Converter**: Convert Markdown to PDF/DOCX with live HTML preview and style presets

### 🚀 Installation

1. Download `PyPDF_Toolbox-v1.0.0-win64.zip`
2. Extract the ZIP file
3. Run `PyPDF_Toolbox.exe`
4. The launcher will appear at the top of your screen

### 🔒 Verification

Verify the integrity of the download using the SHA256 checksum:

```powershell
# Windows PowerShell
Get-FileHash -Algorithm SHA256 PyPDF_Toolbox-v1.0.0-win64.zip
# Compare with the hash in PyPDF_Toolbox-v1.0.0-win64.zip.sha256.txt
```

---

[1.0.1]: https://github.com/alorbach/pypdf-toolbox-gui/releases/tag/v1.0.1
[1.0.0]: https://github.com/alorbach/pypdf-toolbox-gui/releases/tag/v1.0.0
