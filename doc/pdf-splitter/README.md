# PDF Splitter Documentation

The PDF Splitter allows you to manually split PDF documents by selecting split points with visual thumbnail preview.

## Features

- **Drag and Drop**: Drop PDF files directly into the application
- **Page Preview**: Visual preview of PDF pages
- **Custom Page Ranges**: Select specific pages or ranges to extract
- **Multiple Output**: Split into multiple documents at once
- **Batch Processing**: Process multiple PDFs

## Screenshots

![Main Window](screenshots/01-main-window.png)

*Main splitter interface with drag & drop zone and results area*

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main splitter interface |
| `02-drag-drop.png` | Drag and drop zone |
| `03-file-loaded.png` | PDF loaded with page list |
| `04-page-selection.png` | Selecting pages to split |
| `05-split-complete.png` | Split operation completed |

## Usage

### Opening a PDF

1. **Drag and Drop**: Drag a PDF file onto the application window
2. **Browse**: Click the browse button to select a file

### Selecting Pages

- Click individual pages to select them
- Use Ctrl+Click to select multiple pages
- Use Shift+Click to select a range
- Enter page ranges manually (e.g., "1-5, 8, 10-12")

### Splitting the PDF

1. Select the pages you want to extract
2. Choose the output location
3. Click "Split" to create the new PDF(s)

## Technical Details

- **Source File**: `src/pdf_manual_splitter.py`
- **Launch Script**: `launch_pdf_splitter.bat` / `launch_pdf_splitter.sh`
- **Framework**: tkinter with ttk widgets, tkinterdnd2 for drag-and-drop
- **PDF Library**: PyPDF2 / pypdf
- **Display Name**: PDF Splitter (in launcher)

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Open file |
| `Ctrl+A` | Select all pages |
| `Ctrl+S` | Split selected pages |
| `Escape` | Clear selection |

## Tips

- Use page thumbnails to quickly identify content
- Split large documents into logical sections
- Preview pages before splitting to ensure correct selection
