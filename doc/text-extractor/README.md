# PDF Text Extractor Documentation

The PDF Text Extractor allows you to extract text from PDF documents using multiple methods: Python-based extraction, OCR for scanned documents, and Azure AI Document Intelligence for high-quality extraction.

## Features

- **Multiple Extraction Methods**:
  - **Python (PyMuPDF)**: Fast, offline extraction for digital PDFs
  - **OCR (OCRmyPDF)**: Extract text from scanned documents
  - **Azure AI**: High-quality extraction with layout preservation
- **Multiple Output Formats**: Text (.txt), Markdown (.md), JSON
- **Drag and Drop**: Drop PDF files directly into the application
- **Batch Processing**: Process multiple PDFs at once
- **Folder Processing**: Process entire folders with optional recursion
- **Layout Preservation**: Maintains document structure where possible

## Screenshots

![Main Window](screenshots/01-main-window.png)

*Main interface showing extraction methods, output formats, and drag & drop zone*

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main interface |
| `02-drag-drop.png` | Drag and drop zone |
| `03-file-processing.png` | File being processed |
| `04-results.png` | Extraction results |
| `05-azure-config.png` | Azure configuration dialog |

## Usage

### Opening Files

1. **Drag and Drop**: Drag PDF files onto the drop zone
2. **Select Files**: Click "Select PDF Files" to browse
3. **Select Folder**: Click "Select Folder" to process all PDFs in a directory

### Choosing Extraction Method

Select the appropriate method based on your PDF type:

| Method | Best For | Requirements |
|--------|----------|--------------|
| Python (PyMuPDF) | Digital PDFs with embedded text | None (included) |
| OCR | Scanned documents, image-based PDFs | Tesseract OCR |
| Azure AI | Complex layouts, high accuracy needed | Azure subscription |

### Output Formats

- **Text (.txt)**: Plain text with page markers
- **Markdown (.md)**: Structured text with headings and formatting
- **JSON**: Full structured data with metadata

### Azure AI Configuration

To use Azure AI Document Intelligence:

1. Click "⚙️ Azure Config" in the options panel
2. Enter your Azure Document Intelligence endpoint
3. Enter your API key
4. Click "Test Connection" to verify
5. Click "Save" to store settings

Alternatively, set environment variables:
- `AZURE_DOC_INTEL_ENDPOINT`: Your endpoint URL
- `AZURE_DOC_INTEL_API_KEY`: Your API key

## Command Line Usage

```bash
# Extract text from a single PDF
python src/pdf_text_extractor.py document.pdf

# Extract using OCR
python src/pdf_text_extractor.py --method ocr scanned.pdf

# Extract using Azure AI
python src/pdf_text_extractor.py --method azure document.pdf

# Output as Markdown
python src/pdf_text_extractor.py --format markdown document.pdf

# Process folder recursively
python src/pdf_text_extractor.py --recursive folder/

# Force overwrite existing files
python src/pdf_text_extractor.py --force document.pdf

# Start GUI mode
python src/pdf_text_extractor.py --gui
```

### Command Line Options

| Option | Description |
|--------|-------------|
| `--method python` | Use PyMuPDF extraction (default) |
| `--method ocr` | Use OCR-based extraction |
| `--method azure` | Use Azure AI extraction |
| `--format text` | Output as plain text (default) |
| `--format markdown` | Output as Markdown |
| `--format json` | Output as JSON |
| `--recursive, -r` | Process folders recursively |
| `--force` | Overwrite existing output files |
| `--output-dir, -o` | Custom output directory |
| `--gui` | Start GUI mode |
| `--debug` | Enable debug output |

## Technical Details

- **Source File**: `src/pdf_text_extractor.py`
- **Launch Script**: `launch_pdf_text_extractor.bat` / `launch_pdf_text_extractor.sh`
- **Framework**: tkinter with ttk widgets, tkinterdnd2 for drag-and-drop
- **Display Name**: PDF Text Extractor (in launcher)
- **PDF Libraries**: 
  - PyMuPDF (fitz) for Python extraction
  - OCRmyPDF + Tesseract for OCR
  - Azure Document Intelligence for AI extraction

## Dependencies

### Required
- `pymupdf` - PDF reading and text extraction
- `pillow` - Image processing
- `pyyaml` - Configuration files

### Optional
- `ocrmypdf` - OCR support (requires Tesseract OCR)
- `requests` - Azure API calls
- `azure-identity` - Azure authentication
- `tkinterdnd2` - Drag and drop support

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Open file |
| `Escape` | Cancel operation |

## Tips

- **For digital PDFs**: Use Python method for fastest results
- **For scanned PDFs**: Use OCR method to extract text from images
- **For complex layouts**: Use Azure AI for best layout preservation
- **For batch processing**: Select a folder and enable recursive search
- **To skip existing**: Leave "Overwrite existing files" unchecked

## Troubleshooting

### "PyMuPDF not available"
Install with: `pip install pymupdf`

### "OCRmyPDF not available"
1. Install Tesseract OCR on your system
2. Install package: `pip install ocrmypdf`

### "Azure AI not configured"
1. Create an Azure Document Intelligence resource
2. Configure endpoint and API key in the Azure Config dialog
3. Or set environment variables

### OCR quality is poor
- Ensure the scanned document is high resolution
- Try preprocessing the PDF (deskew, enhance contrast)
- Use Azure AI for better results
