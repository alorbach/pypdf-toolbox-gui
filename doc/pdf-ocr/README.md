# PDF OCR Tool Documentation

The PDF OCR Tool adds OCR (Optical Character Recognition) text to PDF files, making scanned documents searchable. It can also convert image files (JPG, PNG, TIFF, BMP) to searchable PDFs.

## Features

- **PDF OCR**: Add searchable text layer to existing PDF files
- **Image to PDF**: Convert image files to searchable PDFs with OCR
- **Multiple Languages**: Support for English, German, French, Spanish, and combinations
- **Batch Processing**: Process multiple files or entire folders at once
- **Drag and Drop**: Drop PDF and image files directly into the application
- **Recursive Folder Processing**: Process subfolders recursively
- **Automatic Image Conversion**: Automatically converts images to PDF before OCR

## Screenshots

![Main Window](screenshots/01-main-window.png)

*Main interface showing the drag & drop zone, language selection, processing log, and action buttons after a successful OCR operation*

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main interface with drag & drop zone, showing successful OCR processing |

## Usage

### Opening Files

1. **Drag and Drop**: Drag PDF or image files onto the drop zone
2. **Select Files**: Click "Select Files" to browse for PDF and image files
3. **Select Folder**: Click "Select Folder" to process all PDFs and images in a directory

### Supported File Types

- **PDF files** (`.pdf`): Existing PDFs that need OCR text layer added
- **Image files**:
  - JPEG (`.jpg`, `.jpeg`)
  - PNG (`.png`)
  - TIFF (`.tiff`)
  - BMP (`.bmp`)

### Language Selection

Choose the OCR language based on your document content:

| Language | Code | Best For |
|----------|------|----------|
| English | `eng` | English documents |
| German | `deu` | German documents |
| English + German | `eng+deu` | Mixed language documents |
| French | `fra` | French documents |
| Spanish | `spa` | Spanish documents |

### Processing Workflow

1. **PDF Files**: OCR text layer is added directly to the existing PDF file (in-place processing)
2. **Image Files**: 
   - Images are first converted to a single PDF
   - OCR is then applied to the PDF
   - The resulting PDF is saved in the same directory as the source images
   - Original image files are preserved

### Batch Processing

When processing multiple files:
- PDFs are processed individually
- Images from the same directory are grouped and converted to a single PDF
- Each file/directory is processed sequentially with progress updates

### Folder Processing

When selecting a folder:
- You can choose to process recursively (including subfolders)
- All PDF and image files found are processed
- Results are saved in the same location as the source files

## Technical Details

- **Source File**: `src/pdf_ocr.py`
- **Launch Script**: `launch_pdf_ocr.bat` / `launch_pdf_ocr.sh`
- **Dependencies**:
  - `ocrmypdf>=16.0.0` - OCR processing engine
  - `pillow>=10.0.0` - Image processing
  - `img2pdf>=0.5.0` - Image to PDF conversion
  - `tkinterdnd2>=0.3.0` - Drag and drop support

### OCR Engine

The tool uses **OCRmyPDF**, which is built on Tesseract OCR. OCRmyPDF:
- Adds a searchable text layer to PDFs without changing the visual appearance
- Skips pages that already contain text (unless forced)
- Handles various image formats and quality levels
- Supports multiple languages

### Image Processing

- Images are converted to RGB format if necessary
- Multiple images are combined into a single PDF
- The PDF is then processed with OCR
- Original image files are preserved (not deleted)

## Error Handling

The tool handles common scenarios:

- **PDFs with existing text**: Skips OCR if text layer already exists
- **Unsupported image formats**: Converts to JPEG if needed
- **Processing errors**: Logs errors and continues with remaining files
- **Missing dependencies**: Shows clear error messages with installation instructions

## Limitations

- **Large files**: Very large PDFs or many images may take significant time to process
- **Image quality**: OCR accuracy depends on image quality and resolution
- **Language support**: Limited to languages supported by Tesseract OCR
- **File size**: Very large image files may require significant memory

## Tips for Best Results

1. **Image Quality**: Use high-resolution images (300 DPI or higher) for best OCR accuracy
2. **Language Selection**: Choose the correct language for your documents
3. **Batch Processing**: Process similar documents together for consistency
4. **File Organization**: Keep related images in the same folder for automatic grouping

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Open file dialog (via Select Files button) |
| `Ctrl+F` | Select folder (via Select Folder button) |

## Troubleshooting

### "OCRmyPDF not available" Error

Install the required dependency:
```bash
pip install ocrmypdf
```

Note: OCRmyPDF requires Tesseract OCR to be installed on your system.

### "Pillow not available" Error

Install Pillow for image processing:
```bash
pip install pillow
```

### "img2pdf not available" Error

Install img2pdf for image to PDF conversion:
```bash
pip install img2pdf
```

### OCR Quality Issues

- Ensure images are high resolution (300+ DPI)
- Check that the correct language is selected
- Clean scanned images before processing (remove noise, adjust contrast)
- For mixed language documents, use language combinations (e.g., `eng+deu`)

### Processing Takes Too Long

- Large files or many images will take time
- OCR processing is CPU-intensive
- Consider processing files in smaller batches
- Close other applications to free up system resources

## Related Tools

- **PDF Text Extractor**: Extract text from PDFs (including OCR'd PDFs)
- **PDF Splitter**: Split large PDFs into smaller files
- **PDF Combiner**: Combine multiple PDFs into one
