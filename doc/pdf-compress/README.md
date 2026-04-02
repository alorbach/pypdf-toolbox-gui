# PDF / Image Recompress

Shrink embedded raster images inside PDFs (linear scale 50–100%) and recompress as JPEG, or batch-process standalone image files to smaller JPEGs.

## Features

- PDF: each embedded image is decoded, optionally scaled down, and written back as JPEG (vectors and text are not rasterized).
- Images: JPG, PNG, WebP, BMP, TIFF supported as input; output is JPEG.
- Adjustable JPEG quality (40–95).
- Drag and drop or file picker; batch output to a chosen folder (`*_compressed.pdf` / `*_compressed.jpg`).

## Technical

- **Source**: `src/pdf_compress.py`
- **Launch**: `launch_pdf_compress.bat` / `launch_pdf_compress.sh`
- **Dependencies**: PyMuPDF, Pillow, tkinterdnd2 (optional)

## Screenshots

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main interface (placeholder) |
