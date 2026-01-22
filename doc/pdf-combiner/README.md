# PDF Combiner Documentation

The PDF Combiner allows you to combine multiple PDF files by visually selecting individual pages from thumbnails. This advanced tool provides a two-panel interface for intuitive page selection and combination.

## Features

- **Visual Page Selection**: Click thumbnails to select pages in the order you want them combined
- **Multiple PDF Support**: Load multiple PDF files and select pages from any of them
- **Drag and Drop**: Drop PDF files directly onto the application
- **Configurable Thumbnail Sizes**: Choose from 6 different thumbnail sizes (Small to Giant)
- **Auto-Selection Patterns**: 
  - Alternate Pages: Automatically select pages alternating between files
  - Alternate + Reverse: Alternate with the second file in reverse order
- **Visual Feedback**: Selected pages are highlighted with numbered badges
- **Page Preview**: See actual page content in thumbnails before selecting
- **Real-time Selection Display**: View your selected pages list as you build it

## Screenshots

![Main Window](screenshots/01-main-window.png)

*Main combiner interface with two-panel layout - controls on left, page thumbnails on right*

| Screenshot | Description |
|------------|-------------|
| `01-main-window.png` | Main combiner interface with two-panel layout |
| `02-drag-drop.png` | Drag and drop zone |
| `03-files-loaded.png` | Multiple PDFs loaded with thumbnails |
| `04-pages-selected.png` | Pages selected with visual feedback |
| `05-auto-selection.png` | Auto-selection pattern applied |
| `06-combine-complete.png` | Combine operation completed |

## Usage

### Loading PDF Files

1. **Drag and Drop**: Drag PDF files onto the drop zone in the left panel
2. **Browse**: Click the drop zone or use the file dialog to select PDF files
3. **Multiple Files**: Select multiple PDF files at once in the file dialog

### Selecting Pages

1. **Click Thumbnails**: Click on any page thumbnail in the right panel to select it
2. **Selection Order**: Pages are selected in the order you click them (numbered badges appear)
3. **Deselect**: Click a selected page again to deselect it
4. **Clear All**: Use the "Clear Selection" button to deselect all pages

### Thumbnail Sizes

Adjust the thumbnail size using the radio buttons in the "Thumbnail Size" section:

- **Small**: 120x150px - See many pages at once (8 columns)
- **Big**: 200x250px - Balanced view (5 columns) - *Default*
- **Biggest**: 640x480px - Larger preview (2 columns)
- **Huge**: 800x600px - Very large preview (1 column)
- **Massive**: 1024x768px - Maximum detail (1 column)
- **Giant**: 1280x960px - Ultra-high detail (1 column)

Changing the size regenerates all thumbnails with the new dimensions.

### Auto-Selection Patterns

When you have 2 or more PDF files loaded, you can use auto-selection:

1. **Auto: Alternate Pages**: 
   - Selects pages alternating between files
   - Pattern: File1 Page1, File2 Page1, File1 Page2, File2 Page2, etc.
   - Useful for interleaving documents

2. **Auto: Alternate + Reverse**:
   - Alternates with the second file in reverse order
   - Pattern: File1 Page1, File2 LastPage, File1 Page2, File2 2nd-LastPage, etc.
   - Useful for creating book-style layouts

### Combining PDFs

1. Select the pages you want in the desired order
2. Click "💾 Save Combined PDF" button
3. Choose the output location and filename
4. The combined PDF will be created with pages in your selected order

## Technical Details

- **Source File**: `src/pdf_combiner.py`
- **Launch Script**: `launch_pdf_visual_combiner.bat` / `launch_pdf_visual_combiner.sh`
- **Display Name**: PDF Visual Combiner (in launcher)
- **Framework**: tkinter with ttk widgets, tkinterdnd2 for drag-and-drop
- **PDF Library**: PyPDF2 / pypdf for PDF manipulation
- **Rendering Library**: PyMuPDF (fitz) for high-quality thumbnail generation
- **Image Library**: PIL/Pillow for image processing

## Dependencies

- PyPDF2 or pypdf (for PDF manipulation)
- PyMuPDF (for PDF rendering and thumbnail generation)
- Pillow (for image processing)
- tkinterdnd2 (for drag and drop support, optional)

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Mouse Wheel` | Scroll through thumbnails |
| `Click Thumbnail` | Select/deselect page |
| `Ctrl+Click` | (Future: Multi-select) |

## Tips

- **Large Files**: Very large PDFs may take time to generate thumbnails. Be patient!
- **Thumbnail Quality**: Larger thumbnail sizes provide better preview quality but take longer to generate
- **Selection Order**: The order you click pages determines their order in the combined PDF
- **Multiple Files**: You can select pages from different files in any order
- **Auto-Selection**: Use auto-selection patterns to quickly create common page arrangements
- **Clear Selection**: Use "Clear Selection" to start over without reloading files

## Workflow Example

1. Load two PDF files: "Document1.pdf" (5 pages) and "Document2.pdf" (3 pages)
2. Use "Auto: Alternate Pages" to create: Doc1-P1, Doc2-P1, Doc1-P2, Doc2-P2, Doc1-P3, Doc2-P3, Doc1-P4, Doc1-P5
3. Or manually click pages in your desired order
4. Click "Save Combined PDF" to create the final document
