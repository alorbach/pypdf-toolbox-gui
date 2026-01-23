# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PyPDF Toolbox
Creates a single executable containing launcher and all tools
"""

# Block cipher for encryption (can be None for no encryption)
block_cipher = None

a = Analysis(
    ['PyPDF_Toolbox.py'],  # Use relative path - spec file is in project root
    pathex=['src'],  # Add src directory to path so PyInstaller can find modules
    binaries=[],
    datas=[
        # Include config template if it exists (relative paths)
        ('config/azure_ai.yaml.template', 'config'),
    ],
    hiddenimports=[
        # Core libraries
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'tkinterdnd2',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        
        # PDF libraries
        'PyPDF2',
        'pypdf',
        'pymupdf',
        'pdf2image',
        'reportlab',
        
        # OCR
        'pytesseract',
        'ocrmypdf',
        'img2pdf',
        
        # Utilities
        'yaml',
        'requests',
        'tqdm',
        'markdown',
        'docx',
        'html2text',
        'tkinterweb',
        
        # Azure AI (optional)
        'openai',
        'azure.identity',
        
        # Tool modules
        'launcher_gui',
        'pdf_ocr',
        'pdf_text_extractor',
        'pdf_combiner',
        'pdf_manual_splitter',
        'pdf_md_converter',
        'utils.azure_config',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PyPDF_Toolbox',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console by default (use --debug flag in build script to enable)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Can add icon file path here if you have one
)
