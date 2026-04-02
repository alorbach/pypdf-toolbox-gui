"""
PDF / Image Recompress Tool

Resize embedded images in PDFs (linear scale down to 50%) and recompress as JPEG,
or batch-resize and recompress standalone raster image files.

Copyright 2025-2026 Andre Lorbach

Licensed under the Apache License, Version 2.0
"""

import io
import logging
import os
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

import tkinter.ttk as ttk

try:
    import fitz
except ImportError:
    fitz = None

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger(__name__)


class UIColors:
    PRIMARY = "#2563eb"
    PRIMARY_HOVER = "#1d4ed8"
    SUCCESS = "#16a34a"
    SUCCESS_HOVER = "#15803d"
    ERROR = "#dc2626"
    ERROR_HOVER = "#b91c1c"
    BG_PRIMARY = "#ffffff"
    BG_SECONDARY = "#f8fafc"
    BG_TERTIARY = "#f1f5f9"
    BORDER = "#e2e8f0"
    TEXT_PRIMARY = "#1e293b"
    TEXT_SECONDARY = "#64748b"
    TEXT_MUTED = "#94a3b8"
    DROP_ZONE_BG = "#f8fafc"
    DROP_ZONE_BORDER = "#94a3b8"
    DROP_ZONE_ACTIVE = "#dbeafe"
    DROP_ZONE_BORDER_ACTIVE = "#2563eb"


class UIFonts:
    TITLE = ("Segoe UI", 18, "bold")
    HEADING = ("Segoe UI", 12, "bold")
    BODY = ("Segoe UI", 10)
    SMALL = ("Segoe UI", 9)
    BUTTON = ("Segoe UI", 10, "bold")


class UISpacing:
    SM = 5
    MD = 10
    LG = 15


PDF_SUFFIX = ".pdf"
RASTER_SUFFIXES = frozenset({
    ".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff",
})
ALL_SUFFIXES = {PDF_SUFFIX} | RASTER_SUFFIXES


def create_rounded_button(parent, text, command, style="primary", width=None):
    colors = {
        "primary": (UIColors.PRIMARY, UIColors.PRIMARY_HOVER, "#ffffff"),
        "secondary": (UIColors.BG_TERTIARY, UIColors.BORDER, UIColors.TEXT_PRIMARY),
        "success": (UIColors.SUCCESS, UIColors.SUCCESS_HOVER, "#ffffff"),
        "danger": (UIColors.ERROR, UIColors.ERROR_HOVER, "#ffffff"),
    }
    bg, hover_bg, fg = colors.get(style, colors["primary"])
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        font=UIFonts.BUTTON,
        bg=bg,
        fg=fg,
        activebackground=hover_bg,
        activeforeground=fg,
        relief="flat",
        cursor="hand2",
        padx=UISpacing.MD,
        pady=UISpacing.SM,
        bd=0,
        highlightthickness=0,
    )
    if width:
        btn.config(width=width)

    def on_enter(_):
        btn.config(bg=hover_bg)

    def on_leave(_):
        btn.config(bg=bg)

    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    return btn


def _pil_to_jpeg_bytes(pil_img: Image.Image, quality: int) -> bytes:
    if pil_img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", pil_img.size, (255, 255, 255))
        if pil_img.mode == "RGBA":
            bg.paste(pil_img, mask=pil_img.split()[3])
        else:
            bg.paste(pil_img, mask=pil_img.split()[1])
        pil_img = bg
    elif pil_img.mode == "P":
        if pil_img.info.get("transparency") is not None:
            pil_img = pil_img.convert("RGBA")
            bg = Image.new("RGB", pil_img.size, (255, 255, 255))
            bg.paste(pil_img, mask=pil_img.split()[3])
            pil_img = bg
        else:
            pil_img = pil_img.convert("RGB")
    elif pil_img.mode != "RGB":
        pil_img = pil_img.convert("RGB")
    buf = io.BytesIO()
    pil_img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def recompress_image_file(src: str, dest: str, scale: float, jpeg_quality: int) -> None:
    with Image.open(src) as im:
        w, h = im.size
        nw = max(1, int(w * scale))
        nh = max(1, int(h * scale))
        if nw != w or nh != h:
            im = im.resize((nw, nh), Image.Resampling.LANCZOS)
        data = _pil_to_jpeg_bytes(im, jpeg_quality)
    Path(dest).parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        f.write(data)


def recompress_pdf_images(src: str, dest: str, scale: float, jpeg_quality: int) -> tuple[int, int]:
    """Replace embedded images; returns (images_replaced, images_skipped)."""
    doc = fitz.open(src)
    replaced = 0
    skipped = 0
    seen_xrefs = set()
    try:
        for pno in range(len(doc)):
            page = doc[pno]
            for info in page.get_images(full=True):
                xref = info[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                try:
                    raw = doc.extract_image(xref)
                except Exception as e:
                    logger.warning("extract_image xref=%s: %s", xref, e)
                    skipped += 1
                    continue
                data = raw.get("image")
                if not data:
                    skipped += 1
                    continue
                try:
                    pil = Image.open(io.BytesIO(data))
                    pil.load()
                except Exception as e:
                    logger.warning("PIL decode xref=%s: %s", xref, e)
                    skipped += 1
                    continue
                w, h = pil.size
                nw = max(1, int(w * scale))
                nh = max(1, int(h * scale))
                if nw != w or nh != h:
                    pil = pil.resize((nw, nh), Image.Resampling.LANCZOS)
                try:
                    jpeg_bytes = _pil_to_jpeg_bytes(pil, jpeg_quality)
                    page.replace_image(xref, stream=jpeg_bytes)
                    replaced += 1
                except Exception as e:
                    logger.warning("replace_image xref=%s: %s", xref, e)
                    skipped += 1
        doc.save(dest, garbage=4, deflate=True, clean=True)
    finally:
        doc.close()
    return replaced, skipped


class PDFCompressApp:
    def __init__(self):
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("PDF / Image Recompress")
        self.root.minsize(720, 520)
        self.root.configure(bg=UIColors.BG_SECONDARY)

        x = int(os.environ.get("TOOL_WINDOW_X", 100))
        y = int(os.environ.get("TOOL_WINDOW_Y", 100))
        w = int(os.environ.get("TOOL_WINDOW_WIDTH", 900))
        h = int(os.environ.get("TOOL_WINDOW_HEIGHT", 700))
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self.file_paths: list[str] = []
        self.output_dir = tk.StringVar(value="")
        self.scale_percent = tk.IntVar(value=50)
        self.jpeg_quality = tk.IntVar(value=80)
        self._busy = False

        self._build_ui()
        if DND_AVAILABLE:
            self._setup_dnd()

    def _build_ui(self):
        main = tk.Frame(self.root, bg=UIColors.BG_SECONDARY, padx=UISpacing.LG, pady=UISpacing.LG)
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main,
            text="Shrink & recompress images",
            font=UIFonts.TITLE,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.PRIMARY,
        ).pack(anchor=tk.W)

        tk.Label(
            main,
            text="PDFs: embedded images are resized (linear scale) and saved as JPEG. "
            "Images: saved as compressed JPEG. Vectors and text in PDFs are unchanged.",
            font=UIFonts.SMALL,
            bg=UIColors.BG_SECONDARY,
            fg=UIColors.TEXT_SECONDARY,
            wraplength=820,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(UISpacing.SM, UISpacing.MD))

        drop = tk.Frame(
            main,
            bg=UIColors.DROP_ZONE_BG,
            highlightbackground=UIColors.DROP_ZONE_BORDER,
            highlightthickness=2,
            padx=UISpacing.MD,
            pady=UISpacing.MD,
        )
        drop.pack(fill=tk.X, pady=(0, UISpacing.MD))
        self.drop_frame = drop
        self._drop_widgets = []
        msg = (
            "Drag and drop PDF or image files here"
            if DND_AVAILABLE
            else "Use Add files to select PDF or images"
        )
        la = tk.Label(
            drop,
            text=msg,
            font=UIFonts.HEADING,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_PRIMARY,
            cursor="hand2",
        )
        la.pack()
        lb = tk.Label(
            drop,
            text="or click to add files",
            font=UIFonts.SMALL,
            bg=UIColors.DROP_ZONE_BG,
            fg=UIColors.TEXT_MUTED,
            cursor="hand2",
        )
        lb.pack()
        for w in (drop, la, lb):
            w.bind("<Button-1>", lambda e: self.add_files())
        self._drop_widgets = [drop, la, lb]

        row = tk.Frame(main, bg=UIColors.BG_SECONDARY)
        row.pack(fill=tk.X, pady=(0, UISpacing.SM))
        create_rounded_button(row, "Add files…", self.add_files, style="primary").pack(
            side=tk.LEFT, padx=(0, UISpacing.SM)
        )
        create_rounded_button(row, "Select all", self.select_all_files, style="secondary").pack(
            side=tk.LEFT, padx=(0, UISpacing.SM)
        )
        create_rounded_button(row, "Clear list", self.clear_files, style="danger").pack(side=tk.LEFT)

        list_frame = tk.Frame(main, bg=UIColors.BG_PRIMARY, bd=1, relief=tk.SOLID)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, UISpacing.MD))
        self.file_list = tk.Listbox(
            list_frame,
            height=8,
            font=("Consolas", 9),
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            selectbackground=UIColors.DROP_ZONE_ACTIVE,
            selectmode=tk.EXTENDED,
            exportselection=False,
        )
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=sb.set)
        self.file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_list.bind("<Control-a>", self._on_select_all_key)
        self.file_list.bind("<Control-A>", self._on_select_all_key)

        opts = tk.LabelFrame(main, text=" Options ", font=UIFonts.HEADING, bg=UIColors.BG_PRIMARY, fg=UIColors.TEXT_PRIMARY)
        opts.pack(fill=tk.X, pady=(0, UISpacing.MD))

        sf = tk.Frame(opts, bg=UIColors.BG_PRIMARY)
        sf.pack(fill=tk.X, padx=UISpacing.MD, pady=UISpacing.SM)
        tk.Label(
            sf,
            text="Image linear scale (% of original width/height, min 50%):",
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            font=UIFonts.BODY,
        ).pack(anchor=tk.W)
        sc = tk.Scale(
            sf,
            from_=50,
            to=100,
            orient=tk.HORIZONTAL,
            variable=self.scale_percent,
            tickinterval=10,
            resolution=1,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            highlightthickness=0,
        )
        sc.pack(fill=tk.X)

        qf = tk.Frame(opts, bg=UIColors.BG_PRIMARY)
        qf.pack(fill=tk.X, padx=UISpacing.MD, pady=UISpacing.SM)
        tk.Label(
            qf,
            text="JPEG quality (lower = smaller files):",
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            font=UIFonts.BODY,
        ).pack(anchor=tk.W)
        jq = tk.Scale(
            qf,
            from_=40,
            to=95,
            orient=tk.HORIZONTAL,
            variable=self.jpeg_quality,
            tickinterval=10,
            resolution=1,
            bg=UIColors.BG_PRIMARY,
            fg=UIColors.TEXT_PRIMARY,
            highlightthickness=0,
        )
        jq.pack(fill=tk.X)

        out_row = tk.Frame(main, bg=UIColors.BG_SECONDARY)
        out_row.pack(fill=tk.X, pady=(0, UISpacing.SM))
        tk.Label(out_row, text="Output folder:", bg=UIColors.BG_SECONDARY, fg=UIColors.TEXT_PRIMARY).pack(
            side=tk.LEFT
        )
        self.out_entry = tk.Entry(out_row, textvariable=self.output_dir, width=60, font=UIFonts.SMALL)
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=UISpacing.SM)
        create_rounded_button(out_row, "Browse…", self.browse_output, style="primary").pack(side=tk.LEFT)

        self.process_btn = create_rounded_button(
            main, "Process all files", self.process_all, style="success"
        )
        self.process_btn.pack(fill=tk.X, pady=(0, UISpacing.SM))

        tk.Label(main, text="Log", font=UIFonts.HEADING, bg=UIColors.BG_SECONDARY, fg=UIColors.TEXT_PRIMARY).pack(
            anchor=tk.W
        )
        self.log = scrolledtext.ScrolledText(
            main,
            height=12,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
        )
        self.log.pack(fill=tk.BOTH, expand=True)

    def _log(self, line: str):
        self.log.insert(tk.END, line + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    def _setup_dnd(self):
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_frame.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self.drop_frame.dnd_bind("<<DragLeave>>", self._on_drag_leave)

    def _on_drag_enter(self, _event):
        self.drop_frame.config(
            bg=UIColors.DROP_ZONE_ACTIVE, highlightbackground=UIColors.DROP_ZONE_BORDER_ACTIVE
        )
        for w in self._drop_widgets:
            if isinstance(w, tk.Label):
                w.config(bg=UIColors.DROP_ZONE_ACTIVE)

    def _on_drag_leave(self, _event):
        self.drop_frame.config(bg=UIColors.DROP_ZONE_BG, highlightbackground=UIColors.DROP_ZONE_BORDER)
        for w in self._drop_widgets:
            if isinstance(w, tk.Label):
                w.config(bg=UIColors.DROP_ZONE_BG)

    def _parse_dnd_paths(self, data: str) -> list[str]:
        files = []
        if "{" in data:
            files = re.findall(r"\{([^}]+)\}", data)
            remaining = re.sub(r"\{[^}]+\}", "", data).strip()
            if remaining:
                files.extend(remaining.split())
        else:
            files = data.split()
        return [f for f in files if Path(f).suffix.lower() in ALL_SUFFIXES]

    def _on_drop(self, event):
        self._on_drag_leave(event)
        paths = self._parse_dnd_paths(event.data)
        self._add_paths(paths)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF or image files",
            filetypes=[
                ("All supported", "*.pdf *.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"),
                ("PDF", "*.pdf"),
                ("Images", "*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"),
                ("All files", "*.*"),
            ],
        )
        self._add_paths(list(paths))

    def _add_paths(self, paths: list[str]):
        for p in paths:
            if not p or Path(p).suffix.lower() not in ALL_SUFFIXES:
                continue
            if p not in self.file_paths:
                self.file_paths.append(p)
                self.file_list.insert(tk.END, p)
        if self.file_paths and not self.output_dir.get().strip():
            self.output_dir.set(str(Path(self.file_paths[0]).parent))

    def clear_files(self):
        self.file_paths.clear()
        self.file_list.delete(0, tk.END)

    def select_all_files(self):
        """Select every row in current list order (same as Ctrl+A)."""
        n = self.file_list.size()
        if n == 0:
            return
        self.file_list.selection_clear(0, tk.END)
        self.file_list.select_set(0, tk.END)
        self.file_list.activate(0)
        self.file_list.focus_set()
        self.file_list.see(0)

    def _on_select_all_key(self, _event=None):
        self.select_all_files()
        return "break"

    def browse_output(self):
        d = filedialog.askdirectory(title="Output folder for compressed files")
        if d:
            self.output_dir.set(d)

    def process_all(self):
        if self._busy:
            return
        if not self.file_paths:
            messagebox.showwarning("No files", "Add at least one PDF or image file.")
            return
        out = self.output_dir.get().strip()
        if not out:
            messagebox.showwarning("Output folder", "Choose an output folder.")
            return
        out_path = Path(out)
        if not out_path.is_dir():
            messagebox.showerror("Output folder", "Output path is not a folder or does not exist.")
            return

        scale = self.scale_percent.get() / 100.0
        quality = self.jpeg_quality.get()
        self._busy = True
        self.process_btn.config(state=tk.DISABLED)
        self.root.config(cursor="wait")

        def work():
            try:
                for src in self.file_paths:
                    sp = Path(src)
                    suffix = sp.suffix.lower()
                    stem = sp.stem
                    if suffix == PDF_SUFFIX:
                        dest = str(out_path / f"{stem}_compressed.pdf")
                        line = f"PDF: {sp.name} -> {Path(dest).name}"
                        self.root.after(0, lambda m=line: self._log(m))
                        try:
                            rep, sk = recompress_pdf_images(src, dest, scale, quality)
                        except Exception as ex:
                            err = str(ex)
                            self.root.after(
                                0,
                                lambda n=sp.name, e=err: self._log(f"  ERROR {n}: {e}"),
                            )
                            continue
                        line2 = f"  replaced {rep} image(s), skipped {sk}"
                        self.root.after(0, lambda m=line2: self._log(m))
                    elif suffix in RASTER_SUFFIXES:
                        dest = str(out_path / f"{stem}_compressed.jpg")
                        line = f"Image: {sp.name} -> {Path(dest).name}"
                        self.root.after(0, lambda m=line: self._log(m))
                        try:
                            recompress_image_file(src, dest, scale, quality)
                        except Exception as ex:
                            err = str(ex)
                            self.root.after(
                                0,
                                lambda n=sp.name, e=err: self._log(f"  ERROR {n}: {e}"),
                            )
                            continue
                        self.root.after(0, lambda: self._log("  done"))
                    else:
                        line = f"Skip unsupported: {sp.name}"
                        self.root.after(0, lambda m=line: self._log(m))
                self.root.after(0, self._done_ok)
            except Exception as e:
                logger.exception("process")
                self.root.after(0, lambda err=str(e): self._done_err(err))

        threading.Thread(target=work, daemon=True).start()

    def _done_ok(self):
        self._busy = False
        self.process_btn.config(state=tk.NORMAL)
        self.root.config(cursor="")
        self._log("Finished.")
        messagebox.showinfo("Done", "Processing finished. See log for details.")

    def _done_err(self, err: str):
        self._busy = False
        self.process_btn.config(state=tk.NORMAL)
        self.root.config(cursor="")
        messagebox.showerror("Error", err)

    def run(self):
        self.root.mainloop()


def main():
    if fitz is None:
        print("[ERROR] PyMuPDF (pymupdf) required: pip install pymupdf")
        return
    if not PIL_AVAILABLE:
        print("[ERROR] Pillow required: pip install Pillow")
        return
    PDFCompressApp().run()


if __name__ == "__main__":
    main()
