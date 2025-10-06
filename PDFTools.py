import os
import platform
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfMerger


def is_windows():
    return platform.system().lower().startswith("win")


def is_macos():
    return platform.system().lower() == "darwin"


class Launcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Tools")
        self.resizable(False, False)

        # Store user-provided Poppler bin directory (session-level)
        self.poppler_bin_dir = tk.StringVar(value="")

        c = ttk.Frame(self, padding=16)
        c.grid(sticky="nsew")

        ttk.Label(c, text="Choose a tool:").grid(row=0, column=0, columnspan=3, pady=(0, 10))

        # Tool buttons
        ttk.Button(c, text="PDF → JPEG", command=self.open_pdf2jpg).grid(row=1, column=0, padx=4, pady=4, sticky="ew")
        ttk.Button(c, text="JPEG → PDF", command=self.open_jpg2pdf).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        ttk.Button(c, text="Combine PDF", command=self.open_combine).grid(row=1, column=2, padx=4, pady=4, sticky="ew")

        # Poppler bin directory controls
        row2 = ttk.Frame(c)
        row2.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 4))
        ttk.Label(row2, text="Poppler bin directory (optional):").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(row2, textvariable=self.poppler_bin_dir, width=60)
        entry.grid(row=1, column=0, padx=(0, 8), pady=(2, 0), sticky="ew")
        ttk.Button(row2, text="Browse", command=self.browse_poppler_bin).grid(row=1, column=1, sticky="w")

        # Help text depending on OS
        help_msg = [
            "If Poppler is not on PATH, set the folder containing pdftoppm.",
        ]
        if is_windows():
            help_msg.append("Windows tip: this is usually the ...\\poppler-xx\\Library\\bin folder.")
        elif is_macos():
            help_msg.append("macOS tip: if installed via Homebrew, Poppler is typically on PATH already.")
            help_msg.append("Otherwise, point this to the directory containing pdftoppm.")
        else:
            help_msg.append("Linux tip: install poppler-utils via your package manager, or set the bin folder here.")

        ttk.Label(c, text="\n".join(help_msg), foreground="#555").grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))

        for i in range(3):
            c.columnconfigure(i, weight=1)

    def browse_poppler_bin(self):
        d = filedialog.askdirectory(title="Select Poppler bin directory (folder containing pdftoppm)")
        if d:
            self.poppler_bin_dir.set(d)

    def open_pdf2jpg(self):
        ToolPDF2JPG(self)

    def open_jpg2pdf(self):
        ToolJPG2PDF(self)

    def open_combine(self):
        ToolCombinePDF(self)


def resolve_pdftoppm_path(poppler_bin_dir: str) -> tuple[bool, str | None]:
    """
    Determine if pdftoppm is available via PATH or via user-provided poppler_bin_dir.
    Returns (available, poppler_path_for_pdf2image)
      - available: True if pdftoppm can be executed one way or the other
      - poppler_path_for_pdf2image: the directory to pass as poppler_path to pdf2image (or None to use PATH)
    """
    # If user supplied a directory and it contains pdftoppm, prefer that.
    if poppler_bin_dir:
        exe_name = "pdftoppm.exe" if is_windows() else "pdftoppm"
        candidate = os.path.join(poppler_bin_dir, exe_name)
        if os.path.isfile(candidate):
            return True, poppler_bin_dir  # use explicit path
        # Allow user to pick the root and still succeed if a 'bin' subdir exists
        bin_candidate = os.path.join(poppler_bin_dir, "bin", exe_name)
        if os.path.isfile(bin_candidate):
            return True, os.path.join(poppler_bin_dir, "bin")

    # Fallback to PATH
    if shutil.which("pdftoppm"):
        return True, None  # None means "use PATH"

    return False, None


def poppler_help_text():
    base = [
        "pdftoppm not found.",
        "Fix: either set the Poppler bin directory (folder with pdftoppm) or add Poppler to your PATH.",
        "Download Poppler: https://poppler.freedesktop.org/.",
    ]
    if is_windows():
        base.append("On Windows, the correct folder is usually ...\\poppler-XX\\Library\\bin.")
    elif is_macos():
        base.append("On macOS, try: brew install poppler (then restart the app) or set the folder containing pdftoppm.")
    else:
        base.append("On Linux, install poppler-utils via your package manager.")
    return "\n".join(base)


class ToolPDF2JPG(tk.Toplevel):
    def __init__(self, master: Launcher):
        super().__init__(master)
        self.master = master
        self.title("PDF → JPEG")
        self.resizable(False, False)
        self.pdf_path = tk.StringVar()
        c = ttk.Frame(self, padding=12)
        c.grid(sticky="nsew")
        ttk.Entry(c, textvariable=self.pdf_path, width=60).grid(row=0, column=0, padx=(0, 8), pady=(0, 8), sticky="ew")
        ttk.Button(c, text="Browse", command=self.browse).grid(row=0, column=1, pady=(0, 8))
        ttk.Button(c, text="Convert", command=self.convert).grid(row=1, column=0, padx=(0, 8), sticky="w")
        ttk.Button(c, text="Close", command=self.destroy).grid(row=1, column=1, sticky="w")
        self.status = tk.StringVar(value="Ready")
        ttk.Label(c, textvariable=self.status, relief="sunken", anchor="w").grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        c.columnconfigure(0, weight=1)

    def browse(self):
        f = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if f:
            self.pdf_path.set(f)
            self.status.set("Selected: " + os.path.basename(f))

    def convert(self):
        pdf = self.pdf_path.get().strip()
        if not pdf or not os.path.isfile(pdf) or not pdf.lower().endswith(".pdf"):
            messagebox.showerror("Invalid file", "Please select a valid PDF file.")
            return

        available, poppler_path = resolve_pdftoppm_path(self.master.poppler_bin_dir.get().strip())
        if not available:
            messagebox.showerror("Poppler required", poppler_help_text())
            return

        base_dir = os.path.dirname(pdf)
        base_name = os.path.splitext(os.path.basename(pdf))[0]
        out_dir = os.path.join(base_dir, f"{base_name}_images")
        os.makedirs(out_dir, exist_ok=True)
        try:
            self.status.set("Converting at 300 DPI...")
            self.update_idletasks()

            # Pass poppler_path when we have an explicit directory; otherwise rely on PATH
            kwargs = dict(dpi=300, output_folder=out_dir, fmt="ppm", paths_only=True)
            if poppler_path:
                kwargs["poppler_path"] = poppler_path

            ppm_paths = convert_from_path(pdf, **kwargs)
            count = 0
            for i, ppm in enumerate(sorted(ppm_paths)):
                out_path = os.path.join(out_dir, f"{base_name}[{i}].jpeg")
                with Image.open(ppm) as im:
                    im.convert("RGB").save(out_path, "JPEG", quality=92, optimize=True, progressive=True)
                count += 1
            self.status.set(f"Done. Saved {count} JPEG(s) to: {out_dir}")
            messagebox.showinfo("Completed", f"Converted {count} page(s).\nFolder: {out_dir}")
        except Exception as e:
            self.status.set("Error during conversion.")
            messagebox.showerror("Conversion failed", str(e))


class ToolJPG2PDF(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("JPEG → PDF")
        self.resizable(False, False)
        self.img_paths = []
        self.folder = tk.StringVar()
        c = ttk.Frame(self, padding=12)
        c.grid(sticky="nsew")
        ttk.Entry(c, textvariable=self.folder, width=60).grid(row=0, column=0, padx=(0, 8), pady=(0, 8), sticky="ew")
        ttk.Button(c, text="Browse", command=self.browse_folder).grid(row=0, column=1, pady=(0, 8))
        ttk.Button(c, text="Convert", command=self.convert).grid(row=1, column=0, padx=(0, 8), sticky="w")
        ttk.Button(c, text="Close", command=self.destroy).grid(row=1, column=1, sticky="w")
        self.status = tk.StringVar(value="Select a folder containing image(s). Files will be ordered by name.")
        ttk.Label(c, textvariable=self.status, relief="sunken", anchor="w").grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        c.columnconfigure(0, weight=1)

    def browse_folder(self):
        d = filedialog.askdirectory(title="Select folder with images")
        if d:
            self.folder.set(d)
            self.status.set("Selected: " + d)

    def convert(self):
        d = self.folder.get().strip()
        if not d or not os.path.isdir(d):
            messagebox.showerror("Invalid folder", "Please select a valid folder.")
            return
        imgs = [
            os.path.join(d, f)
            for f in os.listdir(d)
            if f.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"))
        ]
        if not imgs:
            messagebox.showwarning("No images", "No compatible images found in the selected folder.")
            return
        imgs.sort(key=lambda x: os.path.basename(x).lower())
        save_path = filedialog.asksaveasfilename(
            title="Save PDF as", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile="output.pdf"
        )
        if not save_path:
            return
        try:
            self.status.set("Building PDF...")
            self.update_idletasks()
            pil_images = []
            for p in imgs:
                im = Image.open(p)
                if im.mode in ("RGBA", "P"):
                    im = im.convert("RGB")
                else:
                    im = im.copy()
                pil_images.append(im)
            first, rest = pil_images[0], pil_images[1:]
            first.save(save_path, "PDF", resolution=300.0, save_all=True, append_images=rest)
            for im in pil_images:
                im.close()
            self.status.set(f"Saved PDF: {save_path}")
            messagebox.showinfo("Completed", f"Created PDF from {len(imgs)} image(s).\n{save_path}")
        except Exception as e:
            self.status.set("Error during conversion.")
            messagebox.showerror("Conversion failed", str(e))


class ToolCombinePDF(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Combine PDF")
        self.resizable(False, False)
        self.folder = tk.StringVar()
        c = ttk.Frame(self, padding=12)
        c.grid(sticky="nsew")
        ttk.Entry(c, textvariable=self.folder, width=60).grid(row=0, column=0, padx=(0, 8), pady=(0, 8), sticky="ew")
        ttk.Button(c, text="Browse", command=self.browse_folder).grid(row=0, column=1, pady=(0, 8))
        ttk.Button(c, text="Combine", command=self.combine).grid(row=1, column=0, padx=(0, 8), sticky="w")
        ttk.Button(c, text="Close", command=self.destroy).grid(row=1, column=1, sticky="w")
        self.status = tk.StringVar(value="Select a folder with PDFs. Files will be merged in name order.")
        ttk.Label(c, textvariable=self.status, relief="sunken", anchor="w").grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        c.columnconfigure(0, weight=1)

    def browse_folder(self):
        d = filedialog.askdirectory(title="Select folder with PDFs")
        if d:
            self.folder.set(d)
            self.status.set("Selected: " + d)

    def combine(self):
        d = self.folder.get().strip()
        if not d or not os.path.isdir(d):
            messagebox.showerror("Invalid folder", "Please select a valid folder.")
            return
        files = [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(".pdf")]
        if not files:
            messagebox.showwarning("No PDFs", "No PDF files found in the selected folder.")
            return
        files.sort(key=lambda x: os.path.basename(x).lower())
        save_path = filedialog.asksaveasfilename(
            title="Save merged PDF as", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile="merged.pdf"
        )
        if not save_path:
            return
        try:
            self.status.set("Merging PDFs...")
            self.update_idletasks()
            merger = PdfMerger()
            for f in files:
                merger.append(f)
            merger.write(save_path)
            merger.close()
            self.status.set(f"Merged {len(files)} PDF(s) to: {save_path}")
            messagebox.showinfo("Completed", f"Merged {len(files)} file(s).\n{save_path}")
        except Exception as e:
            self.status.set("Error during merge.")
            messagebox.showerror("Merge failed", str(e))


if __name__ == "__main__":
    Launcher().mainloop()
