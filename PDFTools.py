import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfMerger

# Windows-only assumption; Poppler must be on PATH (pdftoppm).
# One file, three small windows launched from a simple launcher.


class Launcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Tools (Windows)")
        self.resizable(False, False)
        c = ttk.Frame(self, padding=16)
        c.grid(sticky="nsew")
        ttk.Label(c, text="Choose a tool:").grid(
            row=0, column=0, columnspan=3, pady=(0, 10)
        )
        ttk.Button(c, text="PDF → JPEG", command=self.open_pdf2jpg).grid(
            row=1, column=0, padx=4, pady=4, sticky="ew"
        )
        ttk.Button(c, text="JPEG → PDF", command=self.open_jpg2pdf).grid(
            row=1, column=1, padx=4, pady=4, sticky="ew"
        )
        ttk.Button(c, text="Combine PDF", command=self.open_combine).grid(
            row=1, column=2, padx=4, pady=4, sticky="ew"
        )
        for i in range(3):
            c.columnconfigure(i, weight=1)

    def open_pdf2jpg(self):
        ToolPDF2JPG(self)

    def open_jpg2pdf(self):
        ToolJPG2PDF(self)

    def open_combine(self):
        ToolCombinePDF(self)


class ToolPDF2JPG(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("PDF → JPEG")
        self.resizable(False, False)
        self.pdf_path = tk.StringVar()
        c = ttk.Frame(self, padding=12)
        c.grid(sticky="nsew")
        ttk.Entry(c, textvariable=self.pdf_path, width=60).grid(
            row=0, column=0, padx=(0, 8), pady=(0, 8), sticky="ew"
        )
        ttk.Button(c, text="Browse", command=self.browse).grid(
            row=0, column=1, pady=(0, 8)
        )
        ttk.Button(c, text="Convert", command=self.convert).grid(
            row=1, column=0, padx=(0, 8), sticky="w"
        )
        ttk.Button(c, text="Close", command=self.destroy).grid(
            row=1, column=1, sticky="w"
        )
        self.status = tk.StringVar(value="Ready")
        ttk.Label(c, textvariable=self.status, relief="sunken", anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
        c.columnconfigure(0, weight=1)

    def browse(self):
        f = filedialog.askopenfilename(
            title="Select PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if f:
            self.pdf_path.set(f)
            self.status.set("Selected: " + os.path.basename(f))

    def convert(self):
        if not shutil.which("pdftoppm"):
            messagebox.showerror(
                "Poppler not found",
                "pdftoppm not found on PATH. Install Poppler for Windows and add its bin to PATH.",
            )
            return
        pdf = self.pdf_path.get().strip()
        if not pdf or not os.path.isfile(pdf) or not pdf.lower().endswith(".pdf"):
            messagebox.showerror("Invalid file", "Please select a valid PDF file.")
            return
        base_dir = os.path.dirname(pdf)
        base_name = os.path.splitext(os.path.basename(pdf))[0]
        out_dir = os.path.join(base_dir, f"{base_name}_images")
        os.makedirs(out_dir, exist_ok=True)
        try:
            self.status.set("Converting at 300 DPI...")
            self.update_idletasks()
            ppm_paths = convert_from_path(
                pdf, dpi=300, output_folder=out_dir, fmt="ppm", paths_only=True
            )
            count = 0
            for i, ppm in enumerate(sorted(ppm_paths)):
                out_path = os.path.join(out_dir, f"{base_name}[{i}].jpeg")
                with Image.open(ppm) as im:
                    im.convert("RGB").save(
                        out_path, "JPEG", quality=92, optimize=True, progressive=True
                    )
                count += 1
            self.status.set(f"Done. Saved {count} JPEG(s) to: {out_dir}")
            messagebox.showinfo(
                "Completed", f"Converted {count} page(s).\nFolder: {out_dir}"
            )
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
        ttk.Entry(c, textvariable=self.folder, width=60).grid(
            row=0, column=0, padx=(0, 8), pady=(0, 8), sticky="ew"
        )
        ttk.Button(c, text="Browse", command=self.browse_folder).grid(
            row=0, column=1, pady=(0, 8)
        )
        ttk.Button(c, text="Convert", command=self.convert).grid(
            row=1, column=0, padx=(0, 8), sticky="w"
        )
        ttk.Button(c, text="Close", command=self.destroy).grid(
            row=1, column=1, sticky="w"
        )
        self.status = tk.StringVar(
            value="Select a folder containing JPEG(s). Files will be ordered by name."
        )
        ttk.Label(c, textvariable=self.status, relief="sunken", anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
        c.columnconfigure(0, weight=1)

    def browse_folder(self):
        d = filedialog.askdirectory(title="Select folder with JPEGs")
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
            messagebox.showwarning(
                "No images", "No compatible images found in the selected folder."
            )
            return
        imgs.sort(key=lambda x: os.path.basename(x).lower())
        save_path = filedialog.asksaveasfilename(
            title="Save PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="output.pdf",
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
            first.save(
                save_path, "PDF", resolution=300.0, save_all=True, append_images=rest
            )
            for im in pil_images:
                im.close()
            self.status.set(f"Saved PDF: {save_path}")
            messagebox.showinfo(
                "Completed", f"Created PDF from {len(imgs)} image(s).\n{save_path}"
            )
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
        ttk.Entry(c, textvariable=self.folder, width=60).grid(
            row=0, column=0, padx=(0, 8), pady=(0, 8), sticky="ew"
        )
        ttk.Button(c, text="Browse", command=self.browse_folder).grid(
            row=0, column=1, pady=(0, 8)
        )
        ttk.Button(c, text="Combine", command=self.combine).grid(
            row=1, column=0, padx=(0, 8), sticky="w"
        )
        ttk.Button(c, text="Close", command=self.destroy).grid(
            row=1, column=1, sticky="w"
        )
        self.status = tk.StringVar(
            value="Select a folder with PDFs. Files will be merged in name order."
        )
        ttk.Label(c, textvariable=self.status, relief="sunken", anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
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
        files = [
            os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(".pdf")
        ]
        if not files:
            messagebox.showwarning(
                "No PDFs", "No PDF files found in the selected folder."
            )
            return
        files.sort(key=lambda x: os.path.basename(x).lower())
        save_path = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="merged.pdf",
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
            messagebox.showinfo(
                "Completed", f"Merged {len(files)} file(s).\n{save_path}"
            )
        except Exception as e:
            self.status.set("Error during merge.")
            messagebox.showerror("Merge failed", str(e))


if __name__ == "__main__":
    Launcher().mainloop()
