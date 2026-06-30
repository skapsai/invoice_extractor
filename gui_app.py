"""
gui_app.py
----------
Desktop GUI for the Invoice Extractor pipeline.

Built with: Python · Tkinter · ttkbootstrap (theme: darkly)

Features:
  - Upload a single PDF or a ZIP containing PDFs
  - Background processing (UI stays responsive)
  - Progress spinner + status label
  - "Save Excel" button (enabled after successful extraction)
  - Every run creates a fresh Excel (no appending)
  - All verbose logs go to invoice_extraction.log on disk (no log box in UI)

Usage:
  python gui_app.py
"""

import queue
import shutil
import tempfile
import threading
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# ── Import the extraction engine ──────────────────────────────────────────────
from invoice_extractor import (
    EXCEL_COLUMNS,
    EXTRACTED_TEXT_DIR,
    INVOICES_DIR,
    OUTPUT_EXCEL,
    process_pdf,
    write_excel,
)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

APP_TITLE   = "Invoice Extractor"
APP_THEME   = "darkly"
WIN_WIDTH   = 520
WIN_HEIGHT  = 380

# Accepted extensions
ACCEPTED_EXT = {".pdf", ".zip"}


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APPLICATION CLASS
# ─────────────────────────────────────────────────────────────────────────────

class InvoiceExtractorApp:
    """Main GUI window for the Invoice Extractor."""

    def __init__(self, root: ttk.Window) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.resizable(False, False)

        # State
        self._selected_files: list[Path] = []
        self._output_excel: Path | None  = None
        self._msg_queue: queue.Queue     = queue.Queue()

        self._build_ui()
        self._center_window()

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        """Assemble all widgets."""
        root = self.root
        style = ttk.Style()

        # ── Header banner ─────────────────────────────────────────────────────
        banner = ttk.Frame(root, bootstyle="primary", padding=(0, 16))
        banner.pack(fill=X)

        ttk.Label(
            banner,
            text="🧾  Invoice Extractor",
            font=("Segoe UI", 20, "bold"),
            bootstyle="inverse-primary",
            anchor=CENTER,
        ).pack(fill=X)

        ttk.Label(
            banner,
            text="PDF → Structured Excel in one click",
            font=("Segoe UI", 10),
            bootstyle="inverse-primary",
            foreground="#adb5bd",
            anchor=CENTER,
        ).pack(fill=X)

        # ── Main card ─────────────────────────────────────────────────────────
        card = ttk.Frame(root, padding=(32, 16))
        card.pack(fill=BOTH, expand=True)

        # Upload button
        self._btn_upload = ttk.Button(
            card,
            text="📂  Upload PDF(s) or ZIP",
            bootstyle="primary-outline",
            width=28,
            command=self._on_upload,
        )
        self._btn_upload.pack(pady=(0, 10))

        # Selected file label
        self._lbl_file = ttk.Label(
            card,
            text="No files selected",
            font=("Segoe UI", 9),
            foreground="#6c757d",
            anchor=CENTER,
            wraplength=440,
        )
        self._lbl_file.pack(pady=(0, 10))

        # Separator
        ttk.Separator(card, orient=HORIZONTAL).pack(fill=X, pady=(0, 15))

        # Status label
        self._lbl_status = ttk.Label(
            card,
            text="",
            font=("Segoe UI", 10, "bold"),
            anchor=CENTER,
            wraplength=440,
        )
        self._lbl_status.pack(pady=(0, 8))

        # Progress bar (indeterminate)
        self._progress = ttk.Progressbar(
            card,
            bootstyle="primary-striped",
            mode="indeterminate",
            length=380,
        )
        self._progress.pack(pady=(0, 16))

        # Download button
        self._btn_save = ttk.Button(
            card,
            text="💾  Save Excel",
            bootstyle="success",
            width=22,
            state=DISABLED,
            command=self._on_download,
        )
        self._btn_save.pack()

        # ── Footer ────────────────────────────────────────────────────────────
        footer = ttk.Frame(root, padding=(0, 4))
        footer.pack(fill=X, side=BOTTOM)
        ttk.Label(
            footer,
            text="Logs are saved to  invoice_extraction.log",
            font=("Segoe UI", 8),
            foreground="#6c757d",
            anchor=CENTER,
        ).pack(fill=X)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _center_window(self) -> None:
        """Center the window on screen."""
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x  = (sw - WIN_WIDTH)  // 2
        y  = (sh - WIN_HEIGHT) // 2
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{x}+{y}")

    def _set_status(self, text: str, color: str = "#f8f9fa") -> None:
        """Update the status label (thread-safe via .after scheduling)."""
        self._lbl_status.configure(text=text, foreground=color)

    def _set_ui_busy(self, busy: bool) -> None:
        """Lock/unlock interactive controls during processing."""
        state = DISABLED if busy else NORMAL
        self._btn_upload.configure(state=state)
        if busy:
            self._btn_save.configure(state=DISABLED)
            self._progress.start(10)
        else:
            self._progress.stop()
            self._progress["value"] = 0

    # ── File Upload ───────────────────────────────────────────────────────────

    def _on_upload(self) -> None:
        """Open file dialog, validate selection, then start processing."""
        paths = filedialog.askopenfilenames(
            title="Select Invoice PDF(s) or ZIP",
            filetypes=[
                ("PDF / ZIP files", "*.pdf *.zip"),
                ("PDF files",       "*.pdf"),
                ("ZIP archives",    "*.zip"),
            ],
        )
        if not paths:
            return  # User cancelled

        selected_files = [Path(p) for p in paths]

        # Validate extension
        for selected in selected_files:
            if selected.suffix.lower() not in ACCEPTED_EXT:
                messagebox.showerror(
                    "Unsupported Format",
                    f"Unsupported file format in '{selected.name}'. Please upload a PDF or a ZIP containing PDFs.",
                )
                return

        self._selected_files = selected_files
        
        # Display selection
        if len(selected_files) == 1:
            lbl_text = f"📄  {selected_files[0].name}"
        else:
            lbl_text = f"📄  {len(selected_files)} files selected"
            
        self._lbl_file.configure(
            text=lbl_text,
            foreground="#e9ecef",
        )

        # Reset state from previous run
        self._output_excel = None
        self._btn_save.configure(state=DISABLED)
        self._set_status("")

        # Start background worker
        self._set_ui_busy(True)
        thread = threading.Thread(target=self._worker, daemon=True)
        thread.start()

        # Begin polling the message queue
        self.root.after(100, self._poll_queue)

    # ── Background Worker ─────────────────────────────────────────────────────

    def _worker(self) -> None:
        """
        Runs in a daemon thread.
        Handles PDF or ZIP input, calls the extractor engine, writes Excel.
        Communicates back to the GUI via self._msg_queue.
        """
        temp_dirs: list[Path] = []
        staged: list[Path] = []
        try:
            self._msg_queue.put(("status", "⏳  Processing…", "#f8f9fa"))

            pdf_paths = []
            for path in self._selected_files:
                pdf_paths.extend(self._resolve_pdfs(path, temp_dirs))

            if not pdf_paths:
                self._msg_queue.put(("error", "❌  No valid PDFs found in the uploaded file(s)."))
                return

            # Ensure required directories exist
            for folder in [INVOICES_DIR, EXTRACTED_TEXT_DIR, OUTPUT_EXCEL.parent]:
                folder.mkdir(parents=True, exist_ok=True)

            # Copy PDFs into the invoices/ folder (for engine to read)
            for src in pdf_paths:
                dest = INVOICES_DIR / src.name
                shutil.copy2(str(src), str(dest))
                staged.append(dest)

            # Extract records from each staged PDF
            all_records: list[dict] = []
            for staged_pdf in staged:
                try:
                    records = process_pdf(staged_pdf)
                    all_records.extend(records)
                except Exception as e:
                    import logging
                    logging.getLogger("invoice_extractor").error(f"Error processing {staged_pdf.name}: {e}")
                finally:
                    # Delete the copied PDF immediately after processing
                    if staged_pdf.exists():
                        try:
                            staged_pdf.unlink()
                        except Exception:
                            pass

            # Clean up generated raw text files in EXTRACTED_TEXT_DIR
            for staged_pdf in staged:
                stem = staged_pdf.stem
                for txt_file in EXTRACTED_TEXT_DIR.glob(f"{stem}__*.txt"):
                    try:
                        txt_file.unlink()
                    except Exception:
                        pass

            if not all_records:
                self._msg_queue.put(("error", "⚠️  Processing finished but no records were extracted."))
                return

            # Fresh Excel every run — overwrite any existing file
            write_excel(all_records, OUTPUT_EXCEL)
            self._output_excel = OUTPUT_EXCEL

            n = len(all_records)
            self._msg_queue.put(("done", f"✅  Done — {n} record{'s' if n != 1 else ''} extracted."))

        except Exception as exc:
            # Short message in UI; full traceback goes to the log file
            short = str(exc)[:120]
            self._msg_queue.put(("error", f"❌  Error: {short}"))
        finally:
            # Clean up temporary folders from ZIP extraction
            for tmp in temp_dirs:
                if tmp.exists():
                    try:
                        shutil.rmtree(tmp)
                    except Exception:
                        pass

    def _resolve_pdfs(self, path: Path, temp_dirs: list[Path]) -> list[Path]:
        """
        Given the uploaded path, return a list of PDF Path objects to process.
          - PDF  → [path]
          - ZIP  → extract all *.pdf members to a temp dir, return those paths
        """
        if path.suffix.lower() == ".pdf":
            return [path]

        # ZIP: extract PDFs to a temporary directory
        tmp_dir = Path(tempfile.mkdtemp(prefix="invoice_zip_"))
        temp_dirs.append(tmp_dir)
        pdf_list: list[Path] = []
        with zipfile.ZipFile(str(path), "r") as zf:
            for member in zf.namelist():
                if member.lower().endswith(".pdf") and not member.startswith("__MACOSX"):
                    # Extract preserving only the filename (strip sub-folders)
                    filename = Path(member).name
                    target   = tmp_dir / filename
                    with zf.open(member) as src, open(target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
                    pdf_list.append(target)

        return pdf_list

    # ── Queue Polling ─────────────────────────────────────────────────────────

    def _poll_queue(self) -> None:
        """
        Called by the Tkinter event loop every 100 ms.
        Drains messages from the worker thread and updates the GUI.
        """
        try:
            while True:
                msg = self._msg_queue.get_nowait()
                kind = msg[0]

                if kind == "status":
                    _, text, color = msg
                    self._set_status(text, color)

                elif kind == "done":
                    _, text = msg
                    self._set_status(text, "#2fb344")   # green
                    self._set_ui_busy(False)
                    self._btn_save.configure(state=NORMAL)
                    return  # Stop polling

                elif kind == "error":
                    _, text = msg
                    self._set_status(text, "#e5383b")   # red
                    self._set_ui_busy(False)
                    return  # Stop polling

        except queue.Empty:
            pass

        # Re-schedule if worker is still running
        self.root.after(100, self._poll_queue)

    # ── Excel Download ────────────────────────────────────────────────────────

    def _on_download(self) -> None:
        """Open a Save-As dialog so the user can copy the Excel output anywhere."""
        if not self._output_excel or not self._output_excel.exists():
            messagebox.showerror("File Not Found", "The output Excel file could not be found.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save Extracted Excel",
            defaultextension=".xlsx",
            initialfile="extracted_invoice_records.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not save_path:
            return  # User cancelled

        shutil.copy2(str(self._output_excel), save_path)
        messagebox.showinfo(
            "Saved",
            f"Excel file saved successfully:\n{save_path}",
        )


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    root = ttk.Window(
        title=APP_TITLE,
        themename=APP_THEME,
        size=(WIN_WIDTH, WIN_HEIGHT),
        resizable=(False, False),
    )
    app = InvoiceExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
