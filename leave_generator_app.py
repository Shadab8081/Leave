#!/usr/bin/env python3
"""
Leave Document Generator — Desktop App
========================================
A simple offline, double-clickable GUI wrapping generate_leave_documents.py.

Run with:   python3 leave_generator_app.py

No internet connection is used at runtime. The only external tool it calls
is LibreOffice (installed locally) to convert filled Excel/Word files to PDF.
"""

import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import generate_leave_documents as gen


class LeaveGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Leave Document Generator")
        self.geometry("760x560")
        self.minsize(640, 460)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar(value=str(Path.cwd() / "Leave Documents"))
        self.is_running = False

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        header = tk.Label(
            self, text="Automated Leave Application & Clearance Generator",
            font=("Segoe UI", 14, "bold")
        )
        header.pack(anchor="w", **pad)

        subtitle = tk.Label(
            self,
            text="Select your filled employee Excel sheet, choose where to save the "
                 "documents, then click Generate.",
            font=("Segoe UI", 9), fg="#555555", justify="left"
        )
        subtitle.pack(anchor="w", padx=12)

        # --- Input file row ---
        row1 = tk.Frame(self)
        row1.pack(fill="x", **pad)
        tk.Label(row1, text="Input Excel file:", width=16, anchor="w").pack(side="left")
        tk.Entry(row1, textvariable=self.input_path).pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(row1, text="Browse…", command=self.browse_input).pack(side="left")

        # --- Output folder row ---
        row2 = tk.Frame(self)
        row2.pack(fill="x", **pad)
        tk.Label(row2, text="Output folder:", width=16, anchor="w").pack(side="left")
        tk.Entry(row2, textvariable=self.output_path).pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(row2, text="Browse…", command=self.browse_output).pack(side="left")

        # --- Generate button + progress ---
        row3 = tk.Frame(self)
        row3.pack(fill="x", **pad)
        self.generate_btn = tk.Button(
            row3, text="Generate Leave Documents", font=("Segoe UI", 11, "bold"),
            bg="#2e7d32", fg="white", activebackground="#256428",
            command=self.on_generate_clicked, height=2
        )
        self.generate_btn.pack(side="left", fill="x", expand=True)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12, pady=(0, 6))

        # --- Log box ---
        log_frame = tk.Frame(self)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        tk.Label(log_frame, text="Log:", anchor="w").pack(anchor="w")

        text_frame = tk.Frame(log_frame)
        text_frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        self.log_box = tk.Text(
            text_frame, height=14, wrap="word", yscrollcommand=scrollbar.set,
            font=("Consolas", 9), bg="#111111", fg="#dddddd", insertbackground="white"
        )
        self.log_box.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log_box.yview)

        open_row = tk.Frame(self)
        open_row.pack(fill="x", padx=12, pady=(0, 12))
        tk.Button(open_row, text="Open Output Folder", command=self.open_output_folder).pack(side="left")

    # ------------------------------------------------------------------
    def browse_input(self):
        path = filedialog.askopenfilename(
            title="Select employee Excel sheet",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="Select (or create) output folder")
        if path:
            self.output_path.set(path)

    def open_output_folder(self):
        out = Path(self.output_path.get())
        out.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            import os
            os.startfile(out)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.run(["open", str(out)])
        else:
            import subprocess
            subprocess.run(["xdg-open", str(out)])

    def log(self, msg):
        def append():
            self.log_box.insert("end", str(msg) + "\n")
            self.log_box.see("end")
        self.after(0, append)

    # ------------------------------------------------------------------
    def on_generate_clicked(self):
        if self.is_running:
            return
        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()

        if not input_path:
            messagebox.showwarning("Missing file", "Please select an input Excel file first.")
            return
        if not Path(input_path).exists():
            messagebox.showerror("File not found", f"Could not find:\n{input_path}")
            return
        if not output_path:
            messagebox.showwarning("Missing folder", "Please choose an output folder.")
            return

        self.log_box.delete("1.0", "end")
        self.is_running = True
        self.generate_btn.config(state="disabled", text="Generating…")
        self.progress.start(12)

        thread = threading.Thread(
            target=self._run_generation_thread, args=(input_path, output_path), daemon=True
        )
        thread.start()

    def _run_generation_thread(self, input_path, output_path):
        try:
            generated, errors = gen.run_generation(input_path, output_path, log=self.log)
            self.after(0, lambda: self._on_done(len(generated), len(errors)))
        except Exception as exc:
            tb = traceback.format_exc()
            self.log("ERROR:\n" + tb)
            self.after(0, lambda: self._on_failed(str(exc)))

    def _on_done(self, n_ok, n_err):
        self.progress.stop()
        self.generate_btn.config(state="normal", text="Generate Leave Documents")
        self.is_running = False
        if n_err:
            messagebox.showwarning(
                "Done — with some issues",
                f"Generated documents for {n_ok} employee(s).\n"
                f"{n_err} row(s) had problems — see Error_Report.txt in the output folder "
                f"and the log above."
            )
        else:
            messagebox.showinfo("Done", f"Generated documents for {n_ok} employee(s). All rows succeeded.")

    def _on_failed(self, err):
        self.progress.stop()
        self.generate_btn.config(state="normal", text="Generate Leave Documents")
        self.is_running = False
        messagebox.showerror("Generation failed", f"Something went wrong:\n\n{err}")


if __name__ == "__main__":
    app = LeaveGeneratorApp()
    app.mainloop()
