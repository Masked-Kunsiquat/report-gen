"""
report_gui.py  —  Desktop GUI for the Facility Inspection Report generator.

Usage:
    uv run --with pandas --with openpyxl report_gui.py
"""

import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import generate_report_latex as gen

_company = gen.COMPANY.get("name", "") or "Inspection"

# ---------------------------------------------------------------------------
# Brand colours
# ---------------------------------------------------------------------------
NAVY  = "#25408F"
SLATE = "#4D5B82"
GRAY  = "#F0F3F6"
WHITE = "#FFFFFF"
GREEN = "#85CF5F"
RED   = "#F05773"


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{_company} — Report Generator")
        self.geometry("700x520")
        self.resizable(False, False)
        self.configure(bg=WHITE)
        self._build_ui()
        self._running = False

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self, bg=NAVY)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text=f"{_company} — Facility Inspection Report Generator",
            font=("Segoe UI", 13, "bold"), fg=WHITE, bg=NAVY, pady=14
        ).pack(side="left", padx=20)

        # Body
        body = tk.Frame(self, bg=WHITE, padx=24, pady=20)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        # --- Excel file ---
        tk.Label(body, text="EXCEL FILE", font=("Segoe UI", 8, "bold"),
                 fg=SLATE, bg=WHITE, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w")

        self._input_var = tk.StringVar()
        input_row = tk.Frame(body, bg=WHITE)
        input_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 14))
        input_row.columnconfigure(0, weight=1)

        tk.Entry(input_row, textvariable=self._input_var,
                 font=("Segoe UI", 10), relief="solid", bd=1,
                 highlightthickness=0).grid(row=0, column=0, sticky="ew", ipady=5)
        tk.Button(input_row, text="Browse…", command=self._browse_input,
                  font=("Segoe UI", 9), bg=GRAY, fg=SLATE,
                  relief="flat", padx=12, pady=5, cursor="hand2").grid(
            row=0, column=1, padx=(8, 0))

        # --- Output folder ---
        tk.Label(body, text="OUTPUT FOLDER", font=("Segoe UI", 8, "bold"),
                 fg=SLATE, bg=WHITE, anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="w")

        self._output_var = tk.StringVar()
        out_row = tk.Frame(body, bg=WHITE)
        out_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 18))
        out_row.columnconfigure(0, weight=1)

        tk.Entry(out_row, textvariable=self._output_var,
                 font=("Segoe UI", 10), relief="solid", bd=1,
                 highlightthickness=0).grid(row=0, column=0, sticky="ew", ipady=5)
        tk.Button(out_row, text="Browse…", command=self._browse_output,
                  font=("Segoe UI", 9), bg=GRAY, fg=SLATE,
                  relief="flat", padx=12, pady=5, cursor="hand2").grid(
            row=0, column=1, padx=(8, 0))

        # --- Generate button ---
        self._gen_btn = tk.Button(
            body, text="Generate Report",
            font=("Segoe UI", 11, "bold"),
            bg=NAVY, fg=WHITE, activebackground=SLATE, activeforeground=WHITE,
            relief="flat", pady=10, cursor="hand2",
            command=self._generate
        )
        self._gen_btn.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 14))

        # --- Log area ---
        tk.Label(body, text="LOG", font=("Segoe UI", 8, "bold"),
                 fg=SLATE, bg=WHITE, anchor="w").grid(
            row=5, column=0, columnspan=2, sticky="w")

        self._log_box = scrolledtext.ScrolledText(
            body, height=10, font=("Consolas", 9),
            state="disabled", bg=GRAY, relief="flat",
            wrap="word", padx=8, pady=8
        )
        self._log_box.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
        body.rowconfigure(6, weight=1)

        # Tag colours for log messages
        self._log_box.tag_config("ok",    foreground=GREEN)
        self._log_box.tag_config("err",   foreground=RED)
        self._log_box.tag_config("plain", foreground="#333333")

    # ------------------------------------------------------------------
    # File pickers
    # ------------------------------------------------------------------
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select inspection Excel export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self._input_var.set(path)
            if not self._output_var.get():
                self._output_var.set(str(Path(path).parent))

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self._output_var.set(path)

    # ------------------------------------------------------------------
    # Logging (thread-safe via after())
    # ------------------------------------------------------------------
    def _log(self, msg: str):
        if "[Done]" in msg:
            tag = "ok"
        elif "ERROR" in msg or "failed" in msg.lower():
            tag = "err"
        else:
            tag = "plain"
        self.after(0, lambda m=msg, t=tag: self._append(m, t))

    def _append(self, msg: str, tag: str):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n", tag)
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    # ------------------------------------------------------------------
    # Generation
    # ------------------------------------------------------------------
    def _generate(self):
        if self._running:
            return

        input_file = self._input_var.get().strip()
        if not input_file or not Path(input_file).exists():
            self._log("ERROR: Please select a valid Excel file.")
            return

        out_dir    = self._output_var.get().strip() or str(Path(input_file).parent)
        stem       = Path(input_file).stem
        output_pdf = str(Path(out_dir) / f"{stem}_report.pdf")

        self._running = True
        self._gen_btn.configure(state="disabled", text="Generating…")
        self._append("─" * 60, "plain")

        def _worker():
            try:
                gen.generate_report(input_file, output_pdf, log=self._log)
                self.after(0, lambda: self._finish(True))
            except Exception as exc:
                self._log(f"ERROR: {exc}")
                self.after(0, lambda: self._finish(False))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish(self, success: bool):
        self._running = False
        self._gen_btn.configure(state="normal", text="Generate Report")
        if not success:
            self._append("✗  Generation failed — see log above.", "err")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
