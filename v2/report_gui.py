"""
report_gui.py  —  Desktop GUI for the Facility Inspection Report generator.

Usage:
    uv run --with pandas --with openpyxl report_gui.py
"""

import os
import re
import sys
import time
import threading
import traceback
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
AMBER = "#C98A2E"

_STEP_RE = re.compile(r"^\s*\[\d+\]")


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{_company} — Report Generator")
        self.geometry("720x600")
        self.resizable(False, True)
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

        # Scrollable body
        outer = tk.Frame(self, bg=WHITE)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=WHITE, highlightthickness=0)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=WHITE, padx=24, pady=20)
        body_window = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_resize(event):
            canvas.itemconfig(body_window, width=event.width)
        canvas.bind("<Configure>", _on_resize)

        def _on_body_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        body.bind("<Configure>", _on_body_resize)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

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

        # --- Comments (optional, always visible) ---
        tk.Label(body, text="COMMENTS (OPTIONAL)", font=("Segoe UI", 8, "bold"),
                 fg=SLATE, bg=WHITE, anchor="w").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(0, 4))

        comments_frame = tk.Frame(body, bg=WHITE)
        comments_frame.grid(row=5, column=0, columnspan=2, sticky="ew")
        comments_frame.columnconfigure(0, weight=1)

        self._comment_boxes = {}
        for i, (key, label) in enumerate([
            ("client",      "Client Comments"),
            ("maintenance", "Maintenance Notes"),
            ("general",     "General Remarks"),
        ]):
            tk.Label(comments_frame, text=label.upper(),
                     font=("Segoe UI", 8), fg=SLATE, bg=WHITE, anchor="w"
                     ).grid(row=i*2, column=0, sticky="w", pady=(6 if i else 0, 2))
            box = tk.Text(comments_frame, height=3, font=("Segoe UI", 9),
                          relief="solid", bd=1, wrap="word", padx=6, pady=4)
            box.grid(row=i*2+1, column=0, sticky="ew", pady=(0, 2))
            self._comment_boxes[key] = box

        # --- Generate button ---
        self._gen_btn = tk.Button(
            body, text="Generate Report",
            font=("Segoe UI", 11, "bold"),
            bg=NAVY, fg=WHITE, activebackground=SLATE, activeforeground=WHITE,
            relief="flat", pady=10, cursor="hand2",
            command=self._generate
        )
        self._gen_btn.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(12, 14))

        # --- Log area ---
        tk.Label(body, text="LOG", font=("Segoe UI", 8, "bold"),
                 fg=SLATE, bg=WHITE, anchor="w").grid(
            row=7, column=0, columnspan=2, sticky="w")

        self._log_box = scrolledtext.ScrolledText(
            body, height=8, font=("Consolas", 9),
            state="disabled", bg=GRAY, relief="flat",
            wrap="word", padx=8, pady=8
        )
        self._log_box.grid(row=8, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
        body.rowconfigure(8, weight=1)

        # Tag colours for log messages
        self._log_box.tag_config("ok",    foreground="#3C9A2E")
        self._log_box.tag_config("err",   foreground=RED)
        self._log_box.tag_config("warn",  foreground=AMBER)
        self._log_box.tag_config("step",  foreground=NAVY, font=("Consolas", 9, "bold"))
        self._log_box.tag_config("time",  foreground="#9AA3B2")
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
    @staticmethod
    def _classify(msg: str) -> str:
        low = msg.lower()
        if "[done]" in low:
            return "ok"
        if "error" in low or "failed" in low or msg.strip().startswith("!"):
            return "err"
        if "hint" in low or "[warn]" in low or "missing" in low:
            return "warn"
        if _STEP_RE.match(msg):
            return "step"
        return "plain"

    def _log(self, msg: str):
        ts  = time.strftime("%H:%M:%S")
        tag = self._classify(msg)
        self.after(0, lambda m=msg, t=tag, s=ts: self._append(m, t, s))

    def _append(self, msg: str, tag: str = "plain", ts: str = None):
        self._log_box.configure(state="normal")
        if ts:
            self._log_box.insert("end", f"{ts}  ", "time")
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

        comments = {k: v.get("1.0", "end-1c") for k, v in self._comment_boxes.items()}

        def _worker():
            try:
                gen.generate_report(input_file, output_pdf, log=self._log, comments=comments)
                self.after(0, lambda: self._finish(True, output_pdf))
            except Exception as exc:
                self._log(f"ERROR: {type(exc).__name__}: {exc}")
                for line in traceback.format_exc().strip().splitlines()[-4:]:
                    self._log(f"    {line}")
                self.after(0, lambda: self._finish(False, None))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish(self, success: bool, output_pdf: str = None):
        self._running = False
        self._gen_btn.configure(state="normal", text="Generate Report")
        if success:
            self._append("✓  Report generated successfully.", "ok")
            if output_pdf and Path(output_pdf).exists() and hasattr(os, "startfile"):
                try:
                    os.startfile(output_pdf)  # open in default PDF viewer (Windows)
                except OSError:
                    pass
        else:
            self._append("✗  Generation failed — see log above.", "err")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
