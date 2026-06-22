# Facility Inspection Report Generator

Converts a facility inspection Excel export into a branded, print-ready PDF via XeLaTeX.

## Output structure

| Page | Content |
|------|---------|
| 1 | Cover — overall score, stat cards |
| 2 | Charts — Top/Bottom 5 by Zone, Location Type, and Element |
| 3+ | Deficiency Work Orders table with photo strips |

---

## Prerequisites

| Dependency | Notes |
|------------|-------|
| [uv](https://docs.astral.sh/uv/) | Python package runner — handles deps automatically |
| [MiKTeX](https://miktex.org/) | XeLaTeX engine for PDF compilation |
| Segoe UI | Bundled with Windows; required by the LaTeX template |

Python packages (`pandas`, `openpyxl`) are fetched automatically by `uv` on first run.

---

## Setup

### 1. Company config (gitignored)

Create `company.json` in the `v2/` directory:

```json
{
  "name": "Your Company Name",
  "address": "123 Main St, City, ST 00000",
  "phone": "(555) 000-0000"
}
```

### 2. Logo (gitignored)

Place a logo file named `logo.png`, `logo.jpg`, or `logo.svg` in the `v2/` directory. It appears in the running page header at 1cm height.

---

## Usage

### GUI (recommended)

Double-click **`Report Generator.bat`** — a window opens with file pickers and a live progress log.

### CLI

```bash
uv run --with pandas --with openpyxl generate_report_latex.py <input.xlsx> [output.pdf]
```

Output defaults to `<input>_report.pdf` in the same directory as the input file.

---

## File layout

```
v2/
├── generate_report_latex.py   # Core generator — also importable as a library
├── report_gui.py              # tkinter GUI
├── Report Generator.bat       # Double-click launcher (Windows)
├── company.json               # ← create this (gitignored)
├── logo.png                   # ← add your logo (gitignored)
└── README.md
```

---

## Configuration

### Score thresholds

In `generate_report_latex.py`, `score_color()` maps scores to bar/cell colours:

```python
def score_color(score: float) -> str:
    if score >= 85: return GREEN   # #85CF5F
    if score >= 80: return AMBER   # #F0B557
    return RED                     # #F05773
```

### Goal line

Charts show a dashed vertical reference line at **80%** by default. Change the `0.80` value in `render_bar_chart_tex()` to adjust.

### Top/Bottom N

Charts display the top N and bottom N entries. Default is **5**. Pass a different `n` to `render_bar_chart_tex()` to override.

---

## Privacy & gitignore

The following are excluded from version control:

- `*.xlsx` — source inspection data
- `*.pdf`, `*.tex`, `*.aux`, `*.log` — generated artifacts
- `company.json`, `logo.*` — company branding

Generated `.tex` files contain venue and inspector names and are never committed.

---

## Roadmap

- [ ] Portable distribution (portable MiKTeX + PyInstaller `.exe`, no admin required)
- [ ] Previous-period comparison (delta scores between two Excel exports)
- [ ] Cover page summary text auto-generated from data
