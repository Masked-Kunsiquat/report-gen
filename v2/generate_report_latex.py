"""
generate_report_latex.py
Reads an inspection Excel export and produces a branded PDF via XeLaTeX.
Charts are rendered as PNGs by matplotlib and included via \\includegraphics.

Usage:
    uv run --with pandas --with openpyxl --with matplotlib generate_report_latex.py <input.xlsx> [output.pdf]
"""

import sys
import re
import json
import base64
import shutil
import tempfile
import zipfile
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches


# ---------------------------------------------------------------------------
# XeLaTeX discovery
# ---------------------------------------------------------------------------

XELATEX_CANDIDATES = [
    "xelatex",
    r"C:\Users\user\AppData\Local\Programs\MiKTeX\miktex\bin\x64\xelatex.exe",
    r"C:\Program Files\MiKTeX\miktex\bin\x64\xelatex.exe",
    r"C:\Program Files (x86)\MiKTeX\miktex\bin\xelatex.exe",
]

def find_xelatex() -> str:
    for candidate in XELATEX_CANDIDATES:
        if shutil.which(candidate) or Path(candidate).exists():
            return candidate
    raise RuntimeError(
        "xelatex not found. Install MiKTeX and ensure it is on PATH."
    )


# ---------------------------------------------------------------------------
# Brand constants
# ---------------------------------------------------------------------------

NAVY   = "#25408F"
SKY    = "#00B0F0"
GRAY   = "#F0F3F6"
GREEN  = "#85CF5F"
AMBER  = "#F0B557"
RED    = "#F05773"
SLATE  = "#4D5B82"


def score_color(score: float) -> str:
    if score >= 85: return GREEN
    if score >= 80: return AMBER
    return RED


def hex_to_rgb(h: str) -> tuple:
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# Company config  (gitignored — same as HTML version)
# ---------------------------------------------------------------------------

def load_company(script_path: Path) -> dict:
    root = script_path.parent
    company = {"name": "", "address": "", "phone": "", "logo_path": ""}
    config = root / "company.json"
    if config.exists():
        with open(config, encoding="utf-8") as f:
            company.update(json.load(f))
    for ext in ("png", "jpg", "jpeg", "svg"):
        p = root / f"logo.{ext}"
        if p.exists():
            company["logo_path"] = str(p)
            break
    return company

COMPANY = load_company(Path(__file__))


# ---------------------------------------------------------------------------
# Data loading  (identical logic to HTML version)
# ---------------------------------------------------------------------------

def load_inspections(file_path: str):
    filters_df = pd.read_excel(file_path, sheet_name="Filters", header=None, engine="openpyxl")
    filters = dict(zip(filters_df[0].str.strip(), filters_df[1]))
    df = pd.read_excel(file_path, sheet_name="Raw Data", engine="openpyxl")
    complete = df[df["Status"].fillna("").str.strip().str.lower() == "complete"].copy()
    complete.reset_index(drop=True, inplace=True)
    return complete, filters, df


def comment_col(df: pd.DataFrame) -> str:
    return "Comments" if "Comments" in df.columns else "Comment"


def parse_location(location: str) -> tuple:
    location = str(location).strip()
    for sep in (" - ", "_ "):
        if sep in location:
            _, space = location.rsplit(sep, 1)
            return "", "", space.strip()
    return "", "", location


def drop_image_stubs(df: pd.DataFrame) -> pd.DataFrame:
    stub_pattern = re.compile(r"^\s*\(([^)]+)\)\s*$")
    def is_stub(comment, group_comments):
        m = stub_pattern.match(comment)
        if not m:
            return False
        tag = m.group(1)
        return any(
            c != comment and re.search(r"\(" + re.escape(tag) + r"\)\s*$", c)
            for c in group_comments
        )
    drop_idx = set()
    cc = comment_col(df)
    for _, group in df.groupby(["Inspection #", "Element"], sort=False):
        comments = group[cc].fillna("").tolist()
        for idx, comment in zip(group.index, comments):
            if is_stub(comment, comments):
                drop_idx.add(idx)
    return df[~df.index.isin(drop_idx)].reset_index(drop=True)


def build_summary(df: pd.DataFrame, filters: dict) -> dict:
    per_insp = (
        df.groupby("Inspection #", sort=False)
        .agg(
            score=("Score In %", "first"),
            location=("Location", "first"),
            completion=("Completion Date", "first"),
            completed_by=("Completed By", "first"),
            venue_=("Venue", "first"),
            building_=("Building", "first"),
            corporation_=("Corporation", "first"),
        )
        .reset_index()
    )

    overall_score = round(per_insp["score"].mean(), 2) if not per_insp.empty else 0
    venue = per_insp["venue_"].iloc[0] if not per_insp.empty else "Unknown Venue"
    building = per_insp["building_"].iloc[0] if not per_insp.empty else ""
    corporation = per_insp["corporation_"].iloc[0] if not per_insp.empty else ""

    date_from = pd.to_datetime(filters.get("From Date"), errors="coerce")
    date_to   = pd.to_datetime(filters.get("To Date"),   errors="coerce")
    if pd.notna(date_from) and pd.notna(date_to):
        if date_from.date() == date_to.date():
            date_str = date_from.strftime("%B %d, %Y")
        else:
            date_str = f"{date_from.strftime('%B %d, %Y')} – {date_to.strftime('%B %d, %Y')}"
    else:
        date_str = ""

    per_insp["_space"] = per_insp["location"].apply(lambda l: parse_location(str(l))[2])
    space_scores = per_insp.groupby("_space")["score"].mean().round(2).sort_values(ascending=True)

    insp_zone = (
        df.groupby("Inspection #")["Zone"].first()
        .reset_index().rename(columns={"Zone": "zone"})
    )
    per_insp_z = per_insp.merge(insp_zone, on="Inspection #", how="left")
    zone_scores = per_insp_z.groupby("zone")["score"].mean().round(2).sort_values(ascending=True)

    _cc = comment_col(df)
    work_order_rows = df[df[_cc].notna() & (df[_cc].astype(str).str.strip() != "")].copy()
    work_order_rows = work_order_rows.rename(columns={_cc: "Comment"})
    work_order_rows = drop_image_stubs(work_order_rows)

    insp_meta = {
        row["Inspection #"]: {
            "completion": pd.to_datetime(row["completion"], errors="coerce"),
            "completed_by": str(row["completed_by"]).strip() if pd.notna(row["completed_by"]) else "",
        }
        for _, row in per_insp.iterrows()
    }

    return {
        "venue": venue, "building": building, "corporation": corporation,
        "date_str": date_str, "overall_score": overall_score,
        "total_inspections": len(per_insp),
        "total_location_types": len(space_scores),
        "total_elements": int(df["Element"].nunique()),
        "loc_scores": space_scores,
        "zone_scores": zone_scores,
        "work_order_rows": work_order_rows,
        "insp_scores": dict(zip(per_insp["Inspection #"], per_insp["score"])),
        "insp_meta": insp_meta,
    }


def extract_images(file_path: str, raw_df: pd.DataFrame) -> dict:
    ns = {
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    result = defaultdict(list)
    with zipfile.ZipFile(file_path) as z:
        names = z.namelist()
        drawing = next((n for n in names if n == "xl/drawings/drawing1.xml"), None)
        rels_path = next((n for n in names if n == "xl/drawings/_rels/drawing1.xml.rels"), None)
        if not drawing or not rels_path:
            return dict(result)
        rels_root = ET.fromstring(z.read(rels_path))
        rid_to_target = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root}
        drawing_root = ET.fromstring(z.read(drawing))
        for anchor in drawing_root.findall("xdr:twoCellAnchor", ns):
            row_el = anchor.find("xdr:from/xdr:row", ns)
            blip = anchor.find(".//xdr:blipFill/a:blip", ns)
            if row_el is None or blip is None:
                continue
            raw_idx = int(row_el.text) - 1  # xdr:row=0 is header; data starts at 1
            rid = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if rid is None or raw_idx < 0 or raw_idx >= len(raw_df):
                continue
            insp_id = str(raw_df.iloc[raw_idx].get("Inspection #", ""))
            if not insp_id:
                continue
            img_rel = rid_to_target.get(rid, "")
            img_path = "xl/media/" + Path(img_rel).name
            if img_path not in names:
                continue
            result[insp_id].append(z.read(img_path))
    return dict(result)


# ---------------------------------------------------------------------------
# Chart rendering  (matplotlib → PNG files in tmp dir)
# ---------------------------------------------------------------------------

def render_bar_chart_png(scores: pd.Series, title: str, out_path: Path) -> None:
    labels = [str(l) for l in scores.index]
    values = scores.values.tolist()
    colors = [hex_to_rgb(score_color(v)) for v in values]

    fig_h = max(2.0, min(len(labels) * 0.32, 3.5))
    fig, ax = plt.subplots(figsize=(6.9, fig_h))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    bars = ax.barh(labels, values, color=colors, height=0.55, zorder=2)

    ax.set_xlim(0, 105)
    ax.xaxis.set_visible(False)
    ax.spines[:].set_visible(False)
    ax.tick_params(left=False)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=9, color="#222222")
    ax.yaxis.set_tick_params(length=0)
    ax.invert_yaxis()
    ax.set_title(title, fontsize=10, fontweight="bold",
                 color=NAVY, loc="left", pad=8)

    # Score labels at end of each bar
    for i, (bar, val) in enumerate(zip(bars, values)):
        ax.text(val + 0.8, bar.get_y() + bar.get_height() / 2,
                f"{val:.2f}%", va="center", ha="left", fontsize=8,
                color=NAVY, fontweight="bold")

    # Light gridlines
    ax.set_axisbelow(True)
    ax.xaxis.grid(False)

    plt.tight_layout(pad=0.5)
    fig.savefig(out_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------------------
# LaTeX escaping
# ---------------------------------------------------------------------------

def tex(s: str) -> str:
    """Escape a string for LaTeX."""
    if not isinstance(s, str):
        s = str(s)
    replacements = [
        ("\\", "\\textbackslash{}"),
        ("&",  "\\&"),
        ("%",  "\\%"),
        ("$",  "\\$"),
        ("#",  "\\#"),
        ("_",  "\\_"),
        ("{",  "\\{"),
        ("}",  "\\}"),
        ("~",  "\\textasciitilde{}"),
        ("^",  "\\textasciicircum{}"),
        ("–", "--"),
        ("—", "---"),
        ("’", "'"),
        ("“", "``"),
        ("”", "''"),
    ]
    for old, new in replacements:
        s = s.replace(old, new)
    return s


def color_cmd(hex_color: str) -> str:
    r, g, b = [int(hex_color.lstrip("#")[i:i+2], 16) for i in (0, 2, 4)]
    return f"{r},{g},{b}"


# ---------------------------------------------------------------------------
# LaTeX document builder
# ---------------------------------------------------------------------------

def fwdslash(p) -> str:
    """Convert a path to forward-slash form safe for LaTeX."""
    return str(p).replace("\\", "/")


def build_latex(summary: dict, zone_chart: Path, loc_chart: Path,
                insp_images: dict, tmp_dir: Path) -> str:

    venue       = tex(summary["venue"])
    date_str    = tex(summary["date_str"])
    overall     = summary["overall_score"]
    logo_path   = COMPANY.get("logo_path", "")
    company_name = tex(COMPANY.get("name", ""))

    thead = (
        r"\rowcolor{theadrow}"
        r"\textcolor{white}{\textbf{Zone}} &"
        r"\textcolor{white}{\textbf{Location Type}} &"
        r"\textcolor{white}{\textbf{Element}} &"
        # Merge cols 4+5 into one cell → no inter-column tabcolsep gap to color
        r"\multicolumn{2}{l}{"
        r"\makebox[1.5cm][r]{\textcolor{white}{\textbf{Rating}}}"
        r"\hspace{2\tabcolsep}"
        r"\textcolor{white}{\textbf{Comments}}"
        r"} \\\hline"
        "\n"
        r"\endhead"
    )

    overall_color = color_cmd(score_color(overall))

    # --- Stat cards (4 minipages) ---
    def stat_card(value, label):
        return (
            r"\begin{minipage}[t]{0.23\textwidth}"
            r"\centering"
            r"\colorbox{statbg}{\parbox{\dimexpr\linewidth-2\fboxsep}{"
            r"\hyphenpenalty=10000\exhyphenpenalty=10000"
            r"\parbox[c][2.4cm][c]{\linewidth}{\centering"
            rf"\textbf{{\LARGE\textcolor{{navy}}{{{value}}}}}\\[6pt]"
            rf"\textbf{{\small\textcolor{{slate}}{{\MakeUppercase{{{label}}}}}}}"
            r"}"
            r"}}"
            r"\end{minipage}"
        )

    cards = r"\hfill".join([
        stat_card(summary["total_inspections"],    "Inspections completed"),
        stat_card(summary["total_location_types"], "Location types inspected"),
        stat_card(summary["total_elements"],       "Unique elements inspected"),
        stat_card(len(summary["work_order_rows"]), "Deficiencies flagged"),
    ])

    # --- Work orders table rows ---
    work_rows = summary["work_order_rows"]
    insp_scores = summary["insp_scores"]
    insp_meta   = summary["insp_meta"]

    table_rows = []
    current_insp = None

    for _, row in work_rows.iterrows():
        insp_id  = row.get("Inspection #", "")
        location = str(row.get("Location", ""))
        _, _, space = parse_location(location)
        zone    = str(row.get("Zone", "") or "").strip()
        element = str(row.get("Element", "") or "")
        comment = str(row.get("Comment", "") or "").strip()
        rating_raw = row.get("Rating", None)
        try:
            rating = f"{int(float(rating_raw))}\\%" if pd.notna(rating_raw) else ""
        except (ValueError, TypeError):
            rating = tex(str(rating_raw))

        if insp_id != current_insp:
            current_insp = insp_id
            insp_score = insp_scores.get(insp_id, "")
            score_txt  = f"{insp_score}\\%" if insp_score != "" else ""
            score_rgb  = color_cmd(score_color(float(insp_score))) if insp_score != "" else color_cmd(SLATE)
            meta       = insp_meta.get(insp_id, {})
            comp_dt    = meta.get("completion")
            date_txt   = comp_dt.strftime("%m/%d/%Y").lstrip("0").replace("/0", "/") if pd.notna(comp_dt) else ""
            inspector  = tex(meta.get("completed_by", ""))

            meta_parts = [rf"\textbf{{\textcolor{{navy}}{{#{tex(str(insp_id))}}}}}", tex(location)]
            if date_txt:
                meta_parts.append(tex(date_txt))
            if inspector:
                meta_parts.append(inspector)
            meta_line  = r"\ \textbullet\ ".join(meta_parts)
            score_cell = rf"\textbf{{\textcolor[RGB]{{{score_rgb}}}{{{score_txt}}}}}"

            table_rows.append(
                rf"\rowcolor{{inspbg}}\multicolumn{{4}}{{l}}{{\small {meta_line}}}"
                rf" & \multicolumn{{1}}{{r}}{{\small {score_cell}}} \\"
                r"\noalign{\hrule\penalty10000}"
            )

            # Photo strip if images exist
            imgs = insp_images.get(str(insp_id), [])
            if imgs:
                img_cells = []
                for i, img_bytes in enumerate(imgs[:6]):  # cap at 6 per inspection
                    img_file = tmp_dir / f"img_{insp_id}_{i}.jpg"
                    img_file.write_bytes(img_bytes)
                    img_cells.append(rf"\includegraphics[height=2cm,keepaspectratio]{{{fwdslash(img_file)}}}")
                photos_tex = r"\quad ".join(img_cells)
                table_rows.append(
                    rf"\rowcolor{{photobg}}\multicolumn{{5}}{{l}}{{{photos_tex}}} \\"
                    r"\noalign{\hrule\penalty10000}"
                )

        table_rows.append(
            rf"{tex(zone)} & {tex(space)} & {tex(element)} & {rating} & {tex(comment)} \\\hline"
        )

    table_body = "\n".join(table_rows)

    # --- Logo ---
    if logo_path and Path(logo_path).exists():
        logo_tex = rf"\includegraphics[height=1cm,keepaspectratio]{{{fwdslash(logo_path)}}}"
    else:
        logo_tex = rf"\textbf{{\textcolor{{navy}}{{{company_name}}}}}"

    # --- Full document ---
    return rf"""
\documentclass[10pt,letterpaper]{{article}}

% ---------- packages ----------
\usepackage{{fontspec}}
\usepackage{{geometry}}
\usepackage{{xcolor}}
\usepackage{{graphicx}}
\usepackage{{xltabular}}
\usepackage{{booktabs}}
\usepackage{{array}}
\usepackage{{colortbl}}
\usepackage{{fancyhdr}}
\usepackage{{textcase}}
\usepackage{{microtype}}
\usepackage{{parskip}}
\usepackage{{multirow}}

% ---------- font ----------
\setmainfont{{Segoe UI}}[
  BoldFont      = Segoe UI Bold,
  ItalicFont    = Segoe UI Italic,
  Ligatures     = TeX
]

% ---------- page layout ----------
\geometry{{letterpaper, top=2.5cm, bottom=2cm, left=2cm, right=2cm,
           headheight=1.4cm, headsep=0.4cm}}

% ---------- brand colours ----------
\definecolor{{navy}}{{RGB}}{{37,64,143}}
\definecolor{{sky}}{{RGB}}{{0,176,240}}
\definecolor{{statbg}}{{RGB}}{{240,243,246}}
\definecolor{{slate}}{{RGB}}{{77,91,130}}
\definecolor{{inspbg}}{{RGB}}{{221,227,240}}
\definecolor{{photobg}}{{RGB}}{{247,248,252}}
\definecolor{{theadrow}}{{RGB}}{{37,64,143}}
\definecolor{{rowodd}}{{RGB}}{{255,255,255}}
\definecolor{{roweven}}{{RGB}}{{240,243,246}}

% ---------- running header (every page) ----------
\pagestyle{{fancy}}
\fancyhf{{}}
\renewcommand{{\headrulewidth}}{{0.4pt}}
\fancyhead[L]{{{logo_tex}}}
\fancyhead[C]{{\small\textcolor{{navy}}{{\textbf{{{venue}}}}}}}
\fancyhead[R]{{\small\textcolor{{slate}}{{{date_str}}}}}
\fancyfoot[C]{{\small\thepage}}

% ---------- table helpers ----------
\newcolumntype{{L}}[1]{{>{{\raggedright\arraybackslash}}p{{#1}}}}
\newcolumntype{{R}}[1]{{>{{\raggedleft\arraybackslash}}p{{#1}}}}
\setlength{{\tabcolsep}}{{4pt}}
\renewcommand{{\arraystretch}}{{1.3}}
\setlength{{\arrayrulewidth}}{{0.4pt}}

\begin{{document}}

% ══════════════════════════════════════════
%  COVER / SUMMARY SECTION
% ══════════════════════════════════════════
\thispagestyle{{fancy}}

\vspace*{{0.4cm}}
% Score hero + title row
\noindent
\begin{{minipage}}[t]{{0.22\textwidth}}
  \centering
  \colorbox{{navy}}{{\parbox{{\dimexpr\linewidth-2\fboxsep}}{{
    \vspace{{6pt}}
    \centering
    \textbf{{\small\textcolor{{white}}{{\MakeUppercase{{Overall Score}}}}}}\\[4pt]
    \textbf{{\Huge\textcolor{{white}}{{{overall}\%}}}}
    \vspace{{6pt}}
  }}}}
\end{{minipage}}%
\hfill
\begin{{minipage}}[t]{{0.74\textwidth}}
  \vspace{{4pt}}
  {{\large\textbf{{\textcolor{{navy}}{{Facility Inspection Report}}}}}}\\[2pt]
  {{\small\textcolor{{slate}}{{{venue}}}}}\\[1pt]
  {{\small\textcolor{{slate}}{{{date_str}}}}}
\end{{minipage}}

\vspace{{0.6cm}}

% Stat cards
\noindent
{cards}

\vspace{{0.8cm}}

% Charts
\noindent\includegraphics[width=\textwidth,height=0.28\textheight,keepaspectratio]{{{fwdslash(zone_chart)}}}

\vspace{{0.3cm}}
\noindent\includegraphics[width=\textwidth,height=0.28\textheight,keepaspectratio]{{{fwdslash(loc_chart)}}}

\vspace{{0.8cm}}

\newpage

% ══════════════════════════════════════════
%  DEFICIENCY WORK ORDERS TABLE
% ══════════════════════════════════════════
{{\large\textbf{{\textcolor{{navy}}{{Deficiency Work Orders}}}}}}
\vspace{{0.3cm}}

\begin{{xltabular}}{{\textwidth}}{{L{{2.5cm}} L{{2.5cm}} L{{3cm}} R{{1.5cm}} X}}
{thead}
{table_body}
\end{{xltabular}}

\end{{document}}
"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: generate_report_latex.py <input.xlsx> [output.pdf]")
        sys.exit(1)

    input_file  = sys.argv[1]
    output_pdf  = sys.argv[2] if len(sys.argv) > 2 else input_file.replace(".xlsx", "_report.pdf")
    xelatex     = find_xelatex()

    print(f"[1] Reading {input_file} ...")
    df, filters, raw_df = load_inspections(input_file)
    print(f"    {len(df)} complete rows across {df['Inspection #'].nunique()} inspections.")

    print("[2] Building summary ...")
    summary = build_summary(df, filters)
    insp_images = extract_images(input_file, raw_df)
    print(f"    Overall score : {summary['overall_score']}%")
    print(f"    Work orders   : {len(summary['work_order_rows'])}")
    print(f"    Images        : {sum(len(v) for v in insp_images.values())}")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)

        print("[3] Rendering charts ...")
        zone_chart = tmp_dir / "zone_chart.png"
        loc_chart  = tmp_dir / "loc_chart.png"
        render_bar_chart_png(summary["zone_scores"], "Performance Scores by Zone", zone_chart)
        render_bar_chart_png(summary["loc_scores"],  "Performance Scores by Location Type", loc_chart)

        print("[4] Building LaTeX document ...")
        latex_src = build_latex(summary, zone_chart, loc_chart, insp_images, tmp_dir)

        tex_file = tmp_dir / "report.tex"
        tex_file.write_text(latex_src, encoding="utf-8")
        # Debug: copy tex to working dir for inspection
        shutil.copy(tex_file, Path(input_file).with_suffix(".tex"))

        print("[5] Compiling PDF (pass 1) ...")
        result = subprocess.run(
            [xelatex, "-interaction=nonstopmode", "-output-directory", str(tmp_dir), str(tex_file)],
            capture_output=True, text=True
        )
        pdf_tmp = tmp_dir / "report.pdf"
        if not pdf_tmp.exists():
            # Fatal — show last 30 lines of log
            log_file = tmp_dir / "report.log"
            if log_file.exists():
                lines = log_file.read_text(encoding="utf-8", errors="replace").splitlines()
                print("\n".join(lines[-30:]))
            sys.exit(1)

        # Second pass for longtable page references
        print("[5] Compiling PDF (pass 2) ...")
        subprocess.run(
            [xelatex, "-interaction=nonstopmode", "-output-directory", str(tmp_dir), str(tex_file)],
            capture_output=True, text=True
        )

        shutil.copy(pdf_tmp, output_pdf)

    print(f"[Done] PDF saved to: {output_pdf}")


if __name__ == "__main__":
    main()
