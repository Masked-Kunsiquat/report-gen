"""
generate_report.py
Reads an inspection Excel export and produces a standalone HTML report
mimicking the Arthur Jackson Company facility quality control format.
No image handling — MVP v2.

Usage:
    uv run --with openpyxl --with pandas generate_report.py <input.xlsx> [output.html]
"""

import sys
import re
import html
import json
import base64
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

import pandas as pd


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_inspections(file_path: str) -> tuple[pd.DataFrame, dict]:
    filters_df = pd.read_excel(file_path, sheet_name="Filters", header=None, engine="openpyxl")
    filters = dict(zip(filters_df[0].str.strip(), filters_df[1]))

    df = pd.read_excel(file_path, sheet_name="Raw Data", engine="openpyxl")
    complete = df[df["Status"].fillna("").str.strip().str.lower() == "complete"].copy()
    complete.reset_index(drop=True, inplace=True)
    return complete, filters, df


def extract_images(file_path: str, df: pd.DataFrame) -> dict[str, list[str]]:
    """
    Return {inspection_id: [base64_data_uri, ...]} by reading images directly
    from the xlsx zip — no Excel or xlwings required.

    Images are anchored to spreadsheet rows via the drawing XML. We map each
    row back to its Inspection # using the raw dataframe (row offset accounts
    for the header row).
    """
    ns = {
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    result: dict[str, list[str]] = defaultdict(list)

    with zipfile.ZipFile(file_path) as z:
        names = z.namelist()
        drawing_path = next((n for n in names if n == "xl/drawings/drawing1.xml"), None)
        rels_path = next((n for n in names if n == "xl/drawings/_rels/drawing1.xml.rels"), None)
        if not drawing_path or not rels_path:
            return result

        rels_root = ET.fromstring(z.read(rels_path))
        rid_to_target = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root}

        drawing_root = ET.fromstring(z.read(drawing_path))
        for anchor in drawing_root.findall("xdr:twoCellAnchor", ns):
            row_el = anchor.find("xdr:from/xdr:row", ns)
            blip = anchor.find(".//xdr:blipFill/a:blip", ns)
            if row_el is None or blip is None:
                continue

            # xdr:row=0 is the header row; subtract 1 to get pandas index
            raw_idx = int(row_el.text) - 1
            rid = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if rid is None or raw_idx < 0 or raw_idx >= len(df):
                continue

            insp_id = str(df.iloc[raw_idx].get("Inspection #", ""))
            if not insp_id:
                continue

            # Resolve relative path: "../media/imageN.jpeg" -> "xl/media/imageN.jpeg"
            img_path = "xl/drawings/" + rid_to_target[rid]
            img_path = str(Path(img_path).resolve().relative_to(Path(".").resolve()))
            # Normalise to forward slashes and strip leading separator
            img_path = img_path.replace("\\", "/").lstrip("/")

            if img_path not in names:
                # Fallback: just strip the leading ../
                img_path = rid_to_target[rid].lstrip("../")
                img_path = "xl/media/" + Path(img_path).name

            if img_path not in names:
                continue

            ext = Path(img_path).suffix.lstrip(".").lower()
            mime = {"jpeg": "image/jpeg", "jpg": "image/jpeg", "png": "image/png"}.get(ext, "image/jpeg")
            data = base64.b64encode(z.read(img_path)).decode()
            result[insp_id].append(f"data:{mime};base64,{data}")

    return dict(result)


def parse_location(location: str) -> tuple[str, str, str]:
    """
    Extract the trailing space/room type from a location string.

    Handles multiple account formats:
      '1920-Dynamic Engineering - Office'             → space='Office'
      '15 - Common Areas - Men's RR'                 → space='Men's RR'
      'L1- DEC_Common Areas_ Pantry-Kitchen-Cafeteria' → space='Pantry-Kitchen-Cafeteria'

    Only the third return value (space) is used downstream; floor and area
    are kept for signature compatibility but are not populated.
    """
    location = str(location).strip()
    # Try separators in priority order: ' - ' is most common, '_ ' handles
    # accounts that use underscores (e.g. Jefferson format).
    for sep in (" - ", "_ "):
        if sep in location:
            _, space = location.rsplit(sep, 1)
            return "", "", space.strip()
    return "", "", location


def drop_image_stubs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove platform image-stub rows — comments that are only a parenthetical
    like '(Market)' or '(East)', where the same parenthetical already appears
    at the end of another comment in the same inspection + element group.
    Both conditions must be true to drop; either alone is not enough.
    """
    stub_pattern = re.compile(r"^\s*\(([^)]+)\)\s*$")

    def is_stub(comment: str, group_comments: list[str]) -> bool:
        m = stub_pattern.match(comment)
        if not m:
            return False
        tag = m.group(1)
        # Check if any sibling comment ends with the same parenthetical
        return any(
            c != comment and re.search(r"\(" + re.escape(tag) + r"\)\s*$", c)
            for c in group_comments
        )

    drop_idx = set()
    for _, group in df.groupby(["Inspection #", "Element"], sort=False):
        comments = group[comment_col(df)].fillna("").tolist()
        for idx, comment in zip(group.index, comments):
            if is_stub(comment, comments):
                drop_idx.add(idx)

    return df[~df.index.isin(drop_idx)].reset_index(drop=True)


def comment_col(df: pd.DataFrame) -> str:
    """Return whichever comment column name exists in this export."""
    return "Comments" if "Comments" in df.columns else "Comment"


def build_summary(df: pd.DataFrame, filters: dict) -> dict:
    """Aggregate inspection data into a summary dict for the report."""
    # Inspection # is the SoT for one inspection instance.
    # Each inspection has exactly one Score In %, Location, etc. — take the first
    # row per Inspection # to get a clean one-row-per-inspection table.
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
    date_to = pd.to_datetime(filters.get("To Date"), errors="coerce")
    if pd.notna(date_from) and pd.notna(date_to):
        if date_from.date() == date_to.date():
            date_str = date_from.strftime("%B %d, %Y")
        else:
            date_str = f"{date_from.strftime('%B %d, %Y')} – {date_to.strftime('%B %d, %Y')}"
    else:
        date_str = ""

    # Space type: the part after " - " in Location (e.g. "Men's RR", "Break Room")
    per_insp["_space"] = per_insp["location"].apply(lambda loc: parse_location(str(loc))[2])

    # Average the platform score across all inspections of each space type
    space_scores = (
        per_insp.groupby("_space")["score"]
        .mean()
        .round(2)
        .sort_values(ascending=True)
    )

    # Zone scores: pull Zone from element rows (consistent per inspection), join to per_insp scores
    insp_zone = (
        df.groupby("Inspection #")["Zone"]
        .first()
        .reset_index()
        .rename(columns={"Zone": "zone"})
    )
    per_insp_z = per_insp.merge(insp_zone, on="Inspection #", how="left")
    zone_scores = (
        per_insp_z.groupby("zone")["score"]
        .mean()
        .round(2)
        .sort_values(ascending=True)
    )

    # Work orders: rows that have a Comment
    _cc = comment_col(df)
    work_order_rows = df[df[_cc].notna() & (df[_cc].astype(str).str.strip() != "")].copy()
    work_order_rows = work_order_rows.rename(columns={_cc: "Comment"})
    work_order_rows = drop_image_stubs(work_order_rows)

    return {
        "venue": venue,
        "building": building,
        "corporation": corporation,
        "date_str": date_str,
        "overall_score": overall_score,
        "total_inspections": len(per_insp),
        "total_location_types": len(space_scores),
        "total_elements": int(df["Element"].nunique()),
        "loc_scores": space_scores,
        "zone_scores": zone_scores,
        "work_order_rows": work_order_rows,
        "insp_scores": dict(zip(per_insp["Inspection #"], per_insp["score"])),
        "insp_meta": {
            row["Inspection #"]: {
                "completion": pd.to_datetime(row["completion"], errors="coerce"),
                "completed_by": str(row["completed_by"]).strip() if pd.notna(row["completed_by"]) else "",
            }
            for _, row in per_insp.iterrows()
        },
    }


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

def load_company(script_path: Path) -> dict:
    root = script_path.parent
    config_path = root / "company.json"
    company = {"name": "", "address": "", "phone": "", "logo_tag": ""}

    if config_path.exists():
        with open(config_path, encoding="utf-8") as f:
            company.update(json.load(f))

    mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "svg": "image/svg+xml"}
    for ext, mime in mime_map.items():
        logo_path = root / f"logo.{ext}"
        if logo_path.exists():
            data = base64.b64encode(logo_path.read_bytes()).decode()
            company["logo_tag"] = f'<img src="data:{mime};base64,{data}" class="company-logo" alt="Company logo">'
            break

    return company

COMPANY = load_company(Path(__file__))

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;700&display=swap');
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Inter', Arial, sans-serif; font-weight: 300; font-size: 13px; color: #222; background: #fff; }
.page { width: 850px; margin: 0 auto; padding: 40px 50px; }

/* Top bar */
.header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #00B0F0; padding-bottom: 12px; margin-bottom: 18px; }
.company-name { font-size: 17px; font-weight: 700; color: #25408F; }
.company-sub { font-size: 11px; color: #4D5B82; margin-top: 3px; }
.header-right { display: flex; flex-direction: column; align-items: flex-end; gap: 8px; }
.report-title { font-size: 11px; color: #4D5B82; text-transform: uppercase; letter-spacing: 0.08em; }
.company-logo { max-height: 48px; max-width: 160px; object-fit: contain; }

/* Venue + date */
.venue-row { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 20px; }
.venue-name { font-size: 22px; font-weight: 700; color: #25408F; }
.venue-period { font-size: 14px; font-weight: 700; color: #25408F; }

/* Score hero + stat cards */
.summary-row { display: flex; gap: 12px; margin-bottom: 28px; align-items: stretch; }
.score-hero { background: #25408F; border-radius: 10px; padding: 18px 24px; min-width: 155px; display: flex; flex-direction: column; justify-content: center; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.score-hero-label { font-size: 10px; color: #00B0F0; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 6px; }
.score-hero-value { font-size: 38px; font-weight: 700; color: #fff; line-height: 1; }
.stat-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; flex: 1; }
.stat-card { background: #F0F3F6; border-radius: 8px; padding: 14px 16px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.stat-value { font-size: 24px; font-weight: 700; color: #25408F; margin-bottom: 6px; }
.stat-label { font-size: 10px; color: #4D5B82; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; }

/* Section titles */
.section-title { font-size: 11px; font-weight: 700; color: #4D5B82; border-bottom: 2px solid #00B0F0; padding-bottom: 4px; margin-bottom: 14px; margin-top: 28px; text-transform: uppercase; letter-spacing: 0.08em; }

/* Work orders table */
table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 12px; }
thead tr { background: #25408F; color: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
th { padding: 8px 10px; text-align: left; font-weight: 700; }
td { padding: 7px 10px; border-bottom: 1px solid #e0e4ec; vertical-align: top; }
tr:nth-child(even) td { background: #F0F3F6; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
tr.insp-header td { background: #dde3f0; padding: 6px 10px; border-bottom: 1px solid #b0bdd4; font-weight: 700; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.insp-id { color: #25408F; margin-right: 10px; }
.insp-location { color: #25408F; margin-right: 10px; }
.insp-score { float: right; }
.insp-sep { color: #4D5B82; margin: 0 6px; font-weight: 400; }
.insp-date { color: #4D5B82; font-weight: 400; font-size: 11px; }
.insp-inspector { color: #4D5B82; font-weight: 400; font-size: 11px; }
tr.insp-photos td { background: #f7f8fc; padding: 8px 10px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.insp-photo-strip { display: flex; flex-wrap: wrap; gap: 8px; }
.insp-photo { height: 120px; width: auto; border-radius: 4px; border: 1px solid #dde3f0; object-fit: cover; }
.zone-col   { width: 140px; }
.space-col  { width: 140px; }
.rating-col { width: 60px; text-align: center; }
.element-col { width: 140px; }
.comment-col { }

@media print {
  .page { padding: 20px 30px; padding-top: 70px; }
  .page-break { page-break-before: always; }
  .print-header {
    display: flex;
    position: fixed;
    top: 0; left: 0; right: 0;
    height: 52px;
    background: #fff;
    border-bottom: 2px solid #00B0F0;
    padding: 0 30px;
    align-items: center;
    justify-content: space-between;
    z-index: 100;
  }
  .print-header-logo { max-height: 36px; max-width: 120px; object-fit: contain; }
  .print-header-venue { font-size: 13px; font-weight: 700; color: #25408F; }
  .print-header-date { font-size: 11px; color: #4D5B82; }
  .company-logo { display: none; }
}

@media screen {
  .print-header { display: none; }
}
"""


def score_color(score: float) -> str:
    if score >= 85:
        return "#85CF5F"   # pass
    if score >= 80:
        return "#F0B557"   # warning
    return "#F05773"       # failing


def bar_color(score: float) -> str:
    return score_color(score)


CHART_LABEL_W = 210  # shared across all charts so bars align between them
CHART_BAR_AREA = 430
CHART_SCORE_W = 52
CHART_FONT_PX = 11
CHART_AVG_CHAR_PX = 6.2  # approximate Arial char width at 11px


def truncate_label(text: str, max_px: int) -> str:
    """Truncate label to fit within max_px, appending ellipsis if cut."""
    max_chars = int((max_px - 8) / CHART_AVG_CHAR_PX)
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1] + "…"


def render_bar_chart(scores: pd.Series) -> str:
    """Render an SVG horizontal bar chart — prints correctly unlike CSS divs."""
    row_h = 24
    padding = 6
    label_w = CHART_LABEL_W
    bar_area = CHART_BAR_AREA
    score_w = CHART_SCORE_W
    total_w = label_w + bar_area + score_w + 20
    total_h = len(scores) * row_h + 10

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{total_w}" height="{total_h}" '
        f'style="font-family:Arial,sans-serif;font-size:{CHART_FONT_PX}px;">',
    ]

    for i, (label, score) in enumerate(scores.items()):
        y = i * row_h
        bar_w = int(min(score, 100) / 100 * bar_area)
        color = bar_color(score)
        label_esc = html.escape(truncate_label(str(label), label_w))
        score_txt = f"{score:.2f}%"

        lines.append(
            f'  <text x="4" y="{y + row_h - padding - 3}" '
            f'text-anchor="start" fill="#4D5B82">{label_esc}</text>'
        )
        lines.append(
            f'  <rect x="{label_w}" y="{y + padding}" '
            f'width="{bar_area}" height="{row_h - padding * 2}" rx="3" fill="#e8edf5"/>'
        )
        lines.append(
            f'  <rect x="{label_w}" y="{y + padding}" '
            f'width="{bar_w}" height="{row_h - padding * 2}" rx="3" fill="{color}"/>'
        )
        lines.append(
            f'  <text x="{label_w + bar_area + 6}" y="{y + row_h - padding - 3}" '
            f'fill="#25408F" font-weight="bold">{score_txt}</text>'
        )

    lines.append("</svg>")
    return "\n".join(lines)


def render_work_orders(work_rows: pd.DataFrame, insp_scores: dict, insp_meta: dict, insp_images: dict) -> str:
    if work_rows.empty:
        return "<p><em>No work orders — all areas passed.</em></p>"

    rows_html = []
    current_insp = None

    for _, row in work_rows.iterrows():
        insp_id = row.get("Inspection #", "")
        location = str(row.get("Location", ""))
        _, _, space = parse_location(location)
        zone = html.escape(str(row.get("Zone", "") or "").strip())
        element = html.escape(str(row.get("Element", "") or ""))
        comment = html.escape(str(row.get("Comment", "") or "").strip())
        rating_raw = row.get("Rating", None)
        try:
            rating = f"{int(float(rating_raw))}%" if pd.notna(rating_raw) else ""
        except (ValueError, TypeError):
            rating = html.escape(str(rating_raw))

        if insp_id != current_insp:
            current_insp = insp_id
            insp_score = insp_scores.get(insp_id, "")
            score_txt = f"{insp_score}%" if insp_score != "" else ""
            score_color_val = score_color(float(insp_score)) if insp_score != "" else "#4D5B82"
            meta = insp_meta.get(insp_id, {})
            completion_dt = meta.get("completion")
            date_txt = completion_dt.strftime("%m/%d/%Y").lstrip("0").replace("/0", "/") if pd.notna(completion_dt) else ""
            inspector = html.escape(meta.get("completed_by", ""))
            rows_html.append(f"""
    <tr class="insp-header">
      <td colspan="5">
        <span class="insp-id">#{insp_id}</span>
        <span class="insp-location">{html.escape(location)}</span>
        <span class="insp-score" style="color:{score_color_val}">{score_txt}</span>
        {f'<span class="insp-sep">·</span><span class="insp-date">{date_txt}</span>' if date_txt else ''}
        {f'<span class="insp-sep">·</span><span class="insp-inspector">{inspector}</span>' if inspector else ''}
      </td>
    </tr>""")

            images = insp_images.get(str(insp_id), [])
            if images:
                imgs_html = "".join(f'<img src="{src}" class="insp-photo">' for src in images)
                rows_html.append(f"""
    <tr class="insp-photos">
      <td colspan="5"><div class="insp-photo-strip">{imgs_html}</div></td>
    </tr>""")

        rows_html.append(f"""
    <tr>
      <td class="zone-col">{zone}</td>
      <td class="space-col">{html.escape(space)}</td>
      <td class="element-col">{element}</td>
      <td class="rating-col">{rating}</td>
      <td class="comment-col">{comment}</td>
    </tr>""")

    return f"""
<table>
  <thead>
    <tr>
      <th class="zone-col">Zone</th>
      <th class="space-col">Location Type</th>
      <th class="element-col">Element</th>
      <th class="rating-col">Rating</th>
      <th class="comment-col">Comments</th>
    </tr>
  </thead>
  <tbody>{"".join(rows_html)}
  </tbody>
</table>"""


def render_html(summary: dict) -> str:
    venue = html.escape(summary["venue"])
    date_str = html.escape(summary["date_str"])
    overall = summary["overall_score"]
    loc_scores = summary["loc_scores"]
    work_rows = summary["work_order_rows"]
    total_location_types = summary["total_location_types"]
    total_inspections = summary["total_inspections"]
    total_elements = summary["total_elements"]
    total_deficiencies = len(work_rows)

    bar_chart = render_bar_chart(loc_scores)
    zone_chart = render_bar_chart(summary["zone_scores"])
    work_table = render_work_orders(work_rows, summary["insp_scores"], summary["insp_meta"], summary.get("insp_images", {}))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Facility Inspection Report – {venue}</title>
  <style>{CSS}</style>
</head>
<body>

<!-- Compact running header: only visible in print, repeats on every page -->
<div class="print-header">
  {COMPANY["logo_tag"].replace('class="company-logo"', 'class="print-header-logo"')}
  <span class="print-header-venue">{venue}</span>
  <span class="print-header-date">{date_str}</span>
</div>

<div class="page">

  <!-- Top bar -->
  <div class="header">
    <div>
      <div class="company-name">{html.escape(COMPANY["name"])}</div>
      <div class="company-sub">{html.escape(COMPANY["address"])} &middot; {html.escape(COMPANY["phone"])}</div>
    </div>
    <div class="header-right">
      {COMPANY["logo_tag"]}
      <div class="report-title">Facility Inspection Report</div>
    </div>
  </div>

  <!-- Venue + date -->
  <div class="venue-row">
    <div class="venue-name">{venue}</div>
    <div class="venue-period">{date_str}</div>
  </div>

  <!-- Score hero + stat cards -->
  <div class="summary-row">
    <div class="score-hero">
      <div class="score-hero-label">Overall score</div>
      <div class="score-hero-value">{overall}%</div>
    </div>
    <div class="stat-cards">
      <div class="stat-card">
        <div class="stat-value">{total_inspections}</div>
        <div class="stat-label">Inspections completed</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">{total_location_types}</div>
        <div class="stat-label">Location types inspected</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">{total_elements}</div>
        <div class="stat-label">Unique elements inspected</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">{total_deficiencies}</div>
        <div class="stat-label">Deficiencies flagged</div>
      </div>
    </div>
  </div>

  <!-- Bar chart -->
  <div class="section-title">Performance Scores by Zone</div>
  {zone_chart}

  <div class="section-title">Performance Scores by Location Type</div>
  {bar_chart}

  <!-- Work orders -->
  <div class="section-title page-break">Inspection Work Orders</div>
  {work_table}

</div>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: generate_report.py <input.xlsx> [output.html]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace(".xlsx", "_report.html")

    print(f"[1] Reading {input_file} ...")
    df, filters, raw_df = load_inspections(input_file)
    print(f"    {len(df)} complete rows across {df['Inspection #'].nunique()} inspections.")

    print("[2] Building summary...")
    summary = build_summary(df, filters)
    insp_images = extract_images(input_file, raw_df)
    summary["insp_images"] = insp_images
    print(f"    Images found: {sum(len(v) for v in insp_images.values())} across {len(insp_images)} inspections.")
    print(f"    Overall score: {summary['overall_score']}%")
    print(f"    Locations: {len(summary['loc_scores'])}")
    print(f"    Work orders: {len(summary['work_order_rows'])}")

    print("[3] Rendering HTML...")
    html_out = render_html(summary)

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html_out)

    print(f"[Done] Report saved to: {output_file}")


if __name__ == "__main__":
    main()
