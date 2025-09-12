# LaTeX Report Generation

This module converts JSON inspection data to professional LaTeX/PDF reports.

## Overview

The report generation system consists of:

1. **LaTeX Template** (`templates/inspection_template.tex`) - Professional template with:
   - Clean title page with inspection overview
   - Color-coded scores (red <50%, orange 50-74%, green ≥75%)
   - Embedded images from attachments folder
   - Conditional sections (only show if data exists)
   - Multi-page element details with summary table

2. **Python Generator** (`generate_reports.py`) - Converts JSON to LaTeX:
   - Individual reports per inspection
   - Combined report with all inspections
   - Filtering by inspector, venue, alert type
   - Automatic PDF compilation (if pdflatex available)

## Quick Start

### 1. Generate Individual Reports

```bash
# Generate .tex files for each inspection
python generate_reports.py --single

# Generate and compile to PDF
python generate_reports.py --single --compile
```

### 2. Generate Combined Report

```bash
# All inspections in one document
python generate_reports.py --combined --compile
```

### 3. Filter Reports

```bash
# Only inspections by specific inspector
python generate_reports.py --single --filter inspector="John Smith"

# Only critical alerts
python generate_reports.py --combined --filter alert_type="Critical"

# Multiple filters
python generate_reports.py --single --filter venue="Building A" --filter inspector="John"
```

## File Structure

```
project/
├── extract.py                          # JSON extraction (your existing script)
├── generate_reports.py                 # LaTeX generation script
├── templates/
│   └── inspection_template.tex         # LaTeX template
├── reports/                            # Generated output
│   ├── inspection_12345.tex
│   ├── inspection_12345.pdf
│   └── combined_inspection_report.pdf
├── attachments/                        # Images from extract.py
│   ├── 12345_img_1.png
│   └── 12345_img_2.png
└── inspection_summary.json            # Input data
```

## Template Features

### Professional Layout
- Clean typography with consistent spacing
- Corporate color scheme (blue headers, gray text)
- Responsive table layouts
- Page headers/footers with metadata

### Smart Content Handling
- **Conditional Sections**: Only displays fields with data
- **Score Coloring**: Visual indicators for performance levels
- **Image Embedding**: Automatic image inclusion from attachments/
- **Text Escaping**: Handles special characters safely

### Template Variables

#### Inspection Metadata
- `{{INSPECTION_ID}}` - Unique identifier
- `{{CORPORATION}}`, `{{VENUE}}`, `{{BUILDING}}` - Location info
- `{{COMPLETED_BY}}` - Inspector name
- `{{COMPLETION_DATE}}` - When inspection finished
- `{{SCORE_PERCENT}}` - Overall score with color coding
- `{{ALERT_TYPE}}` - Critical alerts highlighted

#### Element Details (loops through `{{#each ELEMENTS}}`)
- `{{zone}}`, `{{location}}`, `{{element}}` - Element identification
- `{{score_factor}}`, `{{element_weight_percent}}` - Scoring details
- `{{rating}}`, `{{element_score_percent}}` - Assessment results
- `{{comments}}` - Inspector notes
- `{{attachment}}` - Image path (auto-embedded)

## Customization

### Modify Template
Edit `templates/inspection_template.tex` to:
- Change colors: Modify `\\definecolor{headerblue}{RGB}{41, 98, 167}`
- Add company logo: Include `\\includegraphics{logo.png}` in title page
- Adjust layout: Modify geometry settings

### Custom Filters
Add new filters in `generate_reports.py`:

```python
def filter_inspections(inspections, filters):
    # Add custom filter logic
    if 'score_below' in filters:
        threshold = float(filters['score_below'])
        filtered = [insp for insp in filtered 
                   if float(insp.get('score_percent', 0)) < threshold]
```

## Dependencies

### Python Packages
```bash
pip install pathlib  # (built-in Python 3.4+)
```

### LaTeX Installation
For PDF compilation:

**Windows**: Install MiKTeX or TeX Live
**macOS**: Install MacTeX
**Linux**: `sudo apt-get install texlive-full`

Required LaTeX packages (usually included):
- geometry, graphicx, booktabs, longtable
- xcolor, fancyhdr, hyperref

## Integration with extract.py

The system works seamlessly with your existing `extract.py`:

```bash
# Full pipeline
python extract.py                    # Generate JSON + images
python generate_reports.py --compile # Generate PDFs

# Or combined
python extract.py && python generate_reports.py --single --compile
```

## Troubleshooting

### Common Issues

**"Template not found"**
- Ensure `templates/` directory exists with `inspection_template.tex`

**"pdflatex not found"**
- Install LaTeX distribution or use `--single` without `--compile`
- LaTeX files can be compiled manually: `pdflatex inspection_12345.tex`

**"Image not found"**
- Verify `attachments/` directory contains referenced images
- Check that image paths in JSON use forward slashes

**Compilation errors**
- Check LaTeX log files (`.log`) in reports directory
- Verify special characters are properly escaped

### Performance Notes

- Large images: LaTeX will auto-resize, but pre-resize for faster compilation
- Many inspections: Use `--combined` for better performance than multiple PDFs
- Complex filtering: Consider preprocessing JSON instead of runtime filtering

## Example Output

The generated reports include:

1. **Title Page**: Inspection overview with key metrics
2. **Details Pages**: Element-by-element analysis with images
3. **Summary Table**: Quick reference for multi-element inspections

Each report is self-contained and professional enough for client delivery or regulatory submission.