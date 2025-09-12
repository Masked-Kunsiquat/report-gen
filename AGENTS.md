## Project Overview: Audit Parsing Pipeline

This project automates the transformation of audit data exported from Excel into structured JSON and ultimately into a PDF summary report. It is designed to be modular, secure, and extensible.

---

## Current Module: `extract.py`

### Responsibilities
- Load `audit_data.xlsx` and access the `Raw Data` sheet.
- Filter rows where `Status == "Complete"`.
- Group rows by `Inspection #`.
- Normalize inspection IDs (e.g., convert `1432859.0` → `"1432859"`).
- Extract metadata and element-level details.
- Extract embedded images using `xlwings` and `ImageGrab`.
- Match images to the correct inspection row.
- Save images to `attachments/` folder.
- Update `inspection_summary.json` with image paths.
- Print progress and summary report.

### Progress Reporting
- Each step emits structured progress messages to stdout.
- Summary includes:
  - Total inspections parsed
  - Images matched and saved
  - Unmatched images
  - Errors during image processing

---

## Raw Data Sheet Structure

### Primary Grouping
- `Inspection #`: Unique identifier for each inspection.

### Global Metadata
- `Corporation`, `Venue`, `Building`
- `Scheduled Date`, `Creation Date`, `Completion Date`
- `Completed By`, `Status`, `Overall Comment`, `Score in Percent`
- `Alert Type`

### Element-Level Details
- `Zone`, `Location`, `Element`
- `Score Factor`, `Element Weight In %`, `Rating`, `Element Score in %`
- `Comments`, `Attachment` (embedded image)
- Multiple rows per element are possible.

---

## Planned Modules

### Module 3: PDF Generation
- Read enriched `inspection_summary.json`
- Generate styled PDF summary per inspection
- Embed images and comments
- Apply branding/styling

### Module 4: Orchestrator
- CLI or Python script to run modules in sequence
- Accept filters (e.g., inspector name, alert type)
- Display progress across modules
- Handle configuration and logging

---

## Design Notes

- Modular architecture: Each module handles a distinct task.
- JSON is the central data format for inter-module communication.
- Progress reporting is standardized for orchestrator integration.
- Future-proofing for additional alert types and inspection formats.

---

## Getting Started

### Prerequisites
- Microsoft Excel installed (Windows/macOS).
- xlwings has access to Excel (not supported headless/on Linux without Excel).
- PIL.ImageGrab requires a GUI clipboard.

1. Place `audit_data.xlsx` in the working directory.
2. Run `extract.py` to generate `inspection_summary.json` and extract images.
3. Proceed to PDF generation once JSON is verified.

