import pandas as pd
import json
import os
import xlwings as xw
from PIL import ImageGrab
from collections import defaultdict
import numpy as np
from datetime import datetime

def sanitize_for_json(obj):
    """Convert non-JSON-serializable objects to JSON-safe primitives."""
    if obj is None:
        return None
    
    # Handle collections recursively first (before pd.isna check)
    if isinstance(obj, dict):
        return {key: sanitize_for_json(value) for key, value in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize_for_json(item) for item in obj]
    if isinstance(obj, np.ndarray):
        return [sanitize_for_json(item) for item in obj.tolist()]
    
    # Handle pandas NaT and numpy NaN (only for scalar-like objects)
    if not isinstance(obj, (dict, list, tuple, np.ndarray)):
        try:
            if pd.isna(obj):
                return None
        except (TypeError, ValueError):
            # pd.isna may fail on some object types, continue with other checks
            pass
    
    # Handle numpy scalars
    if isinstance(obj, (np.integer, np.int64, np.int32)):
        return int(obj)
    if isinstance(obj, (np.floating, np.float64, np.float32)):
        return float(obj)
    if isinstance(obj, np.bool_):
        return bool(obj)
    
    # Handle datetime objects
    if isinstance(obj, (pd.Timestamp, np.datetime64)):
        try:
            if pd.isna(obj):
                return None
        except (TypeError, ValueError):
            pass
        return pd.Timestamp(obj).isoformat()
    if isinstance(obj, datetime):
        return obj.isoformat()
    
    # Return primitive types as-is
    return obj

def normalize_inspection_id(value):
    """Convert float-like inspection IDs to clean strings."""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value)

def load_flat_inspection_data(file_path):
    """Load inspection data in flat table format matching original PDF structure."""
    df = pd.read_excel(file_path, sheet_name='Raw Data', engine='openpyxl')
    df_complete = df[df['Status'].fillna('').str.strip().str.lower() == 'complete']
    
    # Helper functions for data cleaning
    def get_safe_value(row, key, default=""):
        value = row.get(key, default)
        if pd.isna(value):
            return None
        # Handle datetime objects
        if isinstance(value, (pd.Timestamp, np.datetime64)):
            return pd.Timestamp(value).isoformat()
        # Handle numpy/pandas scalars
        if hasattr(value, 'item'):
            return value.item()
        return value
    
    def get_safe_string(row, key, default=""):
        value = get_safe_value(row, key, default)
        if value is None:
            return ""
        str_value = str(value).strip()
        return str_value if str_value else ""
    
    def get_safe_number(row, key, default=None):
        value = get_safe_value(row, key, default)
        if value is None:
            return ""
        try:
            # Try to parse as float first, then int if it's a whole number
            float_val = float(value)
            if float_val.is_integer():
                return int(float_val)
            return float_val
        except (ValueError, TypeError):
            return ""

    # Create flat table structure - one row per element
    table_rows = []
    row_map = {}  # Maps Excel row to table row index for image assignment
    
    for idx, row in df_complete.iterrows():
        # Create flat row matching original PDF columns exactly
        table_row = {
            "inspection_number": normalize_inspection_id(row.get("Inspection #", "")),
            "corporation": get_safe_string(row, "Corporation"),
            "venue": get_safe_string(row, "Venue"), 
            "building": get_safe_string(row, "Building"),
            "zone": get_safe_string(row, "Zone"),
            "location": get_safe_string(row, "Location"),
            "status": get_safe_string(row, "Status"),
            "scheduled_date": get_safe_string(row, "Scheduled Date"),
            "creation_date": get_safe_string(row, "Creation Date"),
            "completion_date": get_safe_string(row, "Completion Date"),
            "completed_by": get_safe_string(row, "Completed By"),
            "overall_comment": get_safe_string(row, "Overall Comment"),
            "score_percent": get_safe_number(row, "Score In %"),
            "element": get_safe_string(row, "Element"),
            "score_factor": get_safe_string(row, "Score Factor"),
            "element_weight_percent": get_safe_number(row, "Element Weight In %"),
            "rating": get_safe_string(row, "Rating"),
            "element_score_percent": get_safe_number(row, "Element Score In %"),
            "comment": get_safe_string(row, "Comment"),
            "attachment": ""  # Will be populated by image extraction
        }
        
        table_rows.append(table_row)
        # Map Excel row number (1-based) to table row index
        excel_row = idx + 2  # Convert 0-based pandas index to 1-based Excel row
        row_map[excel_row] = len(table_rows) - 1
    
    return table_rows, row_map

def extract_images_and_update_table(excel_path, table_rows, row_map):
    # Create private Excel app to avoid process leaks
    app = xw.App(visible=False, add_book=False)
    wb = None
    try:
        wb = app.books.open(excel_path)
        sheet = wb.sheets['Raw Data']
        os.makedirs('attachments', exist_ok=True)

        image_index_map = defaultdict(int)
        unmatched_images = []
        matched_images = 0
        errors = 0

        # Collect and sort pictures for deterministic assignment
        pictures_data = []
        for pic in sheet.pictures:
            try:
                # Use TopLeftCell.Row for accurate row detection
                row = pic.api.TopLeftCell.Row
                col = pic.api.TopLeftCell.Column
                inspection_id_raw = sheet.range(f'A{row}').value
                norm_id = normalize_inspection_id(inspection_id_raw)
                pictures_data.append((row, col, pic, norm_id))
            except Exception as e:  # noqa: BLE001 — intentional to keep pipeline resilient
                import traceback
                print(f"[Error] Failed to get position for picture: {e}")
                traceback.print_exc()
                errors += 1

        # Sort by row, then column for deterministic assignment
        pictures_data.sort(key=lambda x: (x[0], x[1]))

        for row, col, pic, norm_id in pictures_data:
            try:
                # Retry image copy
                img = None
                for attempt in range(3):
                    pic.api.Copy()
                    img = ImageGrab.grabclipboard()
                    if img:
                        break

                if img is None:
                    print(f"[Warning] No image found in clipboard for row {row}.")
                    unmatched_images.append(norm_id)
                    continue

                image_index_map[norm_id] += 1
                image_filename = f"{norm_id}_img_{image_index_map[norm_id]}.png"
                image_path = os.path.join('attachments', image_filename)
                img.save(image_path)

                # Match to specific table row using exact Excel row
                matched = False
                if row in row_map:
                    table_row_index = row_map[row]
                    if table_row_index < len(table_rows):
                        table_rows[table_row_index]['attachment'] = image_path.replace("\\", "/")
                        matched = True
                        matched_images += 1
                        print(f"[Image Saved] {image_filename} for row {table_row_index} (Excel row {row})")

                if not matched:
                    unmatched_images.append(norm_id)
                    print(f"[Warning] Could not match image at Excel row {row} to table row")

            except Exception as e:  # noqa: BLE001 — intentional to keep pipeline resilient
                import traceback
                print(f"[Error] Failed to process image: {e}")
                traceback.print_exc()
                errors += 1

    finally:
        # Always close workbook and quit app to prevent process leaks
        if wb:
            wb.close()
        app.quit()

    return matched_images, unmatched_images, errors

if __name__ == "__main__":
    input_file = "audit_data.xlsx"
    output_file = "inspection_table.json"  # Changed to reflect flat table structure

    try:
        print("[Step 1] Loading inspection data in flat table format...")
        table_rows, row_map = load_flat_inspection_data(input_file)
        print(f"[Step 2] Found {len(table_rows)} inspection elements.")
        print("[Step 3] Extracting embedded images using xlwings...")

        matched_images, unmatched_images, errors = extract_images_and_update_table(
            input_file, table_rows, row_map
        )

        print("[Step 4] Image extraction complete. Sanitizing data for JSON...")
        # Sanitize all table data for JSON serialization AFTER image processing
        table_rows = sanitize_for_json(table_rows)
        
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(table_rows, f, indent=2, ensure_ascii=False)

        print(f"[Done] Inspection table saved to '{output_file}'.")
        print("\n--- Summary Report ---")
        print(f"Total table rows parsed: {len(table_rows)}")
        print(f"Images successfully matched and saved: {matched_images}")
        print(f"Unmatched images: {len(unmatched_images)}")
        if unmatched_images:
            print(f"Unmatched Inspection IDs: {set(unmatched_images)}")
        print(f"Errors during image processing: {errors}")

    except FileNotFoundError:
        print(f"[Error] File '{input_file}' not found.")
    except Exception as e:  # noqa: BLE001 — intentional to keep CLI resilient
        import traceback
        print(f"[Error] An unexpected error occurred: {e}")
        traceback.print_exc()
