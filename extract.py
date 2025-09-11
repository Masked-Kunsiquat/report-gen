import pandas as pd
import json
import os
import xlwings as xw
from PIL import ImageGrab
from collections import defaultdict

def normalize_inspection_id(value):
    """Convert float-like inspection IDs to clean strings."""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value)

def load_and_group_inspections(file_path):
    df = pd.read_excel(file_path, sheet_name='Raw Data', engine='openpyxl')
    df_complete = df[df['Status'].str.lower() == 'complete']
    grouped = df_complete.groupby('Inspection #')

    inspections = []
    row_map = defaultdict(list)  # Maps inspection_id to row indices

    for inspection_id, group in grouped:
        norm_id = normalize_inspection_id(inspection_id)
        first_row = group.iloc[0]

        inspection_data = {
            "inspection_id": norm_id,
            "corporation": first_row.get("Corporation", ""),
            "venue": first_row.get("Venue", ""),
            "building": first_row.get("Building", ""),
            "scheduled_date": str(first_row.get("Scheduled Date", "")),
            "creation_date": str(first_row.get("Creation Date", "")),
            "completion_date": str(first_row.get("Completion Date", "")),
            "completed_by": first_row.get("Completed By", ""),
            "overall_comment": first_row.get("Overall Comment", ""),
            "score_percent": first_row.get("Score in Percent", ""),
            "alert_type": first_row.get("Alert Type", ""),
            "elements": []
        }

        for idx, row in group.iterrows():
            element_data = {
                "zone": row.get("Zone", ""),
                "location": row.get("Location", ""),
                "element": row.get("Element", ""),
                "score_factor": row.get("Score Factor", ""),
                "element_weight_percent": row.get("Element Weight In %", ""),
                "rating": row.get("Rating", ""),
                "element_score_percent": row.get("Element Score in %", ""),
                "comments": row.get("Comments", ""),
                "attachment": "NaN"
            }
            inspection_data["elements"].append(element_data)
            row_map[norm_id].append(idx)

        inspections.append(inspection_data)

    return inspections, row_map, df_complete

def extract_images_and_update_json(excel_path, inspections, row_map, df_complete):
    wb = xw.Book(excel_path)
    sheet = wb.sheets['Raw Data']
    os.makedirs('attachments', exist_ok=True)

    image_index_map = defaultdict(int)
    unmatched_images = []
    matched_images = 0
    errors = 0

    for pic in sheet.pictures:
        try:
            # Estimate row based on image position
            top = pic.api.Top
            row = int(top // sheet.range('A1').height) + 1
            inspection_id_raw = sheet.range(f'A{row}').value
            norm_id = normalize_inspection_id(inspection_id_raw)

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

            # Match to correct element in JSON
            matched = False
            for inspection in inspections:
                if inspection['inspection_id'] == norm_id:
                    for element in inspection['elements']:
                        if element['attachment'] == "NaN":
                            element['attachment'] = image_path.replace("\\", "/")
                            matched = True
                            matched_images += 1
                            print(f"[Image Saved] {image_filename} for Inspection #{norm_id}")
                            break
                    break

            if not matched:
                unmatched_images.append(norm_id)

        except Exception as e:
            print(f"[Error] Failed to process image: {e}")
            errors += 1

    return matched_images, unmatched_images, errors

if __name__ == "__main__":
    input_file = "audit_data.xlsx"
    output_file = "inspection_summary.json"

    try:
        print("[Step 1] Loading and grouping inspections...")
        inspections, row_map, df_complete = load_and_group_inspections(input_file)
        print(f"[Step 2] Found {len(df_complete)} 'Complete' rows.")
        print(f"[Step 3] Grouped into {len(inspections)} inspections.")
        print("[Step 4] Initial JSON structure created.")
        print("[Step 5] Extracting embedded images using xlwings...")

        matched_images, unmatched_images, errors = extract_images_and_update_json(
            input_file, inspections, row_map, df_complete
        )

        print("[Step 6] Image extraction complete. Saving updated JSON...")
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(inspections, f, indent=2, ensure_ascii=False)

        print(f"[Done] Inspection summary saved to '{output_file}'.")
        print("\n--- Summary Report ---")
        print(f"Total inspections parsed: {len(inspections)}")
        print(f"Images successfully matched and saved: {matched_images}")
        print(f"Unmatched images: {len(unmatched_images)}")
        if unmatched_images:
            print(f"Unmatched Inspection IDs: {set(unmatched_images)}")
        print(f"Errors during image processing: {errors}")

    except FileNotFoundError:
        print(f"[Error] File '{input_file}' not found.")
    except Exception as e:
        print(f"[Error] An unexpected error occurred: {e}")
