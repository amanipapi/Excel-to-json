"""Reads Excel files in 'xlsx/', extracts sheets with 'student' in the name,
cleans data based on file prefix (F_1 or F_2), and saves to JSON."""

import os
import json
import pandas as pd

INPUT_FOLDER = 'xlsx'
OUTPUT_FOLDER = 'json'

# Create the output folder if it doesn't exist
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def is_valid_id(value):
    """Checks if the ID is a number"""
    return pd.notna(value) and str(value).strip().isdigit()

def is_valid_name(value):
    """Checks if the data has a full name"""
    return pd.notna(value) and str(value).strip() != ''

def safe_value(val):
    """Return a null value for the json"""
    return None if pd.isna(val) else val

def process_f1_row(row):
    """Process a row from F_1 Excel files and return cleaned data."""
    if not (is_valid_id(row.get("GLORIOUS LIFE CHURCH")) and is_valid_name(row.get("Unnamed: 1"))):
        return None
    return {
        "id": int(row["GLORIOUS LIFE CHURCH"]),
        "full_name": str(row["Unnamed: 1"]).strip(),
        "phone_number": safe_value(row.get("Unnamed: 2")),
        "restoration_part_1": safe_value(row.get("Unnamed: 3")),
        "restoration_part_2": safe_value(row.get("Unnamed: 4")),
        "restoration_part_3": safe_value(row.get("Unnamed: 5")),
        "previous_church": safe_value(row.get("Unnamed: 6"))
    }

def process_f2_row(row):
    """Process a row from F_2 Excel files and return cleaned data."""
    if not (is_valid_id(row.get("GLORIOUS LIFE CHURCH ")) and is_valid_name(row.get("Unnamed: 1"))):
        return None
    return {
        "id": int(row["GLORIOUS LIFE CHURCH "]),
        "full_name": str(row["Unnamed: 1"]).strip(),
        "phone_number": safe_value(row.get("Unnamed: 2")),
        "foundation_1_teacher": safe_value(row.get("Unnamed: 3")),
        "foundation_1_batch": safe_value(row.get("Unnamed: 4"))
    }

for filename in os.listdir(INPUT_FOLDER):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(INPUT_FOLDER, filename)
        print(f"Processing {filename}...")

        try:
            xls = pd.ExcelFile(file_path)
            student_sheets = [s for s in xls.sheet_names if "student" in s.lower()]
            if not student_sheets:
                print(f"No 'student' sheet found in {filename}, skipping.")
                continue

            result = []

            for sheet in student_sheets:
                df = pd.read_excel(xls, sheet_name=sheet)

                if filename.startswith("F_1"):
                    cleaned = list(filter(None, map(process_f1_row, df.to_dict(orient='records'))))
                elif filename.startswith("F_2"):
                    cleaned = list(filter(None, map(process_f2_row, df.to_dict(orient='records'))))
                else:
                    print(f"Unknown file prefix in {filename}, skipping.")
                    continue

                result.extend(cleaned)

            if result:
                json_filename = os.path.splitext(filename)[0] + '.json'
                output_path = os.path.join(OUTPUT_FOLDER, json_filename)

                with open(output_path, 'w', encoding='utf-8') as json_file:
                    json.dump(result, json_file, indent=4, ensure_ascii=False)

                print(f"Saved cleaned data to {output_path}")
            else:
                print(f"!!No valid data found in {filename}")

        except (ValueError, KeyError) as e:
            print(f"!!!Error processing {filename}: {e}")
