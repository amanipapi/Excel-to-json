"""Reads Excel files in 'xlsx/', extreact the attendance by topic and students"""
import os
import json
import pandas as pd

# === CONFIG ===
INPUT_FOLDER = 'xlsx'
OUTPUT_FOLDER = 'attend_json'
HEADER_ROW_INDEX = 3  # Row 4 in Excel (0-indexed)
SKIP_KEYWORDS = ['remark', 'student']

# Make sure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# List all Excel files
excel_files = [f for f in os.listdir(INPUT_FOLDER) if f.endswith('.xlsx')]

for excel_file in excel_files:
    print(f"\nProcessing file: {excel_file}")
    excel_path = os.path.join(INPUT_FOLDER, excel_file)
    xls = pd.ExcelFile(excel_path)

    # We'll collect everything per topic
    file_data = []

    for sheet_name in xls.sheet_names:
        if any(keyword in sheet_name.lower() for keyword in SKIP_KEYWORDS):
            print(f"Skipping sheet: {sheet_name}")
            continue

        print(f"Reading topic: {sheet_name}")
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=HEADER_ROW_INDEX,
                               dtype={"Phone Number": str})

            topic_data = []

            for _, row in df.iterrows():
                full_name = row.get("Full Name ")
                phone = row.get("Phone Number")

                if pd.isna(full_name) or pd.isna(phone):
                    continue

                attendance = {}
                for col in df.columns:
                    if isinstance(col, str) and col.lower().startswith("session"):
                        VALUE = str(row.get(col))
                        MARK = VALUE.strip()[0].upper() if VALUE and VALUE.strip() else None
                        if MARK in ['P', 'A']:
                            attendance[col] = MARK
                        else:
                            attendance[col] = None

                topic_data.append({
                    "full_name": str(full_name).strip(),
                    "phone_number": str(phone).strip(),
                    "attendance": attendance
                })

            # Only add the topic if it has data
            if topic_data:
                file_data.append({
                    "topic": sheet_name,
                    "students": topic_data
                })

        except (ValueError, KeyError) as e:
            print(f"!!!Error in sheet '{sheet_name}': {e}")

    # Write to JSON if there was any data
    if file_data:
        OUTPUT_FILENAME = f"A_{os.path.splitext(excel_file)[0]}.json"
        output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(file_data, f, indent=4, ensure_ascii=False)

        print(f"Saved: {output_path}")
    else:
        print(f"!!No valid data found in {excel_file}")
