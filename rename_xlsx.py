"""Renames all teh excel files for leading json  """
import os
import re
import shutil

SOURCE_FOLDER = "original_excel"
DEST_FOLDER = "xlsx"
os.makedirs(DEST_FOLDER, exist_ok=True)

# Pattern to find "foundation 1 batch 123" or just "foundation 1 batch"
pattern = re.compile(r"foundation\s*(\d)\s*batch\s*(\d+)?", re.IGNORECASE)

def get_unique_filename(dest_folder, base_name):
    """Make sure there is no over writing of data"""
    name, ext = os.path.splitext(base_name)
    counter = 1
    new_name = base_name
    while os.path.exists(os.path.join(dest_folder, new_name)):
        new_name = f"{name}_{counter}{ext}"
        counter += 1
    return new_name

for filename in os.listdir(SOURCE_FOLDER):
    if filename.lower().endswith(".xlsx"):
        match = pattern.search(filename)
        if match:
            level = match.group(1)
            batch = match.group(2).zfill(3) if match.group(2) else "000"
            BASE_FILENAME = f"F_{level}_B_{batch}.xlsx"
            SAFE_FILENAME = get_unique_filename(DEST_FOLDER, BASE_FILENAME)

            old_path = os.path.join(SOURCE_FOLDER, filename)
            new_path = os.path.join(DEST_FOLDER, SAFE_FILENAME)

            shutil.move(old_path, new_path)
            print(f"Moved & Renamed: {filename} ➜ {SAFE_FILENAME}")
        else:
            print(f"!!!Skipped (no match): {filename}")
