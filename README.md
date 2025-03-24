# GLC Excel to JSON Converter
This tool processes Excel files with student data and converts them to clean JSON files based on specific rules for Restoration (`F_1`) and Foundation (`F_2`) formats.

## 📁 Project Structure

project/ 
    ├── original_excel/ # original source files 
    ├── xlsx/ # Excel source files 
    ├── json/ # Output JSON files 
    ├── loading_json_v2.py # Main script 
    ├── remane_xlsx.py # Renaming script
    ├── .gitignore 
    └── README.md


## 🚀 Features

- Automatically renames Excel files based on content
- Handles multiple batch formats (e.g., `F_1_B_123`)
- Extracts sheets with "student" in the name
- Cleans and validates data
- Outputs JSON in `json/` folder

## 🔧 Usage
Create a virtual enviroment using bash or any terminal
    python -m venv myenv
activate the venv
    activate myenv/Scripts/activate
Create folders
    original_excel
    xlsx
using bash or any terminal
    python rename-xlsx.py
    python loading_json_v2.py
Make sure Excel files are placed inside the original_excel/ folder before running the script.

## 📦 Requirements

Python 3.7+
pandas
openpyxl

## Install dependencies

- pip install pandas
- pip install openpyxl

## ✅ Output
Cleaned .json files are saved to the json/ folder using the base name of the Excel file.

## 📝 License
MIT License