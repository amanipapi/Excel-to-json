# GLC Excel to JSON Converter
This tool processes Excel files with student data and converts them to clean JSON files based on specific rules for Restoration (`F_1`) and Foundation (`F_2`) formats.

## 📁 Project Structure

project/   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    ├── original_excel/ # original source files   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    ├── xlsx/ # Excel source files   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   ├── json/ # Output JSON files       
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   ├── attend_json/ # Output JSON files for attendace   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    ├── loading_json_v2.py # Main script   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   ├── remane_xlsx.py # Renaming script  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   ├── .gitignore   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    └── README.md  


## 🚀 Features

- Automatically renames Excel files based on content
- Handles multiple batch formats (e.g., `F_1_B_123`)
- Extracts sheets with "student" in the name
- Cleans and validates data
- Outputs JSON in `json/` folder

## 🔧 Usage
Create a virtual enviroment using bash or any terminal  
&nbsp;&nbsp;&nbsp;&nbsp;    python -m venv myenv  
activate the venv  
&nbsp;&nbsp;&nbsp;&nbsp;    activate myenv/Scripts/activate  
Create folders  
&nbsp;&nbsp;&nbsp;&nbsp;    original_excel  
&nbsp;&nbsp;&nbsp;&nbsp;    xlsx  
using bash or any terminal  
&nbsp;&nbsp;&nbsp;&nbsp;    python rename-xlsx.py  
&nbsp;&nbsp;&nbsp;&nbsp;    python loading_json_v2.py  
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
