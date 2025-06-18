# Text To XL

A simple and interactive Python script that allows users to input custom tabular data and export it as an Excel (.xlsx) file using the `pandas` library.

---

## Features

- Validates Excel file names (prevents illegal characters, duplicates, and bad formats)
- Add as many columns and rows as you want, interactively
- Ensures unique, non-empty column names
- Automatically saves to a specified folder (`./main-script/MyXL`)
- Creates the Excel file using `pandas` and `openpyxl`
- Built-in error handling for a smoother experience

---

## Requirements File

This project includes a `requirements.txt` file located at the root of the repository. It lists all the Python dependencies needed to run the script.

### Installing Dependencies from `requirements.txt`

To install all required packages at once, run the following command in your terminal:

### terminal:
pip install -r requirements.txt

---

## How It Works

1. Run the script
2. Enter the name of the Excel file
3. Enter the number of columns and name each one
4. Input data row by row
5. Choose whether to add another row or finish
6. Your Excel file will be saved in `./main-script/MyXL`

---

## Notes

- The script does not overwrite existing Excel files. If a file with the same name exists, the user will be prompted to choose a different name.
- The Excel files are saved to `./main-script/MyXL`. Make sure this directory exists, or the script will attempt to create it automatically.
- The script uses `pandas.DataFrame.to_excel()` which requires `openpyxl` as a dependency.
- Column names must be unique and cannot be empty.
