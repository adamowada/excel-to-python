# Excel-to-Python Converter

## Adam Owada

This project converts an Excel file with formulas into a Python program with equivalent functionality. The conversion includes:

- Parsing Excel sheets and extracting formulas.
- Organizing data into `pandas` DataFrames.
- Generating a Python CLI that mimics the functionality of the Excel file.

## Requirements

Install the necessary Python libraries:

```bash
pip install -r requirements.txt
```

## Usage

1. Run the script, providing the path to your Excel file:

```bash
python script.py path/to/your/excel_file.xlsx
```

2. The tool will generate:

   - CSV files with the parsed data in the `outputs/<excel_file_name>/dataframes/` directory.
   - A Python CLI app in `outputs/<excel_file_name>/main.py`.

3. Navigate to the generated `main.py` and run the CLI app:

```bash
python outputs/<excel_file_name>/main.py
```

## Notes

- This tool supports complex Excel formulas but does not handle VBA macros or external data connections.
- The generated Python CLI replicates the functionality of the original Excel file.
