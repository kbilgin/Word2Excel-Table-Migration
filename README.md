# Word2Excel-Table-Migration
This script provides a simple utility to convert tables from a .docx file to a .xlsx spreadsheet. Each table from the .docx file will be written to a separate sheet in the Excel workbook.

Requirements

To run this script, you need:

Python 3.x
python-docx package: To read .docx files.
openpyxl package: To create and save .xlsx files.
You can install these packages using pip:

pip install python-docx openpyxl

Usage

Modify the script to specify your input .docx file's path and the desired output .xlsx file's path.
Run the script:
python your_script_name.py

Replace your_script_name.py with the name you've saved the script as.

After execution, you'll find the .xlsx file with tables extracted from the provided .docx file. Each table will be in a separate sheet named 'Table X' where X is the table number.
Limitations

This script assumes that the provided .docx file contains only tables. If there are other elements, they will be ignored.
Formatting and styles from the .docx tables are not preserved in the .xlsx file.
