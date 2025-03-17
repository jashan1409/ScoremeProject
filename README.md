Name-Jashanjit Singh
SID=21103087
PDF Table Extraction Project
Overview
This project extracts tables from a text-based PDF and saves them in an Excel file. The script follows these steps:
Extracts text from the PDF using PyPDF2.
Parses and structures tables from the extracted text.
Cleans and formats the data.
Saves the extracted tables into an Excel file using openpyxl.
Requirements
Install the necessary Python libraries before running the script:
pip install PyPDF2 openpyxl

Usage
Set the File Paths


Update pdf_path with the path to your input PDF file.
Update excel_path with the desired output location for the extracted Excel file.
Run the Script


python extract_tables.py

Output
The extracted tables will be stored in an Excel file at the specified excel_path.
Each table will be separated by an empty row for clarity.
Functions
extract_text_from_pdf(pdf_path): Extracts text from the PDF.
parse_tables(text): Identifies and structures tables from extracted text.
clean_data(cell): Cleans special characters from table data.
write_to_excel(tables, excel_path): Saves structured data to an Excel file.
