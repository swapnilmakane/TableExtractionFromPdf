# PDF to Excel Table Extraction

This Python script extracts tables from a PDF file and converts them into an Excel spreadsheet. It uses image processing to detect table regions, converts them to DOCX format, and then exports the tables to Excel sheets.

## Features

- Detects and extracts different types of tables (operation, part details, customer ID, quantity, lot details, target date)
- Processes multi-page PDFs
- Outputs organized Excel sheets with proper formatting

## Installation

1. Clone or download this repository.
2. Install the required dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Dependencies

- opencv-python
- numpy
- PyMuPDF
- pdf2docx
- spire.doc
- spire.xls

## Usage

1. Update the `PDF_PATH` and `OUTPUT_DIR` variables in the script to point to your input PDF and desired output directory.
2. Run the script:

   ```bash
   python pdf_to_excel_conversion.py
   ```

3. The extracted tables will be saved in `final_tables.xlsx` in the specified output directory.

## Notes

- The script is configured for a specific PDF structure. You may need to adjust the `detect_table_type` function and coordinates for different PDF layouts.
- Ensure the input PDF path and output directory exist and are accessible.