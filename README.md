# PDF to Excel Table Extraction

A FastAPI-based application that extracts tables from PDF files and converts them into Excel spreadsheets. It uses advanced image processing techniques to detect table regions, converts them to DOCX format, and exports organized data to Excel sheets.

## Features

- **Intelligent Table Detection**: Automatically detects and extracts different types of tables including:
  - Operation tables
  - Part details
  - Customer ID
  - Quantity information
  - Lot details
  - Target date
- **Multi-page Processing**: Handles PDFs with multiple pages
- **Pattern-based Configuration**: Supports configurable table patterns for different PDF layouts
- **REST API**: Built with FastAPI for easy integration and file uploads
- **Organized Output**: Exports data to properly formatted Excel sheets (`final_tables.xlsx`)

## Installation

1. Clone or download this repository.
2. Install the required dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Dependencies

- **opencv-python**: Image processing and table region detection
- **numpy**: Numerical computations for image analysis
- **PyMuPDF**: PDF file reading and page manipulation
- **pdf2docx**: PDF to DOCX conversion
- **spire.doc**: DOCX document processing
- **spire.xls**: Excel file manipulation and export
- **fastapi**: REST API framework
- **openpyxl**: Excel workbook management

## Usage

### Running the FastAPI Server

```bash
python pdf_to_excel_conversion.py
```

The server will start and be accessible via FastAPI endpoints for PDF upload and table extraction.

### Configuring for Your PDF Structure

The application uses a pattern-based configuration system. Update the `PATTERN_CONFIG` dictionary in the script to match your PDF structure:

```python
PATTERN_CONFIG = {
    "pattern_1": {
        "part_details": {"type": "table", "x": (115, 125), "y": (405, 420), "pages": [0]},
        "customer_id": {"type": "table", "x": (1780, 1790), "y": (405, 420), "pages": [0]},
        # ... more table definitions
    }
}
```

- **x**: X-coordinate range for table detection
- **y**: Y-coordinate range for table detection
- **pages**: Specific pages to process (None = all pages)
- **type**: Either "table" or "region" for different detection methods

## Configuration

- `PDF_PATH`: Path to your input PDF file
- `OUTPUT_DIR`: Directory where extracted Excel files will be saved
- `DPI`: Resolution for image processing (default: 300)
- `PATTERN`: Active pattern configuration to use

## Output

The script generates `final_tables.xlsx` with extracted table data organized into separate sheets based on table type.

## Notes

- The script is configured for a specific PDF structure by default. Adjust the `PATTERN_CONFIG` coordinates for different PDF layouts.
- Ensure input PDF paths and output directories are accessible before running.
- The image processing coordinates (x, y) may need fine-tuning based on your specific PDF dimensions.
- Requires appropriate permissions to read input PDFs and write to the output directory.