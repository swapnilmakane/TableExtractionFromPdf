# PDF Table Extraction API

A FastAPI-based application that extracts tables from PDF files, converts the detected regions into an intermediate Excel workbook, and returns structured JSON.

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
- **Structured JSON Output**: Converts extracted workbook data into a JSON response.

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
python main.py
```

You can also run the app directly with Uvicorn:

```bash
uvicorn pdf_table_extractor.app:app --reload
```

The `/convert` endpoint accepts a PDF upload and returns extracted JSON. Use the `pattern_name` query parameter to select a configured pattern. The older `customerName` query parameter is still accepted for compatibility.

### Configuring for Your PDF Structure

The application uses a pattern-based configuration system. Update `PATTERN_CONFIG` in `pdf_table_extractor/core/config.py` to match your PDF structure:

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

- `DPI`: Resolution for image processing (default: 300)
- `DEFAULT_PATTERN`: Pattern used when no query parameter is provided
- `PATTERN_CONFIG`: Table and region coordinates by pattern
- `EXTRACTOR_CONFIG`: Maps each pattern to its Excel-to-JSON extractor

## Output

The API returns JSON. Temporary PDF and Excel files are created during processing and removed after the request completes.

## Project Structure

```text
pdf_table_extractor/
  app.py                 FastAPI application setup
  api/
    routes.py            HTTP endpoints
  core/
    config.py            DPI, pattern coordinates, extractor registry
  services/
    pdf_conversion_service.py
                         PDF region detection and Excel conversion flow
  extractors/
    pattern_1.py         Pattern-specific Excel-to-JSON parser
main.py                  Local server entry point
```

## Notes

- The app is configured for a specific PDF structure by default. Adjust the `PATTERN_CONFIG` coordinates for different PDF layouts.
- The image processing coordinates (x, y) may need fine-tuning based on your specific PDF dimensions.
- Requires appropriate permissions to create temporary files while processing each request.
