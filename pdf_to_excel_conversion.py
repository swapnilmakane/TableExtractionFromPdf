import cv2
import numpy as np
import fitz
import os
import tempfile

from pdf2docx import Converter
from spire.doc import *
from spire.doc.common import *
from spire.xls import *
from spire.xls.common import *

# =========================
# CONFIG
# =========================
PDF_PATH = "input.pdf"
OUTPUT_DIR = "output_tables"
FINAL_EXCEL = os.path.join(OUTPUT_DIR, "final_tables.xlsx")

PATTERN = "pattern_1"
os.makedirs(OUTPUT_DIR, exist_ok=True)

DPI = 300


PATTERN_CONFIG = {
    "pattern_1": {

        "part_details": {
            "type": "table",
            "x": (115, 125),
            "y": (405, 420),
            "pages": [0],
            "skip_header": False
        },

        "customer_id": {
            "type": "table",
            "x": (1780, 1790),
            "y": (405, 420),
            "pages": [0],
            "skip_header": False
        },

        "quantity": {
            "type": "table",
            "x": (2430, 2445),
            "y": (760, 770),
            "pages": [0],
            "skip_header": False
        },

        "operation": {
            "type": "table",
            "x": (115, 125),
            "y": (850, 860),
            "pages":None,
            "skip_header": True
        },

        "lot_details": {
            "type": "table",
            "x": (115, 2500),
            "y": (1100, 2000),
            "pages": [-1],
            "skip_header": False
        },

        "target_date": {
            "type": "table",
            "x": (115, 125),
            "y": (2060, 2070),
            "pages": [0],
            "skip_header": False
        },

        "work_order": {
            "type": "region",
            "x": (120, 1750),
            "y": (270, 400),
            "pages": [0]
        }
    }
}

# =========================
# GLOBAL STATE
# =========================
final_workbook = Workbook()
final_workbook.Worksheets.Clear()

sheets = {}
row_tracker = {}

# =========================
# PAGE VALIDATION (UPDATED)
# =========================
def is_page_valid(cfg, page_index, total_pages):
    pages = cfg.get("pages", None)

    if pages is None:
        return True

    if isinstance(pages, list) and len(pages) == 0:
        return False


    if -1 in pages:
        return page_index == total_pages - 1

    return page_index in pages

# =========================
# TABLE DETECTION
# =========================
def detect_table_type(x, y, page_index, total_pages):
    config = PATTERN_CONFIG[PATTERN]

    for name, cfg in config.items():
        if cfg["type"] != "table":
            continue

        if not is_page_valid(cfg, page_index, total_pages):
            continue

        if cfg["x"][0] <= x <= cfg["x"][1] and cfg["y"][0] <= y <= cfg["y"][1]:
            return name

    return None

# =========================
# Spire helpers 
# =========================
def CopyTextAndStyle(ws, cell, paragraph):
    cell.RichText.Text = paragraph.Text

def CopyContentInTable(tableCell, cell, worksheet):
    newParagraph = Paragraph(tableCell.Document)

    for i in range(tableCell.ChildObjects.Count):
        obj = tableCell.ChildObjects[i]

        if isinstance(obj, Paragraph):
            para = Paragraph(obj)

            for j in range(para.ChildObjects.Count):
                newParagraph.ChildObjects.Add(para.ChildObjects[j].Clone())

            if i < tableCell.ChildObjects.Count - 1:
                newParagraph.AppendText("\n")

    CopyTextAndStyle(worksheet, cell, newParagraph)

# =========================
# EXPORT TABLE
# =========================
def ExportTableInExcel(ws, row, table, skip_header=False):
    start_row = 1 if skip_header else 0

    for i in range(start_row, table.Rows.Count):
        col = 1

        for j in range(table.Rows[i].Cells.Count):
            cell = ws.Range[row, col]
            cell.BorderAround(LineStyleType.Thin, Color.get_Black())

            CopyContentInTable(
                table.Rows[i].Cells[j],
                cell,
                ws
            )

            col += 1

        row += 1

    return row

# =========================
# IMAGE PROCESSING
# =========================
def get_table_contours(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30))

    h = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel, 2)
    v = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, v_kernel, 2)

    mask = cv2.add(h, v)
    mask = cv2.dilate(mask, np.ones((3, 3), np.uint8), 1)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

# =========================
# PDF → DOCX
# =========================
def pdf_to_docx(doc, page_index, rect):
    new_doc = fitz.open()
    new_page = new_doc.new_page(width=rect.width, height=rect.height)
    new_page.show_pdf_page(new_page.rect, doc, page_index, clip=rect)

    pdf_bytes = new_doc.tobytes()
    new_doc.close()

    tmp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)

    try:
        tmp_pdf.write(pdf_bytes)
        tmp_pdf.close()
        tmp_docx.close()

        cv = Converter(tmp_pdf.name)
        cv.convert(tmp_docx.name)
        cv.close()

        document = Document()
        document.LoadFromFile(tmp_docx.name)

    finally:
        os.remove(tmp_pdf.name)
        os.remove(tmp_docx.name)

    return document

# =========================
# DOCX → EXCEL
# =========================
def convert(doc, table_type):
    cfg = PATTERN_CONFIG[PATTERN][table_type]

    if table_type not in sheets:
        sheets[table_type] = final_workbook.CreateEmptySheet(table_type)
        row_tracker[table_type] = 1

    ws = sheets[table_type]
    row = row_tracker[table_type]

    skip_header = cfg.get("skip_header", False)

    for i in range(doc.Sections.Count):
        section = doc.Sections[i]

        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects[j]

            if isinstance(obj, Table):
                row = ExportTableInExcel(ws, row, Table(obj), skip_header)

    row_tracker[table_type] = row
    doc.Dispose()

# =========================
# REGION EXTRACTION
# =========================
def extract_region(doc, page_index, name, cfg, sx, sy):
    x0, x1 = cfg["x"]
    y0, y1 = cfg["y"]

    rect = fitz.Rect(
        x0 * sx,
        y0 * sy,
        x1 * sx,
        y1 * sy
    )

    d = pdf_to_docx(doc, page_index, rect)
    convert(d, name)

    print(f"✅ REGION → {name}")

# =========================
# MAIN
# =========================
def process_pdf():
    doc = fitz.open(PDF_PATH)
    total_pages = len(doc)

    for page_index, page in enumerate(doc):

        pix = page.get_pixmap(dpi=DPI)
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        contours = get_table_contours(img)

        pw, ph = page.rect.width, page.rect.height
        ih, iw = img.shape[:2]

        sx = pw / iw
        sy = ph / ih

        config = PATTERN_CONFIG[PATTERN]

        # REGION
        for name, cfg in config.items():
            if cfg["type"] == "region" and is_page_valid(cfg, page_index, total_pages):
                extract_region(doc, page_index, name, cfg, sx, sy)

        # TABLES
        for c in contours:
            x, y, w, h = cv2.boundingRect(c)

            if w < 80 or h < 30:
                continue

            table_type = detect_table_type(x, y, page_index, total_pages)
            if not table_type:
                continue

            rect = fitz.Rect(
                x * sx,
                y * sy,
                (x + w) * sx,
                (y + h) * sy
            )

            d = pdf_to_docx(doc, page_index, rect)
            convert(d, table_type)

            print(f"✅ TABLE → {table_type}")

# =========================
# RUN
# =========================
process_pdf()

final_workbook.SaveToFile(FINAL_EXCEL, ExcelVersion.Version2016)
final_workbook.Dispose()

print(f"\nExcel saved: {FINAL_EXCEL}")