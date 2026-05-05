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
PDF_PATH = r"C:\Users\Maka_swa\Documents\Sample\RE_ WT-100　添付資料\ハ゛ランステ゛ィスクA(島津三条).pdf"
OUTPUT_DIR = "output_tables"
FINAL_EXCEL = os.path.join(OUTPUT_DIR, "final_tables.xlsx")

PATTERN = "pattern 1"

os.makedirs(OUTPUT_DIR, exist_ok=True)

DPI = 300

# =========================
# PATTERN CONFIG 
# =========================
PATTERN_CONFIG = {
    "pattern 1": {
        "part_details": {
            "type": "table",
            "x": (115,125),
            "y": (405,420),
            "pages": [0]
        },
        "customer_id": {
            "type": "table",
            "x": (1780,1790),
            "y": (405,420),
            "pages": [0]
        },
        "quantity": {
            "type": "table",
            "x": (2430,2445),
            "y": (760,770),
            "pages": [0]
        },
        "operation": {
            "type": "table",
            "x": (115,125),
            "y": (850,860),
            "pages": [0,1]   # multi-page example
        },
        "lot_details": {
            "type": "table",
            "x": (115,125),
            "y": (1235,1250),
            "pages": [0]
        },
        "target_date": {
            "type": "table",
            "x": (115,125),
            "y": (2060,2070),
            "pages": [0]
        },
        "work_order": {
            "type": "region",
            "x": (100,2500),
            "y": (100,350),
            "pages": [0]
        }
    }
}

# =========================
# GLOBAL STATE
# =========================
final_workbook = Workbook()
final_workbook.Worksheets.Clear()

captured_once = {}
operation_header_written = False

operation_sheet = final_workbook.CreateEmptySheet("operation")
operation_row = 1

sheets_single = {}

# =========================
# PAGE FILTER
# =========================
def is_page_valid(cfg, page_index):
    pages = cfg.get("pages")
    if pages is None:
        return True
    return page_index in pages

# =========================
# TABLE DETECTION
# =========================
def detect_table_type(x, y, page_index):
    config = PATTERN_CONFIG[PATTERN]

    for name, cfg in config.items():
        if cfg["type"] != "table":
            continue

        if not is_page_valid(cfg, page_index):
            continue

        if cfg["x"][0] <= x <= cfg["x"][1] and \
           cfg["y"][0] <= y <= cfg["y"][1]:
            return name

    return None


def CopyText(cell, paragraph):
    cell.RichText.Text = paragraph.Text

def CopyContentInTable(tableCell, cell, ws):
    newParagraph = Paragraph(tableCell.Document)

    for i in range(tableCell.ChildObjects.Count):
        obj = tableCell.ChildObjects[i]

        if isinstance(obj, Paragraph):
            para = Paragraph(obj)

            for j in range(para.ChildObjects.Count):
                newParagraph.ChildObjects.Add(para.ChildObjects[j].Clone())

            if i < tableCell.ChildObjects.Count - 1:
                newParagraph.AppendText("\n")

    CopyText(cell, newParagraph)

def ExportTableInExcel(ws, row, table, skip_header=False):
    start_row = 1 if skip_header else 0

    for i in range(start_row, table.Rows.Count):
        col = 1

        for j in range(table.Rows[i].Cells.Count):
            cell = ws.Range[row, col]
            cell.BorderAround(LineStyleType.Thin, Color.get_Black())

            CopyContentInTable(table.Rows[i].Cells[j], cell, ws)
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

    h_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel, iterations=2)
    v_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, v_kernel, iterations=2)

    mask = cv2.add(h_lines, v_lines)
    mask = cv2.dilate(mask, np.ones((3, 3), np.uint8), iterations=1)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

# =========================
# PDF → DOCX
# =========================
def convert_pdf_region_to_docx(doc, page_index, rect):
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
def convert_docx_to_excel(document, table_type):
    global operation_row, operation_header_written

    if table_type == "operation":
        ws = operation_sheet
        row = operation_row
    else:
        if table_type not in sheets_single:
            sheets_single[table_type] = final_workbook.CreateEmptySheet(table_type)

        ws = sheets_single[table_type]
        row = 1


    for i in range(document.Sections.Count):
        section = document.Sections[i]

        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects[j]

            if isinstance(obj, Table):
                if table_type == "operation":
                    row = ExportTableInExcel(
                        ws,
                        row,
                        Table(obj),
                        skip_header=operation_header_written
                    )
                    operation_header_written = True
                else:
                    row = ExportTableInExcel(ws, row, Table(obj))

    document.Dispose()

    if table_type == "operation":
        operation_row = row
    else:
        captured_once[table_type] = True

# =========================
# REGION EXTRACTION
# =========================
def extract_region(doc, page_index, name, cfg, scale_x, scale_y):
    x_min, x_max = cfg["x"]
    y_min, y_max = cfg["y"]

    rect = fitz.Rect(
        x_min * scale_x,
        y_min * scale_y,
        x_max * scale_x,
        y_max * scale_y
    )

    document = convert_pdf_region_to_docx(doc, page_index, rect)
    convert_docx_to_excel(document, name)

    print(f"✅ REGION → {name}")

# =========================
# MAIN PROCESS
# =========================
def process_pdf():
    doc = fitz.open(PDF_PATH)

    for page_index, page in enumerate(doc):

        pix = page.get_pixmap(dpi=DPI)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        image = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        contours = get_table_contours(image)

        pdf_w, pdf_h = page.rect.width, page.rect.height
        img_h, img_w = image.shape[:2]

        scale_x = pdf_w / img_w
        scale_y = pdf_h / img_h

        config = PATTERN_CONFIG[PATTERN]

        # 🔹 REGION FIRST
        for name, cfg in config.items():
            if cfg["type"] == "region" and is_page_valid(cfg, page_index):
                extract_region(doc, page_index, name, cfg, scale_x, scale_y)

        # 🔹 TABLES
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)

            if w < 80 or h < 30:
                continue

            table_type = detect_table_type(x, y, page_index)
            if table_type is None:
                continue

            if table_type != "operation" and captured_once.get(table_type, False):
                continue

            rect = fitz.Rect(
                x * scale_x,
                y * scale_y,
                (x + w) * scale_x,
                (y + h) * scale_y
            )

            document = convert_pdf_region_to_docx(doc, page_index, rect)
            convert_docx_to_excel(document, table_type)

            print(f"✅ TABLE → {table_type}")

# =========================
# RUN
# =========================
process_pdf()

final_workbook.SaveToFile(FINAL_EXCEL, ExcelVersion.Version2016)
final_workbook.Dispose()

print(f"\nExcel saved: {FINAL_EXCEL}")