import cv2
import numpy as np
import fitz
import os
import tempfile
from datetime import datetime

from fastapi import FastAPI, UploadFile, File

from pdf2docx import Converter
from spire.doc import *
from spire.doc.common import *
from spire.xls import *
from spire.xls.common import *

from openpyxl import load_workbook

# =========================
# APP INIT
# =========================
app = FastAPI()

DPI = 300

# =========================
# PATTERN CONFIG
# =========================
PATTERN = "pattern_1"

PATTERN_CONFIG = {
    "pattern_1": {
        "part_details": {"type": "table", "x": (115, 125), "y": (405, 420), "pages": [0]},
        "customer_id": {"type": "table", "x": (1780, 1790), "y": (405, 420), "pages": [0]},
        "quantity": {"type": "table", "x": (2430, 2445), "y": (760, 770), "pages": [0]},
        "operation": {"type": "table", "x": (115, 125), "y": (850, 860), "pages": None},
        "target_date": {"type": "table", "x": (115, 125), "y": (2060, 2070), "pages": [0]},
        "work_order": {"type": "region", "x": (120, 1750), "y": (270, 400), "pages": [0]}
    }
}

# =========================
# HELPERS
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


def ExportTableInExcel(ws, row, table, table_type, header_written):
    if table_type not in header_written:
        header_written[table_type] = False

    start_row = 1 if header_written[table_type] else 0

    for i in range(start_row, table.Rows.Count):
        col = 1
        for j in range(table.Rows[i].Cells.Count):
            cell = ws.Range[row, col]
            cell.BorderAround(LineStyleType.Thin, Color.get_Black())
            CopyContentInTable(table.Rows[i].Cells[j], cell, ws)
            col += 1
        row += 1

    header_written[table_type] = True
    return row


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


def convert_doc_to_excel(doc, table_type, sheets, row_tracker, workbook, header_written):
    if table_type not in sheets:
        sheets[table_type] = workbook.CreateEmptySheet(table_type)
        row_tracker[table_type] = 1

    ws = sheets[table_type]
    row = row_tracker[table_type]

    for i in range(doc.Sections.Count):
        section = doc.Sections[i]
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects[j]
            if isinstance(obj, Table):
                row = ExportTableInExcel(ws, row, Table(obj), table_type, header_written)

    row_tracker[table_type] = row
    doc.Dispose()


# =========================
# CORE PROCESS
# =========================
def process_pdf(pdf_path, output_excel_path):
    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    workbook = Workbook()
    workbook.Worksheets.Clear()

    sheets = {}
    row_tracker = {}
    header_written = {}  # ✅ FIXED

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
                rect = fitz.Rect(
                    cfg["x"][0] * sx,
                    cfg["y"][0] * sy,
                    cfg["x"][1] * sx,
                    cfg["y"][1] * sy
                )

                d = pdf_to_docx(doc, page_index, rect)
                convert_doc_to_excel(d, name, sheets, row_tracker, workbook, header_written)

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
            convert_doc_to_excel(d, table_type, sheets, row_tracker, workbook, header_written)

    workbook.SaveToFile(output_excel_path, ExcelVersion.Version2016)
    workbook.Dispose()
    doc.close()


# =========================
# EXCEL → JSON
# =========================
def extract_json_pattern_1(excel_path):
    wb = load_workbook(excel_path, data_only=True)

    def get(sheet, cell):
        if sheet not in wb:
            return None
        return wb[sheet][cell].value

    def to_str(val):
        if val is None:
            return None

        text = str(val)

        # Replace Excel encoded vertical tab
        text = text.replace("_x000B_", "\n")

        # Also handle actual VT char if present
        text = text.replace("\x0b", "\n")

        return text.strip()

    def to_int(val):
        if val is None:
            return None

        try:
            cleaned = str(val).replace(",", "").strip()

            return int(cleaned)

        except:
            return None

    def clean_target_date(val):
        if val is None:
            return None

        # Convert to string
        val = str(val)

        # Split using full-width colon (ASCII 65306 → '：')
        if "：" in val:
            val = val.split("：")[-1].strip()

        # Try parsing date
        try:
            dt = datetime.fromisoformat(val)
            return dt.date().isoformat()
        except:
            try:
                dt = datetime.strptime(val, "%d-%m-%Y")
                return dt.date().isoformat()
            except:
                return val  # fallback (raw)

    result = {
        "WorkOrderNumber": "".join(filter(None, [
            to_str(get("work_order", "B1")),
            to_str(get("work_order", "C1")),
            to_str(get("work_order", "D1")),
        ])),
        "CustomerOrderId": to_str(get("customer_id", "B1")),
        "Quantity": to_int(get("quantity", "B1")),
        "PartNumber": to_str(get("part_details", "B2")),
        "PartName": to_str(get("part_details", "B5")),
        "DrawingNumber": to_str(get("part_details", "B3")),
        "DueDate": clean_target_date(get("target_date", "C2")),
        "Operations": []
    }

    # =========================
    # OPERATIONS (SKIP HEADER ROW)
    # =========================
    if "operation" in wb:
        ws = wb["operation"]
        row = 2  # ✅ skip header

        while True:
            op_num = ws[f"A{row}"].value
            op_name = ws[f"B{row}"].value

            if not op_num and not op_name:
                break

            result["Operations"].append({
                "OperationNumber": to_str(op_num),
                "OperationName": to_str(op_name),
                "WorkCenterGroup": to_str(ws[f"D{row}"].value),
                "IsInHouse": bool(ws[f"C{row}"].value) if ws[f"C{row}"].value is not None else False
            })

            row += 1

    return result

def extract_json_from_excel(path, pattern):
    if pattern == "pattern_1":
        return extract_json_pattern_1(path)
    raise ValueError("Unsupported pattern")

# =========================
# API
# =========================
@app.post("/convert")
async def convert_api(file: UploadFile = File(...)):
    import time
    import uuid

    pdf_bytes = await file.read()

    temp_pdf_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4()}.pdf")
    temp_excel_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4()}.xlsx")

    try:
        with open(temp_pdf_path, "wb") as f:
            f.write(pdf_bytes)

        process_pdf(temp_pdf_path, temp_excel_path)

        structured_data = extract_json_from_excel(temp_excel_path,PATTERN)

        return structured_data  # ✅ CLEAN RESPONSE

    finally:
        time.sleep(0.5)

        if os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
            except:
                pass

        if os.path.exists(temp_excel_path):
            try:
                os.remove(temp_excel_path)
            except:
                pass