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
PDF_PATH = r"C:\Users\Maka_swa\Documents\Sample\RE_ WT-100　添付資料\フロコンスプール(コムテスコ).pdf"
OUTPUT_DIR = r"C:\Users\Maka_swa\Documents\Sample\output_tables"
FINAL_EXCEL = os.path.join(OUTPUT_DIR, "final_tables.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

DPI = 300

# =========================
# GLOBAL STATE
# =========================
final_workbook = Workbook()
final_workbook.Worksheets.Clear()

captured_once = {
    "target_date": False,
    "quantity": False,
    "part_details": False,
    "customer_id": False,
    "lot_details": False
}

operation_header_written = False

operation_sheet = final_workbook.CreateEmptySheet("Operation")
operation_row = 1

sheets_single = {}

# =========================
# DETECT TABLE TYPE
# =========================
def detect_table_type(x, y, w, h):

    if y > 2000:
        return "target_date"

    if y > 800 and h > 300:
        return "operation"

    if y > 1200 and h < 200:
        return "lot_details"

    if y > 350 and y < 600 and x > 1500:
        return "customer_id"

    if x > 2000 and h < 150:
        return "quantity"

    if y > 350 and y < 600 and w < 2000:
        return "part_details"

    return None



def CopyTextAndStyle(worksheet: Worksheet, cell: CellRange, paragraph: Paragraph):
    cell.RichText.Text = paragraph.Text


def CopyContentInTable(tableCell: TableCell, cell: CellRange, worksheet: Worksheet):
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


def ExportTableInExcel(worksheet: Worksheet, row: int, table: Table, skip_header=False):
    start_row = 1 if skip_header else 0

    for i in range(start_row, table.Rows.Count):
        col = 1
        for j in range(table.Rows[i].Cells.Count):
            cell = worksheet.Range[row, col]
            cell.BorderAround(LineStyleType.Thin, Color.get_Black())
            CopyContentInTable(table.Rows[i].Cells[j], cell, worksheet)
            col += 1
        row += 1
    return row


# =========================
# IMAGE PROCESSING
# =========================
def get_table_contours(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30))

    horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)

    table_mask = cv2.add(horizontal_lines, vertical_lines)
    table_mask = cv2.dilate(table_mask, np.ones((3, 3), np.uint8), iterations=1)

    contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

    return contours


# =========================
# PDF CROP + CONVERT
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
        if os.path.exists(tmp_pdf.name):
            os.remove(tmp_pdf.name)
        if os.path.exists(tmp_docx.name):
            os.remove(tmp_docx.name)

    return document


# =========================
# Work Order Number
# =========================
def get_work_order_number(doc, page_index, x, y, w, h, scale_x, scale_y):
    extra_top = max(0, y - 140)
    extra_bottom = y

    extra_rect = fitz.Rect(
        x * scale_x,
        extra_top * scale_y,
        (x + w) * scale_x,
        extra_bottom * scale_y
    )

    extra_document = convert_pdf_region_to_docx(doc, page_index, extra_rect)

    sheet_name = "part_details_header"
    if sheet_name not in sheets_single:
        sheets_single[sheet_name] = final_workbook.CreateEmptySheet(sheet_name)

    ws = sheets_single[sheet_name]
    row_header = 1

    for i in range(extra_document.Sections.Count):
        section = extra_document.Sections[i]
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects[j]
            if isinstance(obj, Table):
                row_header = ExportTableInExcel(ws, row_header, Table(obj))

    extra_document.Dispose()


# =========================
# Convert docx to excel
# =========================
def convert_docx_to_excel(document, table_type):
    global operation_row, operation_header_written

    if table_type == "operation":
        worksheet = operation_sheet
        row = operation_row
    else:
        if table_type not in sheets_single:
            sheets_single[table_type] = final_workbook.CreateEmptySheet(table_type)
        worksheet = sheets_single[table_type]
        row = 1

    for i in range(document.Sections.Count):
        section = document.Sections[i]
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects[j]

            if isinstance(obj, Table):

                if table_type == "operation":
                    row = ExportTableInExcel(
                        worksheet,
                        row,
                        Table(obj),
                        skip_header=operation_header_written
                    )
                    operation_header_written = True
                else:
                    row = ExportTableInExcel(worksheet, row, Table(obj))

    document.Dispose()

    if table_type == "operation":
        operation_row = row
    else:
        captured_once[table_type] = True


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

        pdf_width = page.rect.width
        pdf_height = page.rect.height

        img_h, img_w = image.shape[:2]
        scale_x = pdf_width / img_w
        scale_y = pdf_height / img_h

        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)

            if w < 80 or h < 30:
                continue

            table_type = detect_table_type(x, y, w, h)
            if table_type is None:
                continue

            if table_type != "operation" and captured_once.get(table_type, False):
                continue

            x0 = x * scale_x
            x1 = (x + w) * scale_x
            y0 = y * scale_y
            y1 = (y + h) * scale_y

            rect = fitz.Rect(x0, y0, x1, y1)

            
            if table_type == "part_details" and page_index == 0:
                get_work_order_number(doc, page_index, x, y, w, h, scale_x, scale_y)

            # MAIN TABLE
            document = convert_pdf_region_to_docx(doc, page_index, rect)
            convert_docx_to_excel(document, table_type)

            print(f"✅ {table_type} processed")


# =========================
# RUN
# =========================
process_pdf()

final_workbook.SaveToFile(FINAL_EXCEL, ExcelVersion.Version2016)
final_workbook.Dispose()

print(f"\nExcel saved: {FINAL_EXCEL}")