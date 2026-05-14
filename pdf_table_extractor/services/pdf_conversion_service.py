import os
import tempfile
import time
import uuid
from contextlib import suppress

import cv2
import fitz
import numpy as np
from fastapi import UploadFile
from pdf2docx import Converter
from spire.doc import *
from spire.doc.common import *
from spire.xls import *
from spire.xls.common import *

from pdf_table_extractor.core.config import (
    DPI,
    get_pattern_config,
    get_pattern_extractor,
)


def extract_json_from_excel(excel_path, pattern_name):
    extractor = get_pattern_extractor(pattern_name)
    return extractor(excel_path)


def is_page_in_scope(config, page_index, total_pages):
    pages = config.get("pages")
    if pages is None:
        return True
    if isinstance(pages, list) and len(pages) == 0:
        return False
    if -1 in pages:
        return page_index == total_pages - 1
    return page_index in pages


def detect_table_type(x, y, page_index, total_pages, pattern_name):
    pattern_config = get_pattern_config(pattern_name)

    for table_name, table_config in pattern_config.items():
        if table_config["type"] != "table":
            continue
        if not is_page_in_scope(table_config, page_index, total_pages):
            continue
        if table_config["x"][0] <= x <= table_config["x"][1] and table_config["y"][0] <= y <= table_config["y"][1]:
            return table_name

    return None


def copy_text_and_style(cell, paragraph):
    cell.RichText.Text = paragraph.Text


def copy_table_cell_content(table_cell, cell):
    paragraph = Paragraph(table_cell.Document)

    for index in range(table_cell.ChildObjects.Count):
        child_object = table_cell.ChildObjects[index]
        if isinstance(child_object, Paragraph):
            source_paragraph = Paragraph(child_object)
            for child_index in range(source_paragraph.ChildObjects.Count):
                paragraph.ChildObjects.Add(source_paragraph.ChildObjects[child_index].Clone())
            if index < table_cell.ChildObjects.Count - 1:
                paragraph.AppendText("\n")

    copy_text_and_style(cell, paragraph)


def export_table_to_excel(worksheet, row, table, table_type, header_written):
    if table_type not in header_written:
        header_written[table_type] = False

    start_row = 1 if header_written[table_type] else 0

    for table_row_index in range(start_row, table.Rows.Count):
        column = 1
        for cell_index in range(table.Rows[table_row_index].Cells.Count):
            cell = worksheet.Range[row, column]
            cell.BorderAround(LineStyleType.Thin, Color.get_Black())
            copy_table_cell_content(table.Rows[table_row_index].Cells[cell_index], cell)
            column += 1
        row += 1

    header_written[table_type] = True
    return row


def get_table_contours(image):
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, threshold = cv2.threshold(gray_image, 150, 255, cv2.THRESH_BINARY_INV)

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30))

    horizontal_lines = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, horizontal_kernel, 2)
    vertical_lines = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, vertical_kernel, 2)

    mask = cv2.add(horizontal_lines, vertical_lines)
    mask = cv2.dilate(mask, np.ones((3, 3), np.uint8), 1)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return sorted(contours, key=lambda contour: cv2.boundingRect(contour)[1])


def pdf_region_to_docx(document, page_index, rect):
    region_document = fitz.open()
    region_page = region_document.new_page(width=rect.width, height=rect.height)
    region_page.show_pdf_page(region_page.rect, document, page_index, clip=rect)

    pdf_bytes = region_document.tobytes()
    region_document.close()

    temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    temp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)

    try:
        temp_pdf.write(pdf_bytes)
        temp_pdf.close()
        temp_docx.close()

        converter = Converter(temp_pdf.name)
        converter.convert(temp_docx.name)
        converter.close()

        docx_document = Document()
        docx_document.LoadFromFile(temp_docx.name)
        return docx_document
    finally:
        with suppress(FileNotFoundError, PermissionError):
            os.remove(temp_pdf.name)
        with suppress(FileNotFoundError, PermissionError):
            os.remove(temp_docx.name)


def write_docx_tables_to_excel(document, table_type, sheets, row_tracker, workbook, header_written):
    if table_type not in sheets:
        sheets[table_type] = workbook.CreateEmptySheet(table_type)
        row_tracker[table_type] = 1

    worksheet = sheets[table_type]
    row = row_tracker[table_type]

    for section_index in range(document.Sections.Count):
        section = document.Sections[section_index]
        for child_index in range(section.Body.ChildObjects.Count):
            child_object = section.Body.ChildObjects[child_index]
            if isinstance(child_object, Table):
                row = export_table_to_excel(worksheet, row, Table(child_object), table_type, header_written)

    row_tracker[table_type] = row
    document.Dispose()


def convert_pdf_to_excel(pdf_path, output_excel_path, customerName: str):
    pattern_config = get_pattern_config(customerName)
    document = fitz.open(pdf_path)
    workbook = Workbook()
    workbook.Worksheets.Clear()

    try:
        total_pages = len(document)
        sheets = {}
        row_tracker = {}
        header_written = {}

        for page_index, page in enumerate(document):
            pixmap = page.get_pixmap(dpi=DPI)
            image = np.frombuffer(pixmap.samples, np.uint8).reshape(pixmap.height, pixmap.width, pixmap.n)
            image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

            contours = get_table_contours(image)

            page_width = page.rect.width
            image_height, image_width = image.shape[:2]

            scale_x = page_width / image_width
            scale_y = page.rect.height / image_height

            for table_name, table_config in pattern_config.items():
                if table_config["type"] == "region" and is_page_in_scope(table_config, page_index, total_pages):
                    rect = fitz.Rect(
                        table_config["x"][0] * scale_x,
                        table_config["y"][0] * scale_y,
                        table_config["x"][1] * scale_x,
                        table_config["y"][1] * scale_y,
                    )

                    docx_document = pdf_region_to_docx(document, page_index, rect)
                    write_docx_tables_to_excel(docx_document, table_name, sheets, row_tracker, workbook, header_written)

            for contour in contours:
                x, y, width, height = cv2.boundingRect(contour)

                if width < 80 or height < 30:
                    continue

                table_type = detect_table_type(x, y, page_index, total_pages, customerName)
                if not table_type:
                    continue

                rect = fitz.Rect(
                    x * scale_x,
                    y * scale_y,
                    (x + width) * scale_x,
                    (y + height) * scale_y,
                )

                docx_document = pdf_region_to_docx(document, page_index, rect)
                write_docx_tables_to_excel(docx_document, table_type, sheets, row_tracker, workbook, header_written)

        workbook.SaveToFile(output_excel_path, ExcelVersion.Version2016)
    finally:
        workbook.Dispose()
        document.close()


async def process_uploaded_pdf(file: UploadFile, customerName: str):
    pdf_bytes = await file.read()

    temp_pdf_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4()}.pdf")
    temp_excel_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4()}.xlsx")

    try:
        with open(temp_pdf_path, "wb") as pdf_file:
            pdf_file.write(pdf_bytes)

        convert_pdf_to_excel(temp_pdf_path, temp_excel_path, customerName)
        return extract_json_from_excel(temp_excel_path, customerName)
    finally:
        time.sleep(0.5)
        with suppress(FileNotFoundError, PermissionError):
            os.remove(temp_pdf_path)
        with suppress(FileNotFoundError, PermissionError):
            os.remove(temp_excel_path)

