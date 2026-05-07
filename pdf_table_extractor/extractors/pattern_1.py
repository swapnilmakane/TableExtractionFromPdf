from datetime import datetime

from openpyxl import load_workbook


def extract_pattern_1(excel_path):
    workbook = load_workbook(excel_path, data_only=True)

    def get_cell_value(sheet_name, cell_reference):
        if sheet_name not in workbook:
            return None
        return workbook[sheet_name][cell_reference].value

    def to_str(value):
        if value is None:
            return None

        text = str(value)
        text = text.replace("_x000B_", "\n")
        text = text.replace("\x0b", "\n")
        return text.strip()

    def to_int(value):
        if value is None:
            return None

        try:
            cleaned_value = str(value).replace(",", "").strip()
            return int(cleaned_value)
        except (TypeError, ValueError):
            return None

    def clean_target_date(value):
        if value is None:
            return None

        text = str(value)
        if "\uff1a" in text:
            text = text.split("\uff1a")[-1].strip()

        try:
            parsed_date = datetime.fromisoformat(text)
            return parsed_date.date().isoformat()
        except ValueError:
            try:
                parsed_date = datetime.strptime(text, "%d-%m-%Y")
                return parsed_date.date().isoformat()
            except ValueError:
                return text

    result = {
        "WorkOrderNumber": "".join(filter(None, [
            to_str(get_cell_value("work_order", "B1")),
            to_str(get_cell_value("work_order", "C1")),
            to_str(get_cell_value("work_order", "D1")),
        ])),
        "CustomerOrderId": to_str(get_cell_value("customer_id", "B1")),
        "Quantity": to_int(get_cell_value("quantity", "B1")),
        "PartNumber": to_str(get_cell_value("part_details", "B2")),
        "PartName": to_str(get_cell_value("part_details", "B5")),
        "DrawingNumber": to_str(get_cell_value("part_details", "B3")),
        "DueDate": clean_target_date(get_cell_value("target_date", "C2")),
        "Operations": [],
    }

    if "operation" in workbook:
        worksheet = workbook["operation"]
        row = 2

        while True:
            operation_number = worksheet[f"A{row}"].value
            operation_name = worksheet[f"B{row}"].value

            if not operation_number and not operation_name:
                break

            result["Operations"].append({
                "OperationNumber": to_str(operation_number),
                "OperationName": to_str(operation_name),
                "WorkCenterGroup": to_str(worksheet[f"D{row}"].value),
                "IsInHouse": to_str(worksheet[f"C{row}"].value),
            })

            row += 1

    return result

