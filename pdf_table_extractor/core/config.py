from pdf_table_extractor.extractors.pattern_1 import extract_pattern_1


DPI = 300


PATTERN_CONFIG = {
    "pattern_1": {
        "part_details": {"type": "table", "x": (115, 125), "y": (405, 420), "pages": [0]},
        "customer_id": {"type": "table", "x": (1780, 1790), "y": (405, 420), "pages": [0]},
        "quantity": {"type": "table", "x": (2430, 2445), "y": (760, 770), "pages": [0]},
        "operation": {"type": "table", "x": (115, 125), "y": (850, 860), "pages": None},
        "target_date": {"type": "table", "x": (115, 125), "y": (2060, 2070), "pages": [0]},
        "work_order": {"type": "region", "x": (120, 1750), "y": (270, 400), "pages": [0]},
    }
}

EXTRACTOR_CONFIG = {
    "pattern_1": extract_pattern_1,
}


def get_pattern_config(pattern_name):
    try:
        return PATTERN_CONFIG[pattern_name]
    except KeyError as exc:
        raise ValueError(f"Unsupported pattern: {pattern_name}") from exc


def get_pattern_extractor(pattern_name):
    try:
        return EXTRACTOR_CONFIG[pattern_name]
    except KeyError as exc:
        raise ValueError(f"Unsupported pattern: {pattern_name}") from exc

