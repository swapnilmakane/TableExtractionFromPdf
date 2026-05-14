from pdf_table_extractor.extractors.pattern_1 import extract_pattern_1, get_pattern_config_pattern_1


DPI = 300


PATTERN_CONFIG = {
    "pattern_1": get_pattern_config_pattern_1()
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

