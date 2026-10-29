"""Deterministic panorama rules."""

import re

PANORAMA_ANGLES = (0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330)


def check_panorama(ocr_text: str) -> tuple[str, str, str]:
    """Accept only a complete supported angle token from OCR text."""
    for token in re.findall(r"(?<!\d)(\d{1,3})(?!\d)", ocr_text):
        if int(token) in PANORAMA_ANGLES:
            return f"Panorama {token} Degree", "PANORAMA_RULE", f"Found '{token}' in OCR"
    return "", "", ""

