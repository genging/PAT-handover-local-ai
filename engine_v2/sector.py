"""Deterministic sector resolution with explicit priority and trace evidence."""

import re

COLOR_SECTORS = {"white": "1", "ขาว": "1", "red": "2", "แดง": "2", "blue": "3", "น้ำเงิน": "3", "orange": "4", "ส้ม": "4", "yellow": "6", "เหลือง": "6", "green": "7", "เขียว": "7", "purple": "8", "ม่วง": "8"}


def _color_sector(color: str) -> str:
    lower = color.lower()
    for name, sector in COLOR_SECTORS.items():
        if name in lower:
            return sector
    return ""


def resolve_sector(ocr_text: str, font_color: str, tape_color: str, neighbor_sector: str, time_diff: float) -> tuple[str, str, str]:
    match = re.search(r"\b(?:s|sector\s*)([1-8](?:\.\d)?)\b", ocr_text, re.IGNORECASE)
    if match:
        return match.group(1), "SECTOR_OCR_RULE", f"Found '{match.group(0)}' in OCR"
    value = _color_sector(font_color)
    if value:
        return value, "SECTOR_FONT_COLOR", f"Mapped font color '{font_color}'"
    if not re.search(r"ground|สายดิน", tape_color, re.IGNORECASE):
        value = _color_sector(tape_color)
        if value:
            return value, "SECTOR_TAPE_COLOR", f"Mapped tape color '{tape_color}'"
    if time_diff <= 90 and neighbor_sector:
        return neighbor_sector, "SECTOR_TIME_CLUSTER", f"Neighbor sector within {time_diff:g}s"
    return "Unassigned", "SECTOR_FALLBACK", "No sector evidence matched"

