"""Classification entry point for engine v2."""

from pathlib import Path

from .rules import check_panorama
from .sector import resolve_sector


def classify(image_path: str) -> dict:
    """Run deterministic filename/OCR-compatible rules until CV is wired in."""
    text = Path(image_path).stem.replace("_", " ")
    panorama, panorama_source, _ = check_panorama(text)
    sector, sector_source, _ = resolve_sector(text, "", "", "", float("inf"))
    return {
        "category": "Unknown",
        "sector": "" if sector == "Unassigned" else sector,
        "panorama_angle": panorama,
        "decision_source": "+".join(filter(None, (sector_source, panorama_source))) or "stub",
        "score": 0.0,
    }

