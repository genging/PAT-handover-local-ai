"""Local OCR helpers focused on likely watermark/document regions."""

from __future__ import annotations

from pathlib import Path


def extract_text(image_path: str) -> tuple[str, str, str]:
    """Run EasyOCR locally on corner/document regions when the package is available."""
    try:
        import easyocr
        import numpy as np
        image = np.asarray(__import__("engine_v2.ingest", fromlist=["read_image"]).read_image(image_path))
        height, width = image.shape[:2]
        regions = [image[: height // 3, : width // 2], image[: height // 3, width // 2:], image[height // 2 :, :]]
        reader = easyocr.Reader(["en"], gpu=False, verbose=False)
        parts = [text for region in regions for _, text, confidence in reader.readtext(region) if confidence >= 0.35]
        return " ".join(parts), "EASYOCR_WATERMARK_DOCUMENT", f"OCR on {len(regions)} watermark/document regions"
    except (ImportError, OSError, RuntimeError):
        return "", "", "EasyOCR unavailable or image could not be read"


def font_color(image_path: str, text_region=None) -> tuple[str, str, str]:
    """Return a coarse font color decision; region is optional for future OCR boxes."""
    del text_region
    try:
        from collections import Counter
        import numpy as np
        pixels = np.asarray(__import__("engine_v2.ingest", fromlist=["read_image"]).read_image(image_path).convert("RGB"))
        sample = pixels.reshape(-1, 3)[:: max(1, pixels.size // 3000)]
        mean = Counter("white" if int(p.mean()) > 205 else "dark" for p in sample).most_common(1)[0][0]
        return mean, "PIXEL_COLOR_SAMPLE", "Sampled pixels from image"
    except (ImportError, OSError, RuntimeError):
        return "", "", "Pixel sampling unavailable"

