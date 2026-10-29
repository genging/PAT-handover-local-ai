"""Local image loading and metadata extraction helpers."""

from __future__ import annotations

import hashlib
import re
from datetime import datetime
from pathlib import Path


def read_image(image_path: str):
    """Load an image with EXIF orientation applied; HEIC is supported when available."""
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise RuntimeError("Pillow is required for image loading") from exc
    path = Path(image_path)
    if path.suffix.lower() in {".heic", ".heif"}:
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except ImportError as exc:
            raise RuntimeError("pillow-heif is required for HEIC images") from exc
    return ImageOps.exif_transpose(Image.open(path))


def image_hash(image_path: str) -> tuple[str, str, str]:
    """Return (hash, source, evidence), using pHash when imagehash is installed."""
    try:
        import imagehash
        value = str(imagehash.phash(read_image(image_path)))
        return value, "PHASH", f"pHash of {Path(image_path).name}"
    except (ImportError, RuntimeError):
        digest = hashlib.sha256(Path(image_path).read_bytes()).hexdigest()[:16]
        return digest, "SHA256_FALLBACK", f"SHA-256 fallback of {Path(image_path).name}"


def extract_timestamp(image_path: str) -> tuple[str, str, str]:
    """Return an ISO timestamp from EXIF, or a timestamp-like filename token."""
    try:
        image = read_image(image_path)
        exif = image.getexif()
        raw = exif.get(36867) or exif.get(306)
        if raw:
            value = datetime.strptime(str(raw), "%Y:%m:%d %H:%M:%S").isoformat(sep=" ")
            return value, "EXIF_TIMESTAMP", f"EXIF timestamp: {raw}"
    except (OSError, ValueError, RuntimeError):
        pass
    match = re.search(r"(20\d{2}[-_]?\d{2}[-_]?\d{2})[T_ -]?(\d{2})[-:]?(\d{2})[-:]?(\d{2})?", Path(image_path).stem)
    if match:
        date = match.group(1).replace("_", "-")
        date = f"{date[:4]}-{date[4:6]}-{date[6:]}" if "-" not in date else date
        time = ":".join(part for part in match.groups()[1:] if part) or "00:00:00"
        return f"{date} {time}", "FILENAME_TIMESTAMP", f"Timestamp token in filename: {match.group(0)}"
    return "", "", ""

