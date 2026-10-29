from __future__ import annotations

import json
import os
import posixpath
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

SOURCE_ROOT = Path(r"D:\Completed_PAT")
REFERENCE_ROOT = Path(r"D:\Huawei Handover\0_Reference_Photos")
SUMMARY_PATH = Path(r"D:\Huawei Handover\huawei_handover_app\extraction_summary.json")
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}


def xml_root(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def rel_target(source_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target)).lstrip("/")


def relationships(zf: zipfile.ZipFile, source_part: str) -> dict[str, str]:
    rel_part = posixpath.join(posixpath.dirname(source_part), "_rels", posixpath.basename(source_part) + ".rels")
    if rel_part not in zf.namelist():
        return {}
    root = xml_root(zf.read(rel_part))
    return {item.attrib.get("Id", ""): rel_target(source_part, item.attrib.get("Target", "")) for item in root.findall("{*}Relationship")}


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = xml_root(zf.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in si.iter() if node.tag.rsplit("}", 1)[-1] == "t") for si in root.findall("{*}si")]


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find("{*}v")
    if value is None:
        inline = cell.find("{*}is")
        return "".join(node.text or "" for node in inline.iter() if node.tag.rsplit("}", 1)[-1] == "t") if inline is not None else ""
    text = value.text or ""
    return strings[int(text)] if cell.attrib.get("t") == "s" and text.isdigit() and int(text) < len(strings) else text


def safe_name(text: str, limit: int = 120) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r'[<>:"/\\|?*]', "-", text).rstrip(" .")
    return (text or "Unknown Caption")[:limit].rstrip(" .")


def row_col(cell_ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref.upper())
    if not match:
        return 0, 0
    col = 0
    for letter in match.group(1):
        col = col * 26 + ord(letter) - 64
    return int(match.group(2)) - 1, col - 1


def pair_row_texts(values: list[tuple[int, str]], pic_from_col: int, pic_to_col: int) -> str | None:
    matched = [(col, text.strip()) for col, text in sorted(values, key=lambda item: item[0]) if text.strip() and pic_from_col - 2 <= col <= pic_to_col + 2]
    sector = ""
    description = ""
    for _, text in matched:
        if re.fullmatch(r"sector\s*[1-9](?:\.\d+)?", text, re.IGNORECASE):
            sector = text
        elif len(text) >= 3 and not text.casefold().startswith("remark") and not text.casefold().startswith("2.3 site photo"):
            description = text
    if sector and description:
        return f"{sector} - {description}"
    return description or sector or None


def phase_for(caption: str) -> str:
    return "Before" if re.search(r"\bbefore\b|ก่อนทำ|ก่อน", caption, re.IGNORECASE) else "After"


def find_caption_under_picture(rows: dict[int, list[tuple[int, str]]], to_row: int, from_col: int, to_col: int) -> str | None:
    for row in range(to_row, to_row + 5):
        if row not in rows:
            continue
        caption = pair_row_texts(rows[row], from_col, to_col)
        if caption and len(caption) >= 4 and not caption.casefold().startswith("2.3 site photo") and "checklist" not in caption.casefold():
            return caption
    return None


def extract_workbook(path: Path) -> tuple[int, int]:
    extracted = 0
    anchors_seen = 0
    with zipfile.ZipFile(path, "r") as zf:
        names = set(zf.namelist())
        strings = shared_strings(zf)
        workbook = xml_root(zf.read("xl/workbook.xml"))
        workbook_rels = relationships(zf, "xl/workbook.xml")
        ignored = {"content", "cover", "patc", "sow", "boq", "checklist", "assessment", "alarm", "clear deficiency", "assestment"}

        for sheet in workbook.findall(".//{*}sheet"):
            sheet_name = sheet.attrib.get("name", "")
            if sheet.attrib.get("state", "visible") != "visible" or any(term in sheet_name.casefold() for term in ignored):
                continue
            sheet_part = workbook_rels.get(sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"), "")
            if sheet_part not in names:
                continue
            sheet_root = xml_root(zf.read(sheet_part))
            rows: dict[int, list[tuple[int, str]]] = defaultdict(list)
            for cell in sheet_root.findall(".//{*}c"):
                row, col = row_col(cell.attrib.get("r", ""))
                text = cell_value(cell, strings)
                if text.strip():
                    rows[row].append((col, text))
            drawing = sheet_root.find(".//{*}drawing")
            if drawing is None:
                continue
            sheet_rels = relationships(zf, sheet_part)
            drawing_part = sheet_rels.get(drawing.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"), "")
            if drawing_part not in names:
                continue
            drawing_rels = relationships(zf, drawing_part)
            drawing_root = xml_root(zf.read(drawing_part))

            for anchor in drawing_root:
                from_node = anchor.find("{*}from")
                to_node = anchor.find("{*}to")
                if from_node is None or to_node is None:
                    continue
                from_row = int(from_node.findtext("{*}row", "0"))
                from_col = int(from_node.findtext("{*}col", "0"))
                to_row = int(to_node.findtext("{*}row", "0"))
                to_col = int(to_node.findtext("{*}col", "0"))
                if to_row - from_row < 5:
                    continue
                blip = anchor.find(".//{*}blip")
                if blip is None:
                    continue
                media_part = drawing_rels.get(blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"), "")
                if media_part not in names or Path(media_part).suffix.casefold() not in IMAGE_EXTENSIONS:
                    continue
                if zf.getinfo(media_part).file_size < 20 * 1024:
                    continue
                anchors_seen += 1
                caption = find_caption_under_picture(rows, to_row, from_col, to_col)
                if not caption:
                    continue
                phase = phase_for(caption)
                site_match = re.search(r"PAT DOC_([^_]+)_", path.name, re.IGNORECASE)
                site_name = site_match.group(1) if site_match else path.stem[:25]
                out_dir = REFERENCE_ROOT / phase / safe_name(caption)
                os.makedirs(out_dir, exist_ok=True)
                ext = Path(media_part).suffix.lower() or ".jpg"
                out_file = out_dir / f"{safe_name(site_name)}_{safe_name(sheet_name)}_R{to_row}_C{to_col}{ext}"
                counter = 2
                while out_file.exists():
                    out_file = out_dir / f"{safe_name(site_name)}_{safe_name(sheet_name)}_R{to_row}_C{to_col}_{counter}{ext}"
                    counter += 1
                with zf.open(media_part) as source, out_file.open("wb") as target:
                    shutil.copyfileobj(source, target)
                extracted += 1
    return anchors_seen, extracted


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    files = sorted(SOURCE_ROOT.glob("*.xlsx"))
    summary = {"source_root": str(SOURCE_ROOT), "reference_root": str(REFERENCE_ROOT), "files_found": len(files), "files_processed": 0, "files_failed": 0, "anchors_seen": 0, "images_extracted": 0, "file_results": [], "errors": []}
    print(f"พบไฟล์ Excel ทั้งหมด: {len(files)} ไฟล์")
    for index, path in enumerate(files, 1):
        try:
            anchors, extracted = extract_workbook(path)
            summary["files_processed"] += 1
            summary["anchors_seen"] += anchors
            summary["images_extracted"] += extracted
            summary["file_results"].append({"file": str(path), "anchors_seen": anchors, "images_extracted": extracted})
            print(f"[{index}/{len(files)}] {path.name}: พบรูป {anchors} รูป -> ดึงเข้า Reference สำเร็จ {extracted} รูป", flush=True)
        except Exception as exc:
            summary["files_failed"] += 1
            summary["errors"].append({"file": str(path), "error": f"{type(exc).__name__}: {exc}"})
            print(f"[{index}/{len(files)}] ERROR {path.name}: {exc}", flush=True)
    SUMMARY_PATH.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nเสร็จสิ้น! ดึงภาพหน้างานจริงเข้าคลัง Reference ได้ทั้งหมด: {summary['images_extracted']} รูป")


if __name__ == "__main__":
    main()
