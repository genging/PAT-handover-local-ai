from __future__ import annotations

import json
import re
import subprocess
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
import xml.etree.ElementTree as ET


WORKBOOK = Path(r"D:\Huawei Handover\PAT DOC_Huawei Handover_template.xlsx")
APP_DIR = Path(r"D:\Huawei Handover\huawei_handover_app")
OUTPUT = APP_DIR / "template_analysis.json"
REFERENCE_ROOT = Path(r"D:\Huawei Handover\0_Reference_Photos")


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def attr(node: ET.Element, name: str) -> str:
    for key, value in node.attrib.items():
        if key == name or key.rsplit("}", 1)[-1] == name:
            return value
    return ""


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def element_text(node: ET.Element | None) -> str:
    return "" if node is None else "".join(node.itertext()).strip()


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = read_xml(zf, "xl/sharedStrings.xml")
    except KeyError:
        return []
    return [element_text(node) for node in root if local(node.tag) == "si"]


def cell_text(cell: ET.Element, strings: list[str]) -> str:
    cell_type = attr(cell, "t")
    if cell_type == "s":
        value = next((x for x in cell if local(x.tag) == "v"), None)
        try:
            return strings[int(element_text(value))]
        except (ValueError, IndexError, TypeError):
            return ""
    if cell_type == "inlineStr":
        return element_text(next((x for x in cell if local(x.tag) == "is"), None))
    value = next((x for x in cell if local(x.tag) == "v"), None)
    return element_text(value)


CAPTION_TERMS = re.compile(
    r"(?i)before|after|sector|sn\b|serial|installation|article|tilt|background|"
    r"view\s*photo|site\s*photo|activity|remark|label|dismantle|add|remove|"
    r"antenna|rru|aau|card|dcdu|breaker|cable|ground|photo"
)


def ollama_check() -> dict:
    try:
        result = subprocess.run(
            ["ollama", "show", "qwen3.5:9b"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
        )
        output = (result.stdout or "") + (result.stderr or "")
        capabilities = []
        for line in output.splitlines():
            stripped = line.strip()
            if stripped in {"vision", "projector"}:
                capabilities.append(stripped)
        return {
            "command": "ollama show qwen3.5:9b",
            "return_code": result.returncode,
            "vision_detected": "vision" in capabilities,
            "projector_detected": "projector" in capabilities or bool(re.search(r"(?i)^\s*Projector\s*$", output, re.MULTILINE)),
            "capabilities": capabilities,
            "output_excerpt": output[:4000],
        }
    except Exception as exc:
        return {"command": "ollama show qwen3.5:9b", "error": str(exc), "vision_detected": False, "projector_detected": False}


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if not WORKBOOK.exists():
        raise FileNotFoundError(WORKBOOK)
    with zipfile.ZipFile(WORKBOOK, "r") as zf:
        workbook = read_xml(zf, "xl/workbook.xml")
        relationships = read_xml(zf, "xl/_rels/workbook.xml.rels")
        targets = {attr(rel, "Id"): attr(rel, "Target") for rel in relationships if local(rel.tag) == "Relationship"}
        sheets = []
        for node in workbook.iter():
            if local(node.tag) != "sheet":
                continue
            state = attr(node, "state") or "visible"
            target = targets[attr(node, "id")].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheets.append({"name": attr(node, "name"), "state": state, "xml_path": target})

        strings = shared_strings(zf)
        captions_by_sheet: dict[str, list[dict[str, str]]] = defaultdict(list)
        all_captions: dict[str, set[str]] = defaultdict(set)
        for sheet in sheets:
            if sheet["state"] != "visible":
                continue
            root = read_xml(zf, sheet["xml_path"])
            for cell in root.iter():
                if local(cell.tag) != "c":
                    continue
                value = re.sub(r"\s+", " ", cell_text(cell, strings)).strip()
                if not value or not CAPTION_TERMS.search(value):
                    continue
                item = {"cell": attr(cell, "r"), "text": value}
                if item not in captions_by_sheet[sheet["name"]]:
                    captions_by_sheet[sheet["name"]].append(item)
                all_captions[sheet["name"]].add(value)

    category_rules = [
        ("RRU", r"(?i)\bRRU\b"), ("Antenna", r"(?i)antenna|\bant\b"), ("AAU", r"(?i)\bAAU\b"),
        ("Card", r"(?i)\bcard\b|UBB"), ("Cable", r"(?i)cable|CPRI|fiber"),
        ("Breaker", r"(?i)breaker|CB\d"), ("DCDU", r"(?i)DCDU"),
        ("Background", r"(?i)background|barground|overview|wide"),
        ("Ground_Cable", r"(?i)ground|GND"), ("SN_Label", r"(?i)\bSN\b|serial|sticker|label"),
    ]
    recommended = []
    caption_text = "\n".join(
        text for sheet in sheets if sheet["state"] == "visible" for text in all_captions[sheet["name"]]
    )
    for folder, pattern in category_rules:
        if re.search(pattern, caption_text):
            recommended.append(folder)
    for folder in recommended:
        (REFERENCE_ROOT / folder).mkdir(parents=True, exist_ok=True)

    analysis = {
        "workbook": str(WORKBOOK),
        "method": "zipfile + xml.etree.ElementTree; visible worksheet text only",
        "sheet_counts": {
            "total": len(sheets),
            "visible": sum(s["state"] == "visible" for s in sheets),
            "hidden": sum(s["state"] == "hidden" for s in sheets),
            "veryHidden": sum(s["state"] == "veryHidden" for s in sheets),
        },
        "visible_sheets": [s["name"] for s in sheets if s["state"] == "visible"],
        "hidden_sheets": [s["name"] for s in sheets if s["state"] != "visible"],
        "caption_candidates_by_sheet": dict(captions_by_sheet),
        "caption_candidates_unique": sorted({item["text"] for items in captions_by_sheet.values() for item in items}),
        "recommended_reference_folders": recommended,
        "reference_root": str(REFERENCE_ROOT),
        "ollama": ollama_check(),
    }
    OUTPUT.write_text(json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(analysis, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
