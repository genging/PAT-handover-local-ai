from __future__ import annotations

import json
import os
import re
from collections import OrderedDict
from pathlib import Path


APP_ROOT = Path(__file__).resolve().parent
ANALYSIS_PATH = APP_ROOT / "template_analysis.json"
REFERENCE_ROOT = APP_ROOT.parent / "0_Reference_Photos"
MAP_PATH = APP_ROOT / "slot_reference_map.json"

EXCLUDED_SHEET_TERMS = (
    "content", "sow", "boq", "cover", "patc", "assestment", "assessment",
    "checklist", "clear deficiency", "ehs audit", "2.3.1", "2.3.2",
)
GENERIC_TEXT = (
    "remark", "2.3 site photo", "2.3 site photos", "2.3 photo",
    "photo checklist", "site photo", "photo", "activity",
)
CAPTION_TERMS = (
    "before", "after", "ก่อน", "หลัง", "sector", "rru", "aau", "antenna",
    "ground", "grounding", "groundbar", "jumper", "ret", "bbu", "card",
    "dcdu", "breaker", "rectifier", "krone", "cable", "socket", "ladder",
    "inlet", "outlet", "label", "sticker", "serial", "sn no", "dismantle",
    "tilt", "direction", "compass", "j-hanger", "loop", "ปอกสาย", "กราว",
    "ติดตั้ง", "รื้อถอน", "สาย", "ลาเบล", "เปิดฝา", "เช็คสาย",
)


def excluded_sheet(name: str) -> bool:
    lowered = name.casefold().strip()
    return any(term in lowered for term in EXCLUDED_SHEET_TERMS)


def is_caption(text: str) -> bool:
    value = re.sub(r"\s+", " ", str(text or "")).strip()
    lowered = value.casefold()
    if not value or len(value) < 4 or len(value) > 240:
        return False
    if any(lowered == item or lowered.startswith(item + " ") for item in GENERIC_TEXT):
        return False
    if lowered == "sector":
        return False
    if lowered.startswith(("2.3 site photo", "2.3 site photos", "2.3 photo")):
        return False
    if "remark" in lowered or "specification" in lowered or "guideline" in lowered:
        return False
    if not any(term in lowered for term in CAPTION_TERMS):
        return False
    # Exclude obvious equipment/BOM lines unless they explicitly describe a photo slot.
    if re.search(r"\b\d+\s*pcs\b|\b\d+\s*sq\.?\s*mm\b|^cable tie|^clam ", lowered):
        return False
    return True


def phase_for(text: str) -> tuple[str | None, str]:
    lowered = text.casefold()
    before = bool(re.search(r"\bbefore\b|ก่อน", lowered))
    after = bool(re.search(r"\bafter\b|หลัง", lowered))
    if before and not after:
        return "Before", "explicit_before"
    if after and not before:
        return "After", "explicit_after"
    if before and after:
        return None, "conflicting_before_after"
    # These are post-installation evidence captions in the template.
    after_terms = (
        "installation", "ground", "grounding", "groundbar", "jumper", "ret",
        "label", "sticker", "serial", "socket", "ปอกสาย", "เปิดฝา", "เช็คสาย",
        "ติดตั้ง", "กราว", "ลาเบล", "สายทองแดง", "j-hanger", "ladder",
        "dismantle", "รื้อถอน", "direction", "mechanical tilt", "tilt",
        "upper clamp", "lower clamp", "connection under", "stainless cable tie", "clamp",
    )
    if any(term in lowered for term in after_terms):
        return "After", "post_installation_evidence_term"
    return None, "phase_not_determinable"


def safe_folder_name(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r'[<>:"/\\|?*]', "-", text)
    text = re.sub(r"-+", "-", text)
    text = text.rstrip(" .")
    return (text or "Unnamed Slot")[:120].rstrip(" .")


def column_number(cell: str) -> int:
    match = re.match(r"([A-Z]+)", str(cell or "").upper())
    if not match:
        return 0
    number = 0
    for letter in match.group(1):
        number = number * 26 + ord(letter) - ord("A") + 1
    return number


def row_number(cell: str) -> int:
    match = re.search(r"(\d+)$", str(cell or ""))
    return int(match.group(1)) if match else 0


def row_pair_items(items: list[dict]) -> list[dict]:
    """Join a leading Sector X cell with the next caption cell in that row."""
    by_row: dict[int, list[dict]] = {}
    for item in items:
        by_row.setdefault(row_number(item.get("cell", "")), []).append(item)
    output: list[dict] = []
    consumed: set[int] = set()
    for row_items in by_row.values():
        ordered = sorted(row_items, key=lambda item: column_number(item.get("cell", "")))
        for index, item in enumerate(ordered):
            if id(item) in consumed:
                continue
            text = re.sub(r"\s+", " ", str(item.get("text", ""))).strip()
            if re.fullmatch(r"sector\s*[1-9](?:\.\d+)?", text, re.IGNORECASE):
                following = next((candidate for candidate in ordered[index + 1:] if id(candidate) not in consumed and str(candidate.get("text", "")).strip()), None)
                if following is not None:
                    following_text = re.sub(r"\s+", " ", str(following.get("text", ""))).strip()
                    output.append({"cell": f"{item.get('cell', '')}+{following.get('cell', '')}", "text": f"{text} - {following_text}", "paired": True})
                    consumed.add(id(item)); consumed.add(id(following))
                    continue
            output.append(item)
    return output


def main() -> None:
    data = json.loads(ANALYSIS_PATH.read_text(encoding="utf-8"))
    candidates = data.get("caption_candidates_by_sheet", {})
    slots: OrderedDict[tuple[str, str], dict] = OrderedDict()
    skipped = []
    for sheet_name, items in candidates.items():
        if excluded_sheet(sheet_name):
            continue
        for item in row_pair_items(items):
            text = re.sub(r"\s+", " ", str(item.get("text", ""))).strip()
            if not is_caption(text):
                continue
            phase, phase_reason = phase_for(text)
            if phase is None:
                skipped.append({"sheet": sheet_name, "cell": item.get("cell", ""), "caption": text, "reason": phase_reason})
                continue
            key = (phase, text.casefold())
            if key not in slots:
                slots[key] = {
                    "phase": phase,
                    "caption": text,
                    "folder_name": safe_folder_name(text),
                    "phase_reason": phase_reason,
                    "sources": [],
                }
            slots[key]["sources"].append({"sheet": sheet_name, "cell": item.get("cell", ""), "text": text})

    # Prevent case-insensitive folder collisions within the same phase.
    used: dict[str, set[str]] = {"Before": set(), "After": set()}
    result = []
    for slot in slots.values():
        base = slot["folder_name"]
        name = base
        index = 2
        while name.casefold() in used[slot["phase"]]:
            name = f"{base} ({index})"
            index += 1
        used[slot["phase"]].add(name.casefold())
        folder = REFERENCE_ROOT / slot["phase"] / name
        existed = folder.exists()
        os.makedirs(folder, exist_ok=True)
        result.append({
            "phase": slot["phase"],
            "caption": slot["caption"],
            "folder_name": name,
            "folder_path": str(folder),
            "phase_reason": slot["phase_reason"],
            "source_count": len(slot["sources"]),
            "sources": slot["sources"],
            "created_now": not existed,
        })

    output = {
        "source_analysis": str(ANALYSIS_PATH),
        "reference_root": str(REFERENCE_ROOT),
        "generated_at": __import__("datetime").datetime.now().astimezone().isoformat(),
        "before_count": sum(x["phase"] == "Before" for x in result),
        "after_count": sum(x["phase"] == "After" for x in result),
        "total_count": len(result),
        "skipped_ambiguous_count": len(skipped),
        "slots": result,
        "skipped_ambiguous": skipped,
        "limitations": [
            "โครงสร้างนี้อ้างอิง caption_candidates_by_sheet ใน template_analysis.json",
            "ไม่ได้อ่าน Hidden Worksheets และไม่ได้สร้างพิกัดรูปจากชื่อชีทหรือขนาดรูป",
            "caption ที่ไม่มี phase และไม่มีคำบ่งชี้งานหลังติดตั้งถูกข้ามเพื่อไม่เดา",
        ],
    }
    MAP_PATH.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "before": output["before_count"],
        "after": output["after_count"],
        "total": output["total_count"],
        "skipped_ambiguous": output["skipped_ambiguous_count"],
        "created_now": sum(x["created_now"] for x in result),
        "sample": [f"{x['phase']}/{x['folder_name']}" for x in result[:30]],
        "map": str(MAP_PATH),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
