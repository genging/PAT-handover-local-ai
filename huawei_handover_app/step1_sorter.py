from __future__ import annotations

import argparse
import base64
import io
import json
import re
import shutil
import threading
import time
from difflib import SequenceMatcher
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

import requests
from PIL import Image

OLLAMA_URL = "http://127.0.0.1:11434/api/chat"
OLLAMA_MODEL = "qwen3.5:9b"
SECTOR_COLOR_MAP = {"1": "white", "2": "red", "3": "blue", "4": "orange", "6": "yellow", "7": "green", "8": "purple"}
CATEGORIES = ["RRU", "AAU", "Antenna", "Antenna_Tilt_Compass", "BBU_Card", "DCDU", "Breaker_Rectifier", "DC_Socket_Strip", "Cable_Inlet_Ladder", "Barground", "Ground_Cable", "SN_Label", "Dismantle", "Background_Site"]
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
REFERENCE_ROOT = Path(__file__).resolve().parent.parent / "0_Reference_Photos"

@dataclass
class SortResult:
    source: str; destination: str; phase: str; phase_locked: bool; category: str; sector: str; tape_color: str; action: str; sn: str; model: str; watermark_site: str; watermark_time: str; confidence: float; status: str; reason: str; ai_slot: str; ai_reason: str; watermark_text: str = ""; panorama_angle: str = ""; sn_number: str = ""; shared_destinations: list[str] | None = None

def phase_from_source(path: Path) -> tuple[str, bool, str]:
    parts = [part.casefold() for part in path.parts]
    before = any(token in part for part in parts for token in ("before", "befor", "ก่อน"))
    after = any(token in part for part in parts for token in ("after", "หลัง"))
    if before and not after: return "Before", True, "locked_from_source_folder"
    if after and not before: return "After", True, "locked_from_source_folder"
    if before and after: return "Unknown", False, "conflicting_source_phase_folders"
    return "Unknown", False, "source_phase_not_declared"

def rule_classify(path: Path) -> tuple[str, str]:
    text = f"{path.name} {' '.join(path.parts)}".casefold()
    rules = [("Antenna_Tilt_Compass", ("tilt", "compass", "direction", "mechanical tilt")), ("BBU_Card", ("bbu", "card", "ubbp", "wbbp", "upeu", "ce board")), ("Breaker_Rectifier", ("breaker", "rectifier", "krone", "led-load", "cb63", "cb32")), ("DC_Socket_Strip", ("socket", "strip", "panel", "copper", "18mm")), ("Cable_Inlet_Ladder", ("ladder", "j-hanger", "j-loop", "inlet", "outlet", "cpri", "cable", "บากสาย")), ("Barground", ("barground", "gnd bar", "master ground")), ("Ground_Cable", ("grounding kit", "ground cable", "ground", "gnd")), ("SN_Label", ("sn", "serial", "barcode", "sticker", "label", "nbtc")), ("Dismantle", ("dismantle", "remove equipment")), ("Background_Site", ("background", "site clean", "panorama", "overview", "wide view")), ("DCDU", ("dcdu",)), ("Antenna", ("antenna", "ant_", "_ant")), ("AAU", ("aau",)), ("RRU", ("rru",))]
    for category, tokens in rules:
        if any(token in text for token in tokens): return category, "filename_or_folder_rule"
    return "Unknown", "no_safe_rule_match"

def metadata_from_filename(path: Path) -> tuple[str, str]:
    text = path.stem.upper(); sector = re.search(r"(?:^|[_\- ])S([1-9])(?:$|[_\- ])", text); action = re.search(r"(?:^|[_\- ])(ADD|REMOVE)(?:$|[_\- ])", text)
    return (sector.group(1) if sector else "", action.group(1) if action else "")

def capture_timestamp(path: Path) -> float:
    """Best-effort capture time from EXIF, filename, then file mtime."""
    try:
        with Image.open(path) as image:
            value = image.getexif().get(36867) or image.getexif().get(306)
            if value:
                return time.mktime(time.strptime(str(value), "%Y:%m:%d %H:%M:%S"))
    except Exception:
        pass
    match = re.search(r"(20\d{2})[-_]?([01]\d)[-_]?([0-3]\d)[T _-]?([0-2]\d)[ .:_-]?([0-5]\d)[ .:_-]?([0-5]\d)", path.name)
    if match:
        try:
            return time.mktime(time.strptime("".join(match.groups()), "%Y%m%d%H%M%S"))
        except ValueError:
            pass
    try:
        return path.stat().st_mtime
    except OSError:
        return 0.0

MACRO_GROUPS = {
    "PANORAMA_TOWER": ("panorama", "degree", "angle", "องศา"),
    "LABEL_BARCODE": ("barcode", "serial", "sn", "nbtc", "ntc", "sticker", "label", "tag"),
    "INDOOR_CABINET": ("bbu", "card", "dcdu", "breaker", "rectifier", "krone", "led", "socket", "cabinet"),
    "ANTENNA_CLAMP_TILT": ("clamp", "tilt", "compass", "antenna", "jumper", "ret"),
    "RRU_AAU_TOWER": ("rru", "aau", "18mm", "copper", "panel", "port"),
    "GROUND_LADDER_INLET": ("ground", "gnd", "barground", "ladder", "inlet", "outlet", "hanger", "loop", "dismantle"),
}

def macro_scene_group(text: str) -> str:
    lowered = text.casefold()
    scores = {group: sum(1 for token in tokens if token in lowered) for group, tokens in MACRO_GROUPS.items()}
    group, score = max(scores.items(), key=lambda item: item[1])
    return group if score else ""

def macro_slot_matches(slot: str, group: str) -> bool:
    return not group or bool(macro_scene_group(slot) == group)

def normalize_tape_color(value: str) -> str:
    value = str(value or "").casefold().replace("-", "_").replace(" ", "_")
    aliases = {"ฟ้า": "blue", "น้ำเงิน": "blue", "สีน้ำเงิน": "blue", "ม่วง": "purple", "violet": "purple", "เขียว": "green", "เหลือง": "yellow", "ขาว": "white", "แดง": "red", "ส้ม": "orange"}
    parts = re.split(r"[,/|+]+", value)
    for part in parts:
        normalized = aliases.get(part.strip(), part.strip())
        if normalized in set(SECTOR_COLOR_MAP.values()):
            return normalized
    return aliases.get(value, value)

def normalize_abbreviations(text: str) -> str:
    """Normalize watermark abbreviations before Sector/SOW matching."""
    value = str(text or "")
    value = re.sub(r"(?<![A-Za-z])s([1-8](?:\.\d+)?)(?![A-Za-z0-9])", r"Sector \1", value, flags=re.IGNORECASE)
    value = re.sub(r"(?<![A-Za-z])c([1-9](?:\.\d+)?)(?![A-Za-z0-9])", r"Cabinet \1", value, flags=re.IGNORECASE)
    return value

def sector_from_text(text: str) -> str:
    match = re.search(r"(?i)\b(?:s|sector\s*)([1-8](?:\.\d)?)\b", normalize_abbreviations(text))
    return f"S{match.group(1)}" if match else ""

def tape_sector(tape: str, context: str, rules: dict) -> str:
    tape = normalize_tape_color(tape)
    by_color = {value: key for key, value in SECTOR_COLOR_MAP.items()}
    if tape not in by_color:
        return ""
    ground_words = ("ground", "grounding", "gnd", "lug", "sleeve", "bar ground", "barground", "สายดิน", "กราวด์")
    if any(word in context.casefold() for word in ground_words):
        return ""
    sector = by_color[tape]
    if rules.get("allowed_sectors") and sector not in rules["allowed_sectors"]:
        return ""
    return sector

def font_sector(font_color: str, context: str, rules: dict) -> str:
    """Timestamp Camera commonly uses white text; do not turn default white into S1."""
    normalized = normalize_tape_color(font_color)
    if normalized == "white":
        return ""
    return tape_sector(normalized, context, rules)

def image_payload(path: Path) -> str:
    with Image.open(path) as image:
        image = image.convert("RGB"); image.thumbnail((800, 800), Image.Resampling.LANCZOS); buffer = io.BytesIO(); image.save(buffer, format="JPEG", quality=80, optimize=True); return base64.b64encode(buffer.getvalue()).decode("ascii")

def corner_payload(path: Path) -> list[str]:
    """Create four small corner crops for local watermark OCR."""
    with Image.open(path) as source:
        image = source.convert("RGB")
        w, h = image.size
        crops = [image.crop((0, 0, int(w*.42), int(h*.30))), image.crop((int(w*.58), 0, w, int(h*.30))), image.crop((0, int(h*.70), int(w*.42), h)), image.crop((int(w*.58), int(h*.70), w, h))]
        result = []
        for crop in crops:
            crop.thumbnail((500, 300), Image.Resampling.LANCZOS)
            buf = io.BytesIO(); crop.save(buf, format="JPEG", quality=85)
            result.append(base64.b64encode(buf.getvalue()).decode("ascii"))
        return result

def extract_watermark_ocr(path: Path, timeout: int = 45) -> dict:
    try:
        payload = {"model": OLLAMA_MODEL, "messages": [{"role": "user", "content": "Read only watermark/camera text, valid Panorama angle, serial number, and watermark font color from these four corner crops. Valid Panorama angles are only 0,30,60,90,120,150,180,210,240,270,300,330; return empty for any other number. Return JSON only with keys watermark_text, panorama_angle, sn_number, watermark_font_color. Do not guess.", "images": corner_payload(path)}], "stream": False, "think": False, "format": "json", "options": {"temperature": 0.0, "num_ctx": 2048}}
        response = requests.post(OLLAMA_URL, json=payload, timeout=timeout); response.raise_for_status()
        result = response.json().get("message", {}).get("content", "{}")
        return json.loads(result) if isinstance(result, str) else (result if isinstance(result, dict) else {})
    except Exception:
        return {}

def panorama_angle(text: str) -> str:
    allowed = {"0", "30", "60", "90", "120", "150", "180", "210", "240", "270", "300", "330"}
    match = re.search(r"(?:panorama|degree|angle|องศา)\D{0,12}(0|30|60|90|120|150|180|210|240|270|300|330)\b", text, re.I)
    return match.group(1) if match and match.group(1) in allowed else ""

VALID_PANORAMA_ANGLES = {"0", "30", "60", "90", "120", "150", "180", "210", "240", "270", "300", "330"}

def valid_panorama_angle(value: str) -> str:
    value = str(value or "").strip().replace("°", "")
    return value if value in VALID_PANORAMA_ANGLES else ""

def custom_rules() -> list[dict]:
    path = Path(__file__).resolve().parent / "custom_sort_rules.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else data.get("rules", [])
    except Exception:
        return []

def apply_custom_rule(text: str, phase: str) -> tuple[str, str]:
    lowered = text.casefold()
    for rule in custom_rules():
        needle = str(rule.get("contains", "")).casefold()
        target = str(rule.get("target_category", "")).strip()
        if needle and target and needle in lowered and (not rule.get("phase") or str(rule["phase"]).casefold() == phase.casefold()):
            return target, f"custom_rule:{needle}"
    return "", ""

def canonical_slot(name: str) -> str:
    name = re.sub(r"^\s*(?:before|after)\s*[-:]?\s*", "", name, flags=re.IGNORECASE)
    name = re.sub(r"^\s*sector\s*[1-9](?:\.\d+)?\s*[-:]\s*", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\(?\s*sector\s*[1-9](?:\.\d+)?\s*\)?", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\b(?:c[1-4]|s[1-9])\b", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\b\d{1,3}\s*°", "", name)
    name = re.sub(r"\b(?:1st|2nd)\b", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip(" -:")
    return re.sub(r"\s+", " ", name).strip()

def _slot_key(name: str) -> str:
    return re.sub(r"[^a-z0-9ก-๙]+", "", canonical_slot(name).casefold())

def reference_slot_map(phase: str) -> dict[str, list[str]]:
    root = REFERENCE_ROOT / phase
    if not root.exists():
        return {}
    result: dict[str, list[str]] = {}
    for item in root.iterdir():
        if not item.is_dir() or item.name.casefold() == "unknown":
            continue
        if not any(file.is_file() and file.suffix.casefold() in IMAGE_EXTENSIONS for file in item.rglob("*")):
            continue
        result.setdefault(_slot_key(item.name), []).append(item.name)
    return result

def reference_choices(phase: str, macro: str = "") -> list[str]:
    slot_map = reference_slot_map(phase)
    ranked = Counter()
    labels = {}
    for key, values in slot_map.items():
        if macro and not any(macro_slot_matches(name, macro) for name in values):
            continue
        labels[key] = canonical_slot(values[0])
        ranked[key] = sum(1 for name in values for file in (REFERENCE_ROOT / phase / name).rglob("*") if file.is_file() and file.suffix.casefold() in IMAGE_EXTENSIONS)
    # Keep the prompt compact. Fuzzy matching still compares against every real folder.
    return [labels[key] for key, _ in ranked.most_common(30)]

def fuzzy_reference_slot(phase: str, slot_type: str, sector: str) -> tuple[str, float]:
    slots = [name for values in reference_slot_map(phase).values() for name in values]
    if not slot_type or not slots:
        return "", 0.0
    target = f"Sector {sector[1:]} - {slot_type}" if sector else slot_type
    best_name, best_score = "", 0.0
    for name in slots:
        score = max(SequenceMatcher(None, _slot_key(target), _slot_key(name)).ratio(), SequenceMatcher(None, _slot_key(slot_type), _slot_key(name)).ratio())
        if score > best_score:
            best_name, best_score = name, score
    return (best_name, best_score) if best_score >= 0.75 else ("", best_score)

def ask_ollama(path: Path, phase: str, choices: list[str], matcher_hints: list[dict] | None = None, timeout: int = 90) -> dict:
    prompt = "Read the photo only; do not choose a category or folder. Return JSON only with exactly these keys: sector, tape_color, watermark_font_color, panorama_angle, panorama_source, action, sn, model, watermark_site, watermark_time, reason, confidence. Panorama angle is valid only when exactly one of 0,30,60,90,120,150,180,210,240,270,300,330 is read from Timestamp Camera watermark or a paper sign/label via OCR; 38 and 42 are not Panorama. Set panorama_source to watermark or paper_label, otherwise empty. Sector may be S1,S2,S3,S4,S6,S7,S8 or None. S5 does not exist. Check cable tape separately from watermark font color; default white Timestamp text is not evidence of S1. Do not use DINO, filename, category name, landscape or sky to infer category. Do not guess unclear fields."
    payload = {"model": OLLAMA_MODEL, "messages": [{"role": "user", "content": prompt, "images": [image_payload(path)]}], "stream": False, "think": False, "format": "json", "options": {"temperature": 0.1, "num_ctx": 4096}}
    response = requests.post(OLLAMA_URL, json=payload, timeout=timeout); response.raise_for_status(); outer = response.json(); content = outer.get("message", {}).get("content", "{}"); result = json.loads(content) if isinstance(content, str) else content; return result if isinstance(result, dict) else {}

def normalize_ai(result: dict, choices: list[str] | None = None) -> dict:
    valid_categories = set(choices or CATEGORIES)
    category = str(result.get("slot_type") or result.get("category") or "Unknown"); category = category if category in valid_categories else category; raw_sector = str(result.get("sector") or "").upper().replace(" ", ""); sector = raw_sector if raw_sector in {"S1", "S2", "S3", "S4", "S6", "S7", "S8"} else ""; tape = normalize_tape_color(result.get("tape_color")); action = str(result.get("action") or ""); action = action if action in {"Add", "Remove"} else ""
    try: confidence = float(result.get("confidence", 0.0))
    except (TypeError, ValueError): confidence = 0.0
    return {"category": category, "sector": sector, "tape_color": tape, "watermark_font_color": normalize_tape_color(result.get("watermark_font_color")), "panorama_angle": valid_panorama_angle(result.get("panorama_angle")), "panorama_source": str(result.get("panorama_source") or ""), "action": action, "sn": str(result.get("sn") or ""), "model": str(result.get("model") or ""), "watermark_site": str(result.get("watermark_site") or ""), "watermark_time": str(result.get("watermark_time") or ""), "confidence": max(0.0, min(1.0, confidence)), "reason": str(result.get("reason") or "")}

def normalize_sow_rules(rows: list[dict] | None) -> dict:
    rows = rows or []
    allowed_sectors = set()
    actions = set()
    equipment = set()
    for row in rows:
        actions.add(str(row.get("work_type", "")).title())
        equipment.add(str(row.get("equipment", "")).upper())
        for match in re.findall(r"S([1-9])(?:\.\d+)?", str(row.get("sectors", "")).upper()):
            allowed_sectors.add(match)
        for match in re.findall(r"(?:^|[^A-Z])S([1-9])(?:\.\d+)?", str(row.get("details_model", "")).upper()):
            allowed_sectors.add(match)
    return {"allowed_sectors": allowed_sectors, "actions": actions, "equipment": equipment}

def sow_filter_reason(result: dict, phase: str, category: str, action: str, rules: dict) -> str:
    sector = result.get("sector", "")
    if rules["allowed_sectors"] and sector and sector[1:] not in rules["allowed_sectors"]:
        return "sector_not_in_sow"
    # If SOW has no S5-S8, explicitly reject those AI choices.
    if sector in {"S5", "S6", "S7", "S8"} and sector[1:] not in rules["allowed_sectors"]:
        return "sector_5_to_8_not_in_sow"
    if action == "Remove" and "Remove" in rules["actions"] and phase not in {"Before", "Unknown"} and category != "Dismantle":
        return "remove_requires_before_or_dismantle"
    if action == "Add" and "Add" in rules["actions"] and phase not in {"After", "Unknown"}:
        return "add_requires_after"
    return ""

def sort_site(input_dir: str | Path, site_dir: str | Path, copy_files: bool = True, use_ollama: bool = True, stop_event=None, sow_rules: list[dict] | None = None, progress_callback=None) -> list[SortResult]:
    input_path = Path(input_dir); site_path = Path(site_dir); output_root = site_path / "Sorted_Photos"; results: list[SortResult] = []; normalized_sow = normalize_sow_rules(sow_rules)
    sources = [source for source in input_path.rglob("*") if source.is_file() and source.suffix.casefold() in IMAGE_EXTENSIONS]
    sources.sort(key=capture_timestamp)
    total = len(sources); log_path = site_path / "sort_debug.log"
    recent_sector = ""
    recent_timestamp = None
    recent_confidence = 0.0
    with log_path.open("a", encoding="utf-8") as log:
        for index, source in enumerate(sources, 1):
            if stop_event is not None and stop_event.is_set():
                break
            if progress_callback: progress_callback(index - 1, total, source.name, "processing", "")
            phase, locked, phase_reason = phase_from_source(source); choices = reference_choices(phase) if locked else []; matcher_hints = []; rule_category, rule_reason = rule_classify(source); sector, action = metadata_from_filename(source); ai = {}; ai_slot = ""; ai_reason = "ollama_not_used"; image_time = capture_timestamp(source)
            if locked:
                try:
                    from .image_matcher import get_top_matches
                except ImportError:
                    try: from image_matcher import get_top_matches
                    except ImportError: get_top_matches = None
                if get_top_matches is not None:
                    try: matcher_hints = get_top_matches(source, phase, top_k=3)
                    except Exception as exc: ai_reason = f"matcher_error:{type(exc).__name__}"
            # Category is decided by the raw DINOv2 ranking; macro/Qwen labels never override it.
            choices = []
            ocr = extract_watermark_ocr(source) if use_ollama and locked else {}
            watermark_text = normalize_abbreviations(ocr.get("watermark_text") or "")
            ocr_angle = str(ocr.get("panorama_angle") or "")
            sn_number = str(ocr.get("sn_number") or "")
            ocr_font_color = normalize_tape_color(ocr.get("watermark_font_color") or "")
            if use_ollama:
                try:
                    raw_ai = ask_ollama(source, phase, [], [])
                    ai = normalize_ai(raw_ai, choices); ai_slot = str(raw_ai.get("slot_type") or raw_ai.get("category") or ai.get("category") or "")
                    ai_reason = "ollama_qwen3.5_vision"
                except Exception as exc:
                    ai_reason = f"ollama_error:{type(exc).__name__}"
            # Sector is trusted only from Qwen's watermark reading. Matcher never supplies Sector.
            sector = ai.get("sector", "")[1:]; action = ai.get("action") or action; final_phase = phase if locked else "Unknown"; confidence = float(ai.get("confidence", 0.0)); tape = normalize_tape_color(ai.get("tape_color", "")) or "none"
            combined_text = normalize_abbreviations(" ".join((source.name, str(source.parent), watermark_text, ai.get("watermark_site", ""), ai.get("watermark_time", ""), ai.get("reason", ""))))
            watermark_detected_angle = panorama_angle(watermark_text)
            ocr_detected_angle = valid_panorama_angle(ocr_angle)
            ai_angle_source = str(ai.get("panorama_source", "")).casefold()
            label_detected_angle = valid_panorama_angle(ai.get("panorama_angle", "")) if ai_angle_source in {"paper_label", "paper sign", "label", "sign"} else ""
            qwen_watermark_angle = valid_panorama_angle(ai.get("panorama_angle", "")) if ai_angle_source in {"watermark", "timestamp", "timestamp watermark"} else ""
            angle = watermark_detected_angle or ocr_detected_angle or qwen_watermark_angle or label_detected_angle
            angle_source = "Watermark OCR" if (watermark_detected_angle or ocr_detected_angle) else ("Watermark Qwen" if qwen_watermark_angle else ("Paper Label OCR" if label_detected_angle else ""))
            inherited = False
            # 1. Watermark/OCR Regex has priority over all other Sector sources.
            ai_sector_candidate = sector
            sector = ""
            sector_from_regex = sector_from_text(watermark_text)[1:]
            if sector_from_regex:
                sector = sector_from_regex
                reason_from_tape = "Sector_From_Regex_Watermark"
            else:
                # 2. If Regex is absent, use the watermark font color.
                inferred_font_sector = font_sector(ocr_font_color or ai.get("watermark_font_color", ""), combined_text, normalized_sow)
                if inferred_font_sector:
                    sector = inferred_font_sector
                    reason_from_tape = "Sector_From_Font_Color"
                else:
                    # 3. Finally use the tape color around the cable.
                    inferred_tape_sector = tape_sector(tape, combined_text, normalized_sow)
                    if inferred_tape_sector:
                        sector = inferred_tape_sector
                        reason_from_tape = "Sector_From_Tape_Color"
                    else:
                        reason_from_tape = ""
            if not sector and ai_sector_candidate and confidence >= 0.85 and re.search(r"\bS[1-8]\b|sector\s*[1-8]", ai.get("reason", ""), re.IGNORECASE):
                sector = ai_sector_candidate
                reason_from_tape = "Sector_From_Qwen_Evidence"
            # 4. Only inherit from a nearby high-confidence time cluster if no direct evidence exists.
            if not sector and recent_sector and recent_timestamp is not None and abs(image_time - recent_timestamp) <= 90 and recent_confidence >= 0.75:
                sector = recent_sector
                inherited = True
            expected_color = SECTOR_COLOR_MAP.get(sector, "")
            conflict = bool(expected_color and tape not in {"", "none", "unknown"} and tape != expected_color)
            # Reference-first Category pipeline: Panorama -> Custom/SOW -> DINOv2 -> Unknown.
            category = ""
            decision_source = "UNKNOWN"
            matcher_score = float(matcher_hints[0].get("similarity", 0.0)) if matcher_hints else 0.0
            rule_target, rule_reason = apply_custom_rule(combined_text, final_phase)
            if angle:
                panorama_target = next((name for values in reference_slot_map(final_phase).values() for name in values if angle in name and "panorama" in name.casefold()), "")
                category = panorama_target or f"Panorama {angle} Degree"
                decision_source = "PANORAMA_RULE"
                confidence = max(confidence, 0.80)
                reason_parts = []
                reason_parts.append(f"panorama_angle_{angle}")
            elif rule_target:
                category = rule_target
                decision_source = "CUSTOM_RULE"
                confidence = max(confidence, 0.80)
                reason_parts = [rule_reason]
            elif matcher_hints and matcher_score >= 0.55:
                category = str(matcher_hints[0].get("slot") or "")
                decision_source = "MATCHER_PRIMARY"
                confidence = max(confidence, matcher_score)
                reason_parts = [f"matcher_similarity_{matcher_score:.2f}"]
            else:
                category = "Unknown"
                decision_source = "UNKNOWN"
                reason_parts = ["matcher_similarity_below_0.55"]
            if reason_from_tape:
                reason_parts.append(reason_from_tape)
            if inherited:
                reason_parts.append("Sector_Inherited_From_Time_Cluster")
            if ai_reason.startswith("ollama_error") or ai_reason.startswith("matcher_error"):
                reason_parts.append(ai_reason)
            if not category: reason_parts.append("slot_not_found")
            if confidence < 0.55: reason_parts.append(f"confidence_below_0.55 ({confidence:.2f})")
            if conflict: reason_parts.append("sector_tape_conflict")
            sow_reason = sow_filter_reason({"sector": f"S{sector}" if sector else ""}, final_phase, category, action, normalized_sow) if sow_rules else ""
            if sow_reason: reason_parts.append(sow_reason)
            uncertain = confidence < 0.55 or conflict or not locked or not category or category == "Unknown" or bool(sow_reason)
            if uncertain: category = "Unknown"; status = "Pending"
            else: status = "Confirmed"
            reason = "; ".join(reason_parts) or phase_reason
            target_dir = output_root / final_phase / category; target_dir.mkdir(parents=True, exist_ok=True); target = target_dir / source.name
            if target.exists(): target = target_dir / f"{source.stem}_{source.stat().st_size}{source.suffix.lower()}"
            (shutil.copy2 if copy_files else shutil.move)(source, target)
            shared_targets = []
            if matcher_hints and len(matcher_hints) > 1 and abs(float(matcher_hints[0].get("similarity", 0)) - float(matcher_hints[1].get("similarity", 0))) <= 0.08 and not angle:
                for hint in matcher_hints[1:2]:
                    shared_name, shared_score = fuzzy_reference_slot(final_phase, hint.get("slot", ""), f"S{sector}" if sector else "")
                    if shared_name and shared_name != target.parent.name and shared_score >= 0.75:
                        shared_dir = output_root / final_phase / shared_name; shared_dir.mkdir(parents=True, exist_ok=True); shared_path = shared_dir / source.name
                        if shared_path.resolve() != target.resolve():
                            shutil.copy2(source, shared_path); shared_targets.append(str(shared_path))
            result = SortResult(str(source), str(target), final_phase, locked, category, sector, tape, action, ai.get("sn", ""), ai.get("model", ""), ai.get("watermark_site", ""), ai.get("watermark_time", ""), confidence, status, reason, ai_slot, ai.get("reason", ""), watermark_text, angle, sn_number, shared_targets)
            results.append(result)
            if sector and confidence >= 0.75:
                recent_sector, recent_timestamp, recent_confidence = sector, image_time, confidence
            matcher_text = f"{matcher_hints[0]['slot']} ({matcher_hints[0]['similarity']:.2f})" if matcher_hints else "None"
            color_log = f"Color: {ocr_font_color.title()} (Font) | " if ocr_font_color else ""
            angle_log = f"[Panorama Angle: {angle} | Source: {angle_source}] | " if angle else ""
            log_line = f"[{index:02d}/{total}] {source.name} | {final_phase} | {ai_slot or 'None'} | {('S'+sector) if sector else 'None'} | {tape} | {confidence:.2f} | {angle_log}{color_log}[Decision: {decision_source}] -> Category: {category} | Score: {matcher_score:.2f} | Matcher_Top1: {matcher_text} | -> {target.parent.name} | {('Shared_Copy: '+','.join(Path(x).parent.name for x in shared_targets)+' | ') if shared_targets else ''}{('REASON: '+reason) if reason else ''}"
            log.write(log_line + "\n"); log.flush()
            if progress_callback: progress_callback(index, total, source.name, "processing", log_line)
    if progress_callback: progress_callback(len(results), total, "", "done" if not stop_event or not stop_event.is_set() else "stopped", "")
    (site_path / "sort_manifest.json").write_text(json.dumps([asdict(item) for item in results], ensure_ascii=False, indent=2), encoding="utf-8"); return results

def main() -> None:
    parser = argparse.ArgumentParser(); parser.add_argument("input_dir"); parser.add_argument("site_dir"); parser.add_argument("--move", action="store_true"); parser.add_argument("--no-ollama", action="store_true"); parser.add_argument("--json", action="store_true"); args = parser.parse_args(); results = sort_site(args.input_dir, args.site_dir, copy_files=not args.move, use_ollama=not args.no_ollama); payload = [asdict(item) for item in results]; print(json.dumps(payload, ensure_ascii=False, indent=2) if args.json else f"Sorted {len(payload)} photos; Pending={sum(item.status == 'Pending' for item in results)}")

if __name__ == "__main__": main()
