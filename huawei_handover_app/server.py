from __future__ import annotations

import json
import base64
import io
import mimetypes
import re
import sqlite3
import shutil
import sys
import threading
import uuid
import requests
from PIL import Image
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    # รองรับการรันแบบแพ็กเกจ: python -m huawei_handover_app.server
    from .step1_sorter import sort_site
except ImportError:
    # รองรับการเปิดไฟล์ server.py โดยตรงจากโฟลเดอร์แอป
    from step1_sorter import sort_site


ROOT = Path(__file__).resolve().parent
APP_ROOT = ROOT.parent
DATA_ROOT = APP_ROOT / "Sites"
DB_PATH = APP_ROOT / "huawei_handover.db"
STATIC_ROOT = ROOT / "static"
EMERGENCY_STOP = threading.Event()
SITE_STOP_EVENTS: dict[str, threading.Event] = {}
SORT_JOBS: dict[str, dict] = {}
SORT_JOBS_LOCK = threading.Lock()
REBUILD_STATUS = {"status": "idle", "percent": 0, "message": ""}
REBUILD_LOCK = threading.Lock()


class CombinedStopEvent:
    def __init__(self, site_event: threading.Event):
        self.site_event = site_event

    def is_set(self) -> bool:
        return EMERGENCY_STOP.is_set() or self.site_event.is_set()


def set_sort_status(site_id: str, total: int, current: int, current_file: str, status: str, error: str = "", last_log: str = "") -> None:
    percent = int((current / total) * 100) if total else 100
    with SORT_JOBS_LOCK:
        previous = SORT_JOBS.get(site_id, {})
        SORT_JOBS[site_id] = {"total": total, "current": current, "percent": percent, "current_file": current_file, "status": status, "error": error, "last_log": last_log or previous.get("last_log", "")}


def run_sort_job(site_id: str, site: sqlite3.Row, sow_rows: list[dict], site_event: threading.Event) -> None:
    try:
        set_sort_status(site_id, 0, 0, "", "processing")
        with db() as connection:
            connection.execute("UPDATE sites SET status='PROCESSING' WHERE id=?", (site_id,))
        results = sort_site(site["input_dir"], site["site_dir"], copy_files=True, use_ollama=True, stop_event=CombinedStopEvent(site_event), sow_rules=sow_rows, progress_callback=lambda current, total, filename, status, log_line="": set_sort_status(site_id, total, current, filename, status, last_log=log_line))
        stopped = site_event.is_set() or EMERGENCY_STOP.is_set()
        with db() as connection:
            connection.execute("UPDATE sites SET status=? WHERE id=?", ("STOPPED" if stopped else "SORTED", site_id))
            for result in results:
                connection.execute("INSERT INTO photos (id,site_id,source_path,sorted_path,phase,phase_locked,category,sector,tape_color,action,sn,model,watermark_site,watermark_time,confidence,status,reason,watermark_text,panorama_angle,sn_number,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (uuid.uuid4().hex, site_id, result.source, result.destination, result.phase, int(result.phase_locked), result.category, result.sector, result.tape_color, result.action, result.sn, result.model, result.watermark_site, result.watermark_time, result.confidence, result.status, result.reason, result.watermark_text, result.panorama_angle, result.sn_number, now()))
            connection.execute("INSERT INTO audit_log(site_id,event,details,created_at) VALUES (?,?,?,?)", (site_id, "PHOTOS_SORTED", json.dumps({"count": len(results), "stopped": stopped}, ensure_ascii=False), now()))
        set_sort_status(site_id, len(results), len(results), "", "done" if not stopped else "stopped")
    except Exception as exc:
        with db() as connection:
            connection.execute("UPDATE sites SET status='ERROR' WHERE id=?", (site_id,))
        set_sort_status(site_id, 0, 0, "", "error", str(exc))


def now() -> str:
    return datetime.now(timezone.utc).isoformat()


def db() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection

def nearest_existing_directory(raw: str) -> Path:
    try:
        candidate = Path(raw or "D:\\")
        while not candidate.exists() or not candidate.is_dir():
            parent = candidate.parent
            if parent == candidate:
                return Path("D:\\") if Path("D:\\").exists() else Path.cwd()
            candidate = parent
        return candidate
    except Exception:
        return Path("D:\\") if Path("D:\\").exists() else Path.cwd()

def run_rebuild() -> None:
    global REBUILD_STATUS
    with REBUILD_LOCK:
        REBUILD_STATUS = {"status": "processing", "percent": 5, "message": "กำลังโหลด DINOv2 และสแกน Reference"}
        try:
            from .image_matcher import build_reference_embeddings
        except ImportError:
            from image_matcher import build_reference_embeddings
        try:
            def update_rebuild(phase, current, total):
                global REBUILD_STATUS
                phase_base = 10 if phase == "Before" else 55
                phase_span = 45 if phase == "Before" else 45
                REBUILD_STATUS = {"status": "processing", "percent": min(99, int(phase_base + (current / max(total, 1)) * phase_span)), "message": f"กำลังสร้างดัชนี {phase} {current}/{total}"}
            result = build_reference_embeddings(progress_callback=update_rebuild)
            counts = {phase: data.get("images", 0) for phase, data in result.get("phases", {}).items()}
            REBUILD_STATUS = {"status": "done", "percent": 100, "message": f"สร้างดัชนีสำเร็จ Before={counts.get('Before', 0)}, After={counts.get('After', 0)}"}
        except Exception as exc:
            REBUILD_STATUS = {"status": "error", "percent": 100, "message": str(exc)}

def load_custom_rules() -> list[dict]:
    path = ROOT / "custom_sort_rules.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else data.get("rules", [])
    except Exception:
        return []

def save_custom_rules(rules: list[dict]) -> None:
    (ROOT / "custom_sort_rules.json").write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")

def safe_filename(value: str, fallback: str = "learned_photo.jpg") -> str:
    name = Path(str(value or fallback)).name
    name = re.sub(r"[^A-Za-z0-9ก-๙._ -]+", "_", name).strip(" .")
    return name or fallback

def chat_photo_context(site_id: str, image_name: str) -> str:
    needle = Path(image_name or "").name.casefold()
    if not needle:
        return "No matching filename supplied."
    with db() as connection:
        rows = connection.execute("SELECT source_path, sorted_path, phase, category, sector, confidence, reason, watermark_text, panorama_angle, sn_number FROM photos WHERE site_id=?", (site_id,)).fetchall()
    matches = [dict(row) for row in rows if needle in Path(str(row["source_path"] or "")).name.casefold() or needle in Path(str(row["sorted_path"] or "")).name.casefold()]
    return json.dumps(matches[:5], ensure_ascii=False) if matches else "No database photo match."


def init_db() -> None:
    DATA_ROOT.mkdir(exist_ok=True)
    with db() as connection:
        connection.executescript("""
        CREATE TABLE IF NOT EXISTS sites (
            id TEXT PRIMARY KEY, site_code TEXT NOT NULL, site_name TEXT,
            input_dir TEXT, template_path TEXT, site_dir TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'NEW', created_at TEXT NOT NULL,
            stop_requested INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS photos (
            id TEXT PRIMARY KEY, site_id TEXT NOT NULL, source_path TEXT,
            sorted_path TEXT, phase TEXT, phase_locked INTEGER DEFAULT 0,
            category TEXT, sector TEXT, tape_color TEXT, action TEXT,
            sn TEXT, model TEXT, watermark_site TEXT, watermark_time TEXT,
            confidence REAL DEFAULT 0, status TEXT, reason TEXT, watermark_text TEXT,
            panorama_angle TEXT, sn_number TEXT, created_at TEXT NOT NULL,
            FOREIGN KEY(site_id) REFERENCES sites(id)
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, site_id TEXT, event TEXT,
            details TEXT, created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS site_sow (
            id TEXT PRIMARY KEY, site_id TEXT NOT NULL, line_no INTEGER NOT NULL,
            work_type TEXT NOT NULL, equipment TEXT NOT NULL, scope_name TEXT NOT NULL,
            unit TEXT NOT NULL DEFAULT 'pcs.', quantity INTEGER NOT NULL DEFAULT 1,
            details_model TEXT, sectors TEXT, created_at TEXT NOT NULL,
            FOREIGN KEY(site_id) REFERENCES sites(id)
        );
        """)
        existing = {row[1] for row in connection.execute("PRAGMA table_info(photos)")}
        additions = {
            "phase": "TEXT", "phase_locked": "INTEGER DEFAULT 0", "tape_color": "TEXT", "sn": "TEXT",
            "model": "TEXT", "watermark_site": "TEXT", "watermark_time": "TEXT", "confidence": "REAL DEFAULT 0",
            "watermark_text": "TEXT", "panorama_angle": "TEXT", "sn_number": "TEXT",
        }
        for column, definition in additions.items():
            if column not in existing:
                connection.execute(f"ALTER TABLE photos ADD COLUMN {column} {definition}")
        site_columns = {row[1] for row in connection.execute("PRAGMA table_info(sites)")}
        if "stop_requested" not in site_columns:
            connection.execute("ALTER TABLE sites ADD COLUMN stop_requested INTEGER NOT NULL DEFAULT 0")


def json_body(handler: BaseHTTPRequestHandler) -> dict:
    length = int(handler.headers.get("Content-Length", "0"))
    raw = handler.rfile.read(length) if length else b"{}"
    return json.loads(raw.decode("utf-8"))


class Handler(BaseHTTPRequestHandler):
    server_version = "HuaweiHandoverMVP/0.1"

    def log_message(self, *_args):
        return

    def send_json(self, payload: object, status: int = 200) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path: Path) -> None:
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            self.send_file(STATIC_ROOT / "index.html")
            return
        if parsed.path.startswith("/static/"):
            self.send_file(STATIC_ROOT / Path(parsed.path).name)
            return
        if parsed.path == "/api/dashboard":
            with db() as connection:
                sites = connection.execute("SELECT COUNT(*) FROM sites").fetchone()[0]
                photos = connection.execute("SELECT COUNT(*) FROM photos").fetchone()[0]
                unknown = connection.execute("SELECT COUNT(*) FROM photos WHERE status != 'Confirmed'").fetchone()[0]
                return self.send_json({"sites": sites, "photos": photos, "pending": unknown, "jobs": 0, "emergency_stop": EMERGENCY_STOP.is_set()})
        if parsed.path == "/api/status":
            return self.send_json({"emergency_stop": EMERGENCY_STOP.is_set(), "rebuild": REBUILD_STATUS})
        if parsed.path == "/api/reference/rebuild-status":
            return self.send_json(REBUILD_STATUS)
        if parsed.path.startswith("/api/sort_status/"):
            site_id = parsed.path.split("/")[3]
            with SORT_JOBS_LOCK:
                status = dict(SORT_JOBS.get(site_id, {"total": 0, "current": 0, "percent": 0, "current_file": "", "status": "idle"}))
            return self.send_json(status)
        if parsed.path == "/api/fs/list":
            requested = parse_qs(parsed.query).get("path", ["D:\\"])[0]
            current = nearest_existing_directory(requested)
            try:
                dirs = []
                files = []
                for item in sorted(current.iterdir(), key=lambda p: (not p.is_dir(), p.name.casefold())):
                    if item.is_dir():
                        dirs.append({"name": item.name, "path": str(item)})
                    elif item.suffix.casefold() == ".xlsx":
                        files.append({"name": item.name, "path": str(item)})
            except Exception:
                fallback = Path("D:\\") if Path("D:\\").exists() else Path.cwd()
                try:
                    dirs = [{"name": item.name, "path": str(item)} for item in sorted(fallback.iterdir(), key=lambda p: (not p.is_dir(), p.name.casefold())) if item.is_dir()]
                    files = [{"name": item.name, "path": str(item)} for item in sorted(fallback.iterdir(), key=lambda p: (not p.is_dir(), p.name.casefold())) if item.is_file() and item.suffix.casefold() == ".xlsx"]
                    current = fallback
                except Exception:
                    current = fallback
                    dirs, files = [], []
            parent = str(current.parent) if current.parent != current else str(current)
            return self.send_json({"path": str(current), "parent": parent, "dirs": dirs, "files": files})
        if parsed.path == "/api/sites":
            with db() as connection:
                return self.send_json([dict(row) for row in connection.execute("SELECT * FROM sites ORDER BY created_at DESC")])
        if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/sow"):
            site_id = parsed.path.split("/")[3]
            with db() as connection:
                return self.send_json([dict(row) for row in connection.execute("SELECT * FROM site_sow WHERE site_id=? ORDER BY line_no", (site_id,))])
        if parsed.path == "/api/photos":
            site_id = parse_qs(parsed.query).get("site_id", [""])[0]
            with db() as connection:
                return self.send_json([dict(row) for row in connection.execute("SELECT * FROM photos WHERE site_id=? ORDER BY created_at DESC", (site_id,))])
        if parsed.path == "/api/photo-slot-counts":
            site_id = parse_qs(parsed.query).get("site_id", [""])[0]
            with db() as connection:
                rows = connection.execute("SELECT phase, category, COUNT(*) AS count FROM photos WHERE site_id=? GROUP BY phase, category", (site_id,)).fetchall()
            counts = {(row["phase"], row["category"]): int(row["count"]) for row in rows}
            for phase in ("Before", "After"):
                root = APP_ROOT / "0_Reference_Photos" / phase
                if root.exists():
                    for slot in root.iterdir():
                        if slot.is_dir() and slot.name.casefold() != "unknown":
                            counts.setdefault((phase, slot.name), 0)
            return self.send_json([{"phase": phase, "category": category, "count": count} for (phase, category), count in sorted(counts.items())])
        if parsed.path == "/api/errors":
            with db() as connection:
                return self.send_json([dict(row) for row in connection.execute("SELECT * FROM photos WHERE status != 'Confirmed' ORDER BY created_at DESC")])
        if parsed.path == "/api/audit":
            with db() as connection:
                return self.send_json([dict(row) for row in connection.execute("SELECT * FROM audit_log ORDER BY created_at DESC LIMIT 200")])
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            body = json_body(self)
            if parsed.path == "/api/emergency-stop":
                EMERGENCY_STOP.set()
                return self.send_json({"ok": True, "emergency_stop": True})
            if parsed.path == "/api/emergency-reset":
                EMERGENCY_STOP.clear()
                return self.send_json({"ok": True, "emergency_stop": False})
            if parsed.path == "/api/reference/rebuild":
                if REBUILD_STATUS.get("status") == "processing":
                    return self.send_json({"started": False, "status": REBUILD_STATUS}, 409)
                threading.Thread(target=run_rebuild, daemon=True).start()
                return self.send_json({"started": True}, 202)
            if parsed.path == "/api/chat":
                message = str(body.get("message", "")).strip()
                site_id = str(body.get("site_id", ""))
                image_base64 = str(body.get("image_base64", ""))
                image_name = str(body.get("image_name", ""))
                if not message:
                    return self.send_json({"error": "message is required"}, 400)
                site_context = "No site selected."
                log_text = ""
                if site_id:
                    with db() as connection:
                        site = connection.execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
                        photo_count = connection.execute("SELECT COUNT(*) FROM photos WHERE site_id=?", (site_id,)).fetchone()[0]
                    if site:
                        site_context = json.dumps({"site_code": site["site_code"], "status": site["status"], "photo_count": photo_count}, ensure_ascii=False)
                        log_path = Path(site["site_dir"]) / "sort_debug.log"
                        if log_path.exists():
                            log_text = log_path.read_text(encoding="utf-8", errors="replace")[-12000:]
                image_context = chat_photo_context(site_id, image_name) if image_name else "No image filename context."
                compact_site_context = site_context[:900]
                compact_image_context = image_context[:1400]
                compact_log_text = log_text[-3500:]
                prompt = "You are a local Huawei handover debugger. Answer Thai briefly with evidence. Return JSON only: answer, action, target_category, phase, rule_patch. For teaching an attached image use action=learn_photo and target_category/phase. For a watermark rule use rule_patch={contains,target_category,phase}. Site: " + compact_site_context + " Photo history: " + compact_image_context + " Log: " + compact_log_text + " User: " + message
                if image_base64:
                    try:
                        base64.b64decode(image_base64.split(",", 1)[-1], validate=True)
                    except Exception:
                        return self.send_json({"error": "รูปภาพที่แนบมาเป็น Base64 ไม่ถูกต้อง"}, 400)
                ollama_message = {"role": "user", "content": prompt}
                if image_base64:
                    ollama_message["images"] = [image_base64.split(",", 1)[-1]]
                ollama_payload = {"model": "qwen3.5:9b", "messages": [ollama_message], "stream": False, "think": False, "format": "json", "options": {"temperature": 0.1, "num_ctx": 4096}}
                response = requests.post("http://127.0.0.1:11434/api/chat", json=ollama_payload, timeout=90)
                if response.status_code == 400:
                    # Some Ollama builds reject a very long context together with vision JSON mode.
                    compact_prompt = "ตอบ JSON เท่านั้นด้วย keys answer, action, target_category, phase, rule_patch วิเคราะห์ภาพ Huawei นี้สั้นๆ หากผู้ใช้สั่งสอนภาพให้ action=learn_photo ห้ามเดา User: " + message[:500] + " Site: " + site_context[:400] + " Photo: " + image_context[:600] + " Log: " + log_text[-1000:]
                    compact_message = {"role": "user", "content": compact_prompt}
                    if image_base64:
                        compact_message["images"] = [image_base64.split(",", 1)[-1]]
                    response = requests.post("http://127.0.0.1:11434/api/chat", json={"model": "qwen3.5:9b", "messages": [compact_message], "stream": False, "think": False, "format": "json", "options": {"temperature": 0.1, "num_ctx": 4096}}, timeout=90)
                if not response.ok:
                    detail = response.text[:500].replace("\n", " ")
                    return self.send_json({"error": f"Ollama ตอบกลับ {response.status_code}: {detail}"}, 502)
                content = response.json().get("message", {}).get("content", "{}")
                reply = json.loads(content) if isinstance(content, str) else content
                patch_rule = reply.get("rule_patch") if isinstance(reply, dict) else None
                explicit_rule = bool(re.search(r"(กฎ|rule|ถ้าลายน้ำ|ให้แยก|route|จัดเข้าโฟลเดอร์)", message, re.IGNORECASE))
                if not explicit_rule:
                    patch_rule = None
                if isinstance(patch_rule, dict) and patch_rule.get("contains") and patch_rule.get("target_category"):
                    rules = load_custom_rules(); rules.append({"contains": str(patch_rule["contains"]), "target_category": str(patch_rule["target_category"]), "phase": str(patch_rule.get("phase", ""))}); save_custom_rules(rules)
                    reply["answer"] = str(reply.get("answer", "")) + " บันทึกกฎใหม่แล้ว"
                action = str(reply.get("action") or "") if isinstance(reply, dict) else ""
                explicit_learning = bool(re.search(r"(สอน|จำภาพ|เรียนรู้ภาพ|รูปนี้คือ|teach|learn|training|reference photo)", message, re.IGNORECASE))
                if action == "learn_photo" and not explicit_learning:
                    action = ""
                    if isinstance(reply, dict):
                        reply["action"] = None
                        reply["target_category"] = None
                        reply["phase"] = None
                        reply["answer"] = str(reply.get("answer", "")) + " ไม่ได้บันทึกภาพ เพราะข้อความนี้ไม่ได้สั่งสอนภาพ"
                learned = None
                if action == "learn_photo":
                    if not image_base64:
                        return self.send_json({"error": "learn_photo requires an attached image"}, 400)
                    phase = str(reply.get("phase") or body.get("phase") or "After").title()
                    category = str(reply.get("target_category") or "").strip()
                    if phase not in {"Before", "After"} or not category or category.casefold() == "unknown":
                        return self.send_json({"error": "invalid learning phase or target category"}, 400)
                    try:
                        raw = base64.b64decode(image_base64.split(",", 1)[-1], validate=True)
                        with Image.open(io.BytesIO(raw)) as image:
                            image.verify()
                        target_dir = APP_ROOT / "0_Reference_Photos" / phase / safe_filename(category)
                        target_dir.mkdir(parents=True, exist_ok=True)
                        target = target_dir / safe_filename(image_name, f"learned_{uuid.uuid4().hex}.jpg")
                        target.write_bytes(raw)
                        learned = {"phase": phase, "category": category, "path": str(target)}
                        threading.Thread(target=run_rebuild, daemon=True).start()
                    except Exception as exc:
                        return self.send_json({"error": f"invalid image: {exc}"}, 400)
                return self.send_json({"answer": reply.get("answer", str(reply)), "action": action or None, "rule_patch": patch_rule, "learned": learned})
            if parsed.path == "/api/sites":
                site_id = uuid.uuid4().hex
                code = str(body.get("site_code", "")).strip()
                if not code:
                    return self.send_json({"error": "site_code is required"}, 400)
                site_dir = DATA_ROOT / code
                site_dir.mkdir(parents=True, exist_ok=True)
                (site_dir / "Input" / "Before").mkdir(parents=True, exist_ok=True)
                (site_dir / "Input" / "After").mkdir(parents=True, exist_ok=True)
                with db() as connection:
                    connection.execute("INSERT INTO sites (id,site_code,site_name,input_dir,template_path,site_dir,status,created_at) VALUES (?,?,?,?,?,?,?,?)", (site_id, code, body.get("site_name", ""), body.get("input_dir", ""), body.get("template_path", ""), str(site_dir), "NEW", now()))
                    connection.execute("INSERT INTO audit_log(site_id,event,details,created_at) VALUES (?,?,?,?)", (site_id, "SITE_CREATED", json.dumps(body, ensure_ascii=False), now()))
                return self.send_json({"id": site_id, "site_dir": str(site_dir)}, 201)
            if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/stop"):
                site_id = parsed.path.split("/")[3]
                SITE_STOP_EVENTS.setdefault(site_id, threading.Event()).set()
                with db() as connection:
                    connection.execute("UPDATE sites SET stop_requested=1,status='STOP_REQUESTED' WHERE id=?", (site_id,))
                return self.send_json({"ok": True, "site_id": site_id, "stop_requested": True})
            if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/resume"):
                site_id = parsed.path.split("/")[3]
                SITE_STOP_EVENTS.setdefault(site_id, threading.Event()).clear()
                with db() as connection:
                    connection.execute("UPDATE sites SET stop_requested=0,status='READY' WHERE id=?", (site_id,))
                return self.send_json({"ok": True, "site_id": site_id, "stop_requested": False})
            if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/remove"):
                site_id = parsed.path.split("/")[3]
                with db() as connection:
                    site = connection.execute("SELECT site_dir FROM sites WHERE id=?", (site_id,)).fetchone()
                    if site is None:
                        return self.send_json({"error": "site not found"}, 404)
                    connection.execute("DELETE FROM photos WHERE site_id=?", (site_id,))
                    connection.execute("DELETE FROM site_sow WHERE site_id=?", (site_id,))
                    connection.execute("DELETE FROM audit_log WHERE site_id=?", (site_id,))
                    connection.execute("DELETE FROM sites WHERE id=?", (site_id,))
                SITE_STOP_EVENTS.pop(site_id, None)
                return self.send_json({"ok": True, "removed_from_database": True, "site_dir_preserved": site["site_dir"]})
            if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/sow"):
                site_id = parsed.path.split("/")[3]
                rows = body.get("rows", [])
                if not isinstance(rows, list):
                    return self.send_json({"error": "rows must be a list"}, 400)
                with db() as connection:
                    if connection.execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone() is None:
                        return self.send_json({"error": "site not found"}, 404)
                    connection.execute("DELETE FROM site_sow WHERE site_id=?", (site_id,))
                    for index, row in enumerate(rows, 1):
                        scope = str(row.get("scope_name", "")).strip()
                        work_type = str(row.get("work_type", "")).strip().title()
                        equipment = str(row.get("equipment", "")).strip().upper()
                        if work_type not in {"Add", "Remove"} or not scope or not equipment:
                            return self.send_json({"error": f"invalid SOW row {index}"}, 400)
                        try:
                            quantity = max(0, int(row.get("quantity", 1)))
                        except (TypeError, ValueError):
                            return self.send_json({"error": f"invalid quantity on row {index}"}, 400)
                        connection.execute("INSERT INTO site_sow (id,site_id,line_no,work_type,equipment,scope_name,unit,quantity,details_model,sectors,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)", (uuid.uuid4().hex, site_id, index, work_type, equipment, scope, str(row.get("unit", "pcs.") or "pcs."), quantity, str(row.get("details_model", "") or ""), str(row.get("sectors", "") or ""), now()))
                    connection.execute("INSERT INTO audit_log(site_id,event,details,created_at) VALUES (?,?,?,?)", (site_id, "SOW_SAVED", json.dumps({"count": len(rows)}, ensure_ascii=False), now()))
                return self.send_json({"ok": True, "count": len(rows)})
            if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/sow/fill"):
                site_id = parsed.path.split("/")[3]
                with db() as connection:
                    site = connection.execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
                    sow_rows = [dict(row) for row in connection.execute("SELECT * FROM site_sow WHERE site_id=? ORDER BY line_no", (site_id,))]
                if not site or not site["template_path"]:
                    return self.send_json({"error": "site template_path is not configured"}, 400)
                try:
                    from .step2_excel_filler import write_sow_to_excel
                except ImportError:
                    from step2_excel_filler import write_sow_to_excel
                output_path = body.get("output_path") or str(Path(site["site_dir"]) / f"{site['site_code']}_SOW_filled.xlsx")
                result = write_sow_to_excel(site["template_path"], sow_rows, output_path)
                return self.send_json(result)
            if parsed.path.startswith("/api/sites/") and parsed.path.endswith("/sort"):
                if EMERGENCY_STOP.is_set():
                    return self.send_json({"error": "Emergency stop is active. Reset it before starting a new job."}, 409)
                site_id = parsed.path.split("/")[3]
                with SORT_JOBS_LOCK:
                    if SORT_JOBS.get(site_id, {}).get("status") == "processing":
                        return self.send_json({"error": "sort job already running"}, 409)
                with db() as connection:
                    site = connection.execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
                if not site or not site["input_dir"]:
                    return self.send_json({"error": "site or input_dir not configured"}, 400)
                site_event = SITE_STOP_EVENTS.setdefault(site_id, threading.Event())
                if site["stop_requested"]:
                    site_event.set()
                with db() as connection:
                    sow_rows = [dict(row) for row in connection.execute("SELECT * FROM site_sow WHERE site_id=? ORDER BY line_no", (site_id,))]
                threading.Thread(target=run_sort_job, args=(site_id, site, sow_rows, site_event), daemon=True).start()
                return self.send_json({"started": True, "site_id": site_id}, 202)
            if parsed.path.startswith("/api/photos/") and parsed.path.endswith("/review"):
                photo_id = parsed.path.split("/")[3]
                allowed = {"phase", "category", "sector", "action", "status", "tape_color", "sn", "model"}
                values = {key: body[key] for key in allowed if key in body}
                if not values:
                    return self.send_json({"error": "no review fields supplied"}, 400)
                assignments = ", ".join(f"{key}=?" for key in values)
                with db() as connection:
                    photo = connection.execute("SELECT photos.*, sites.site_dir FROM photos JOIN sites ON sites.id=photos.site_id WHERE photos.id=?", (photo_id,)).fetchone()
                    if photo is None:
                        return self.send_json({"error": "photo not found"}, 404)
                    if any(key in values for key in ("phase", "category", "sector")):
                        phase = values.get("phase", photo["phase"]) or "Unknown"
                        category = values.get("category", photo["category"]) or "Unknown"
                        sector = values.get("sector", photo["sector"]) or ""
                        target_dir = Path(photo["site_dir"]) / "Sorted_Photos" / phase / category / (f"S{sector.replace('S','')}" if sector and not sector.startswith('S') else (sector or "Unassigned"))
                        target_dir.mkdir(parents=True, exist_ok=True)
                        old_path = Path(photo["sorted_path"])
                        new_path = target_dir / old_path.name
                        if old_path.exists() and old_path.resolve() != new_path.resolve():
                            shutil.move(str(old_path), str(new_path))
                            values["sorted_path"] = str(new_path)
                    assignments = ", ".join(f"{key}=?" for key in values)
                    connection.execute(f"UPDATE photos SET {assignments} WHERE id=?", (*values.values(), photo_id))
                    connection.execute("INSERT INTO audit_log(event,details,created_at) VALUES (?,?,?)", ("PHOTO_REVIEWED", json.dumps({"photo_id": photo_id, **values}, ensure_ascii=False), now()))
                return self.send_json({"ok": True})
            if parsed.path.startswith("/api/photos/") and parsed.path.endswith("/reference"):
                photo_id = parsed.path.split("/")[3]
                with db() as connection:
                    photo = connection.execute("SELECT photos.*, sites.site_dir FROM photos JOIN sites ON sites.id=photos.site_id WHERE photos.id=?", (photo_id,)).fetchone()
                if photo is None or not photo["sorted_path"]:
                    return self.send_json({"error": "photo not found"}, 404)
                category = str(body.get("category") or photo["category"] or "Unknown")
                reference_dir = APP_ROOT / "0_Reference_Photos" / category
                reference_dir.mkdir(parents=True, exist_ok=True)
                target = reference_dir / Path(photo["sorted_path"]).name
                shutil.copy2(photo["sorted_path"], target)
                with db() as connection:
                    connection.execute("INSERT INTO audit_log(site_id,event,details,created_at) VALUES (?,?,?,?)", (photo["site_id"], "PHOTO_SAVED_AS_REFERENCE", json.dumps({"photo_id": photo_id, "target": str(target)}, ensure_ascii=False), now()))
                return self.send_json({"ok": True, "target": str(target)})
            if parsed.path.startswith("/api/photos/") and parsed.path.endswith("/ocr_update"):
                photo_id = parsed.path.split("/")[3]
                watermark_text = str(body.get("watermark_text", ""))
                panorama_angle = str(body.get("panorama_angle", ""))
                sn_number = str(body.get("sn_number", ""))
                with db() as connection:
                    if connection.execute("SELECT 1 FROM photos WHERE id=?", (photo_id,)).fetchone() is None:
                        return self.send_json({"error": "photo not found"}, 404)
                    connection.execute("UPDATE photos SET watermark_text=?, panorama_angle=?, sn_number=? WHERE id=?", (watermark_text, panorama_angle, sn_number, photo_id))
                return self.send_json({"ok": True, "photo_id": photo_id, "watermark_text": watermark_text, "panorama_angle": panorama_angle, "sn_number": sn_number})
            if parsed.path.startswith("/api/photos/") and parsed.path.endswith("/copy"):
                photo_id = parsed.path.split("/")[3]
                category = str(body.get("category", "")).strip()
                phase = str(body.get("phase", "")).strip() or "After"
                if not category:
                    return self.send_json({"error": "category is required"}, 400)
                with db() as connection:
                    photo = connection.execute("SELECT photos.*, sites.site_dir FROM photos JOIN sites ON sites.id=photos.site_id WHERE photos.id=?", (photo_id,)).fetchone()
                if photo is None or not photo["sorted_path"] or not Path(photo["sorted_path"]).exists():
                    return self.send_json({"error": "photo not found"}, 404)
                target_dir = Path(photo["site_dir"]) / "Sorted_Photos" / phase / category; target_dir.mkdir(parents=True, exist_ok=True)
                target = target_dir / Path(photo["sorted_path"]).name; shutil.copy2(photo["sorted_path"], target)
                new_id = uuid.uuid4().hex
                new_record = {
                    "id": new_id, "site_id": photo["site_id"], "source_path": photo["source_path"],
                    "sorted_path": str(target), "phase": phase, "phase_locked": photo["phase_locked"],
                    "category": category, "sector": photo["sector"], "tape_color": photo["tape_color"],
                    "action": photo["action"], "sn": photo["sn"], "model": photo["model"],
                    "watermark_site": photo["watermark_site"], "watermark_time": photo["watermark_time"],
                    "confidence": photo["confidence"], "status": photo["status"], "reason": "shared_copy",
                    "watermark_text": photo["watermark_text"], "panorama_angle": photo["panorama_angle"],
                    "sn_number": photo["sn_number"], "created_at": now()
                }
                with db() as connection:
                    connection.execute("INSERT INTO photos (id,site_id,source_path,sorted_path,phase,phase_locked,category,sector,tape_color,action,sn,model,watermark_site,watermark_time,confidence,status,reason,watermark_text,panorama_angle,sn_number,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tuple(new_record[key] for key in ("id", "site_id", "source_path", "sorted_path", "phase", "phase_locked", "category", "sector", "tape_color", "action", "sn", "model", "watermark_site", "watermark_time", "confidence", "status", "reason", "watermark_text", "panorama_angle", "sn_number", "created_at")))
                    connection.execute("INSERT INTO audit_log(site_id,event,details,created_at) VALUES (?,?,?,?)", (photo["site_id"], "PHOTO_COPIED_TO_SLOT", json.dumps({"photo_id": photo_id, "target": str(target), "new_photo_id": new_id}, ensure_ascii=False), now()))
                return self.send_json({"ok": True, "target": str(target), "new_photo_record": new_record})
            self.send_error(HTTPStatus.NOT_FOUND)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="Local Huawei Handover Web MVP")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    init_db()
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Huawei Handover MVP: http://127.0.0.1:{args.port}")
    print(f"LAN access: http://<PC-IP>:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
