# Huawei Site Handover Local MVP

This first MVP runs without installing third-party Python packages. It provides a responsive LAN web UI and a safe filename/folder-rule sorter. AI/Ollama vision and Excel writing remain explicit opt-in integrations.

## Start

From `D:\Huawei Handover`:

```powershell
python -m huawei_handover_app.server --host 0.0.0.0 --port 8765
```

Open on the PC: `http://127.0.0.1:8765`  
Open on iPad: `http://<PC-LAN-IP>:8765`

The app creates `huawei_handover.db` and `Sites/` at runtime. Source photos are copied by default, not moved.

## Core scripts

```powershell
python -m huawei_handover_app.step1_sorter "D:\Photos\LP89M" "D:\Huawei Handover\Sites\LP89M" --json
python -m huawei_handover_app.step2_excel_filler "D:\path\template.xlsx" --json
```

The sorter is deliberately conservative: uncertain images go to `Unknown`. It does not claim to read pixels, tape colors, watermark text, or SN without the optional AI/OCR integration.

The Excel script is read-only inspection until a separate approved writer is implemented. It requires `pywin32` for Excel COM on Windows.

## Scope of this MVP

- Dashboard, Site Setup, Photo Review, Error Queue, Activity Review, AI Chat placeholder, Excel Output placeholder, and Audit History screens
- Site/photo SQLite records
- Copy-based photo sorting with Unknown queue
- Review confirmation and audit logging
- LAN-ready touch-friendly layout

Not yet enabled: automatic vision classification, OCR, Ollama calls, semantic Excel caption matching, image placement, full Print Preview comparison, and production authentication.
