from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path


def write_sow_to_excel(workbook_path: str | Path, sow_rows: list[dict], output_path: str | Path) -> dict:
    """Write saved SOW rows to a copy of the workbook, never overwriting the source."""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        return {"status": "blocked_dependency", "message": "pywin32 is not installed.", "error": str(exc)}
    source = Path(workbook_path)
    output = Path(output_path)
    if source.resolve() == output.resolve():
        raise ValueError("output_path must be a copy, not the source workbook")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        sheet = next((item for item in workbook.Worksheets if str(item.Name).casefold() == "sow"), None)
        if sheet is None:
            raise ValueError("SOW worksheet not found")
        header_row = 12
        headers = {}
        for row in range(1, 31):
            for col in range(1, 21):
                value = str(sheet.Cells(row, col).Value or "").casefold()
                if "scope" in value or "work" in value: headers.setdefault("scope", (row, col))
                if value in {"no", "no."} or "ลำดับ" in value: headers.setdefault("no", (row, col))
                if "unit" in value or "หน่วย" in value: headers.setdefault("unit", (row, col))
                if "qty" in value or "quantity" in value or "จำนวน" in value: headers.setdefault("quantity", (row, col))
                if "detail" in value or "model" in value or "รายละเอียด" in value: headers.setdefault("details", (row, col))
        if headers:
            header_row = max(row for row, _ in headers.values())
        cols = {"no": headers.get("no", (header_row, 1))[1], "scope": headers.get("scope", (header_row, 2))[1], "unit": headers.get("unit", (header_row, 3))[1], "quantity": headers.get("quantity", (header_row, 4))[1], "details": headers.get("details", (header_row, 5))[1]}
        start = header_row + 1
        for index, row in enumerate(sow_rows, start):
            sheet.Cells(index, cols["no"]).Value = row.get("line_no", index - header_row)
            sheet.Cells(index, cols["scope"]).Value = row.get("scope_name", "")
            sheet.Cells(index, cols["unit"]).Value = row.get("unit", "pcs.")
            sheet.Cells(index, cols["quantity"]).Value = row.get("quantity", 1)
            detail = row.get("details_model", "")
            sectors = row.get("sectors", "")
            sheet.Cells(index, cols["details"]).Value = f"{detail} {sectors}".strip()
        workbook.Save()
        return {"status": "ok", "output": str(output), "rows_written": len(sow_rows), "worksheet": str(sheet.Name)}
    finally:
        if workbook is not None:
            workbook.Close(False)
        excel.Quit()


def inspect_visible_sheets(workbook_path: str | Path) -> dict:
    """Inspect only visible sheets through Excel COM when pywin32 is available.

    This function never saves the workbook. The actual picture-placement operation is
    intentionally a later, separately approved step.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        return {
            "status": "blocked_dependency",
            "message": "pywin32 is not installed. Install it before running Excel COM inspection.",
            "error": str(exc),
            "workbook": str(workbook_path),
        }
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=True)
        sheets = []
        for sheet in workbook.Worksheets:
            if int(sheet.Visible) != -1:
                continue
            used = sheet.UsedRange
            sheets.append({
                "worksheet": str(sheet.Name),
                "used_range": str(used.Address(False, False)),
                "rows": int(used.Rows.Count),
                "columns": int(used.Columns.Count),
                "shapes": int(sheet.Shapes.Count),
                "print_area": str(sheet.PageSetup.PrintArea or ""),
            })
        return {"status": "ok", "workbook": str(workbook_path), "visible_sheets": sheets}
    finally:
        if workbook is not None:
            workbook.Close(False)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect visible Excel sheets without saving.")
    parser.add_argument("workbook")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = inspect_visible_sheets(args.workbook)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
