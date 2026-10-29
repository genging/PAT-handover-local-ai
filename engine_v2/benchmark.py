"""Run the engine against labeled images and generate an HTML report."""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from html import escape
from pathlib import Path

from .decide import classify


def run_benchmark(labels_path: Path, report_path: Path) -> dict:
    """Run classification for every CSV row and write an HTML report."""
    rows = []
    with labels_path.open(newline="", encoding="utf-8-sig") as handle:
        for row in csv.DictReader(handle):
            prediction = classify(row["file_path"])
            rows.append({"expected": row, "predicted": prediction})

    total = len(rows)
    category_correct = sum(r["expected"]["expected_category"] == r["predicted"]["category"] for r in rows)
    sector_correct = sum(r["expected"]["expected_sector"] == r["predicted"]["sector"] for r in rows)
    panorama_correct = sum(r["expected"]["expected_panorama"] == r["predicted"]["panorama_angle"].replace("Panorama ", "").replace(" Degree", "") for r in rows)
    categories = sorted({r["expected"]["expected_category"] for r in rows} | {r["predicted"]["category"] for r in rows})
    matrix = Counter((r["expected"]["expected_category"], r["predicted"]["category"]) for r in rows)
    errors = [r for r in rows if r["expected"]["expected_category"] != r["predicted"]["category"] or r["expected"]["expected_sector"] != r["predicted"]["sector"]]

    matrix_headers = "".join(f"<th>{escape(category)}</th>" for category in categories)
    matrix_rows = "".join(
        f"<tr><th>{escape(expected)}</th>" + "".join(f"<td>{matrix[(expected, predicted)]}</td>" for predicted in categories) + "</tr>"
        for expected in categories
    )
    error_rows = "".join(
        f"<tr><td>{escape(r['expected']['file_path'])}</td><td>{escape(r['expected']['expected_category'])}</td>"
        f"<td>{escape(r['predicted']['category'])}</td><td>{escape(r['expected']['expected_sector'])}</td>"
        f"<td>{escape(r['predicted']['sector'])}</td></tr>" for r in errors
    ) or "<tr><td colspan='5'>No errors</td></tr>"
    report = f"""<!doctype html><html><head><meta charset='utf-8'><title>Benchmark Report</title>
<style>body{{font-family:Arial,sans-serif}} table{{border-collapse:collapse;margin:1em 0}} th,td{{border:1px solid #bbb;padding:.4em}}</style></head><body>
<h1>Benchmark Report</h1><p>Total samples: {total}</p>
<h2>Overall Accuracy</h2><p>Category: {category_correct / total * 100 if total else 0:.2f}%</p><p>Sector: {sector_correct / total * 100 if total else 0:.2f}%</p><p>Panorama: {panorama_correct / total * 100 if total else 0:.2f}%</p>
<h2>Confusion Matrix (Category)</h2><table><tr><th>Expected \\ Predicted</th>{matrix_headers}</tr>{matrix_rows}</table>
<h2>Error List</h2><table><tr><th>File</th><th>Expected Category</th><th>Predicted Category</th><th>Expected Sector</th><th>Predicted Sector</th></tr>{error_rows}</table>
</body></html>"""
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(report, encoding="utf-8")
    return {"total": total, "category_accuracy": category_correct / total if total else 0.0, "sector_accuracy": sector_correct / total if total else 0.0, "panorama_accuracy": panorama_correct / total if total else 0.0, "errors": len(errors)}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--labels", type=Path, default=Path(__file__).parent / "benchmark" / "labels.csv")
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "benchmark" / "report.html")
    args = parser.parse_args()
    print(run_benchmark(args.labels, args.output))


if __name__ == "__main__":
    main()

