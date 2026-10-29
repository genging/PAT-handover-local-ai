import csv
import tempfile
import unittest
from pathlib import Path

from engine_v2.benchmark import run_benchmark
from engine_v2.decide import classify
from engine_v2.rules import check_panorama
from engine_v2.sector import resolve_sector


class EngineV2Tests(unittest.TestCase):
    def test_classify_stub_contract(self):
        result = classify("sample.jpg")
        self.assertEqual(result["category"], "Unknown")

    def test_panorama_rule(self):
        self.assertEqual(check_panorama("watermark 90"), ("Panorama 90 Degree", "PANORAMA_RULE", "Found '90' in OCR"))
        self.assertEqual(check_panorama("angle 45"), ("", "", ""))

    def test_sector_priority_and_no_s1_default(self):
        self.assertEqual(resolve_sector("sector 3", "red", "blue", "7", 10)[0], "3")
        self.assertEqual(resolve_sector("", "", "", "", 999)[0], "Unassigned")
        self.assertEqual(resolve_sector("", "", "green", "", 10)[0], "7")
        self.assertEqual(resolve_sector("", "", "ground green", "", 10)[0], "Unassigned")

    def test_benchmark_writes_report(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            labels = root / "mock_labels.csv"
            report = root / "report.html"
            with labels.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["file_path", "expected_category", "expected_sector", "expected_panorama"])
                writer.writerows([["a.jpg", "Unknown", "", ""], ["b.jpg", "Tower", "A", "90"], ["c.jpg", "Unknown", "", ""]])
            result = run_benchmark(labels, report)
            self.assertTrue(report.is_file())
            self.assertEqual(result["total"], 3)
            self.assertIn("Confusion Matrix", report.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()

