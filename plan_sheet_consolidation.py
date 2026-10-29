import csv
import hashlib
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

WORKBOOK = Path(r"D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK_hidden_images_removed.xlsx")
REPORT_DIR = Path(r"D:\Huawei Handover\reports")
PLAN_PATH = REPORT_DIR / "sheet_consolidation_plan.csv"
SUMMARY_PATH = REPORT_DIR / "sheet_consolidation_summary.md"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

def local(tag):
    return tag.rsplit("}", 1)[-1]

def attr(elem, name):
    for k, v in elem.attrib.items():
        if k == name or k.rsplit("}", 1)[-1] == name:
            return v
    return ""

def text_of(elem):
    return "" if elem is None else "".join(elem.itertext())

def digest(values):
    return hashlib.sha256("|".join(sorted(values)).encode("utf-8")).hexdigest()

def normalize(base, target):
    return posixpath.normpath(posixpath.join(base, target.lstrip("/")))

def read_xml(zf, path):
    return ET.fromstring(zf.read(path))

def sheet_refs(text, names, ref_regex):
    found = []
    for m in ref_regex.finditer(text or ""):
        raw = m.group(0).rstrip("!")
        if raw.startswith("'"):
            raw = raw[1:-1].replace("''", "'")
        key = raw.casefold()
        if key in names and names[key] not in found:
            found.append(names[key])
    return found

def classify_name(name):
    n = name.strip().lower()
    if n.startswith("sow"): return "SOW"
    if "dismantle" in n: return "Dismantle"
    if re.search(r"\bsn\b|ntc|sticker", n): return "SN"
    if "card" in n: return "Card"
    if "dcdu" in n: return "DCDU"
    if "breaker" in n: return "Breaker"
    if "antenna" in n or re.search(r"\bant\b", n): return "Antenna"
    if "rru" in n or "aau" in n: return "RRU/AAU"
    if "bayface" in n: return "BAYFACE"
    if "content" in n: return "Content"
    if "checklist" in n or "boq" in n: return "Checklist/BOQ"
    if "alarm" in n: return "Alarm"
    if any(x in n for x in ("tower", "map", "j-hanger", "ladder")): return "Site/Structure"
    if any(x in n for x in ("assessment", "assestment", "ehs", "deficiency", "clean", "poe")): return "Audit/Deficiency"
    return "Other"

with zipfile.ZipFile(WORKBOOK, "r") as zf:
    wb = read_xml(zf, "xl/workbook.xml")
    rels = read_xml(zf, "xl/_rels/workbook.xml.rels")
    rel_targets = {attr(r, "Id"): attr(r, "Target") for r in rels.iter() if local(r.tag) == "Relationship"}

    sheets = []
    for node in wb.iter():
        if local(node.tag) != "sheet":
            continue
        state = attr(node, "state") or "visible"
        sheets.append({"index": len(sheets), "name": attr(node, "name"), "state": state.upper(), "path": normalize("xl", rel_targets[attr(node, "id")])})
    names = {s["name"].casefold(): s["name"] for s in sheets}
    alternatives = []
    for name in sorted(names.values(), key=len, reverse=True):
        esc = re.escape(name)
        alternatives.append(r"(?:'" + esc + r"'|" + esc + r")!")
    ref_regex = re.compile("(?i)(?:" + "|".join(alternatives) + ")")

    defined_names = []
    for dn in wb.iter():
        if local(dn.tag) == "definedName":
            refers = text_of(dn)
            refs = sheet_refs(refers, names, ref_regex)
            nm = attr(dn, "name")
            defined_names.append({
                "Name": nm,
                "LocalSheetId": attr(dn, "localSheetId"),
                "RefersTo": refers,
                "ReferencedWorksheets": "; ".join(refs),
                "IsPrintArea": bool(re.search(r"(?i)Print_Area", nm)),
                "IsPrintTitles": bool(re.search(r"(?i)Print_Titles", nm)),
                "HasExternalReference": "[" in refers,
                "HasIndirectReference": bool(re.search(r"(?i)INDIRECT|ADDRESS\s*\(", refers)),
            })

    records = []
    formula_records = []
    for s in sheets:
        root = read_xml(zf, s["path"])
        cells = [x for x in root.iter() if local(x.tag) == "c"]
        formulas = []
        values = []
        styles = Counter()
        outgoing = []
        external_formula_count = 0
        indirect_count = 0
        for c in cells:
            addr = attr(c, "r")
            style = attr(c, "s")
            if style:
                styles[style] += 1
            f = next((x for x in c if local(x.tag) == "f"), None)
            if f is not None:
                formula = text_of(f)
                formulas.append(f"{addr}={formula}")
                refs = sheet_refs(formula, names, ref_regex)
                for ref in refs:
                    if ref not in outgoing:
                        outgoing.append(ref)
                if "[" in formula:
                    external_formula_count += 1
                if re.search(r"(?i)INDIRECT|ADDRESS\s*\(", formula):
                    indirect_count += 1
                formula_records.append({"Worksheet": s["name"], "Cell": addr, "Formula": formula, "ReferencedWorksheets": "; ".join(refs), "HasExternalReference": "[" in formula, "HasIndirectReference": bool(re.search(r"(?i)INDIRECT|ADDRESS\s*\(", formula))})
            else:
                is_text = next((x for x in c.iter() if local(x.tag) == "t"), None)
                value = text_of(is_text) if is_text is not None else text_of(next((x for x in c if local(x.tag) == "v"), None))
                values.append(f"{addr}={value}")
        local_names = [n for n in defined_names if n["LocalSheetId"] and int(n["LocalSheetId"]) == s["index"]]
        referenced_names = [n for n in defined_names if s["name"] in n["ReferencedWorksheets"].split("; ")]
        records.append({
            "Index": s["index"], "Worksheet": s["name"], "Visibility": s["state"], "SheetXmlPath": s["path"],
            "Dimension": next((attr(x, "ref") for x in root.iter() if local(x.tag) == "dimension"), ""),
            "CellCount": len(cells), "FormulaCount": len(formulas), "FormulaPatternHash": digest(formulas), "ValuePatternHash": digest(values),
            "StylePatternHash": digest([f"{k}:{v}" for k, v in styles.items()]),
            "HyperlinkCount": sum(1 for x in root.iter() if local(x.tag) == "hyperlink"),
            "DataValidationCount": sum(1 for x in root.iter() if local(x.tag) == "dataValidation"),
            "ConditionalFormattingCount": sum(1 for x in root.iter() if local(x.tag) == "conditionalFormatting"),
            "DrawingCount": sum(1 for x in root.iter() if local(x.tag) in {"drawing", "legacyDrawing"}),
            "TablePartCount": sum(1 for x in root.iter() if local(x.tag) == "tablePart"),
            "PivotTablePartCount": sum(1 for x in root.iter() if local(x.tag) == "pivotTablePart"),
            "LocalNameCount": len(local_names), "ReferencedNameCount": len(referenced_names),
            "PrintAreaNames": "; ".join(n["Name"] for n in referenced_names if n["IsPrintArea"]),
            "PrintTitleNames": "; ".join(n["Name"] for n in referenced_names if n["IsPrintTitles"]),
            "FormulaExternalCount": external_formula_count, "FormulaIndirectCount": indirect_count, "OutgoingWorksheets": "; ".join(outgoing),
        })

visible = [r for r in records if r["Visibility"] == "VISIBLE"]
plan_rows = []
for record in records:
    formula_sources = [f for f in formula_records if record["Worksheet"] in f["ReferencedWorksheets"].split("; ")]
    name_sources = [n for n in defined_names if record["Worksheet"] in n["ReferencedWorksheets"].split("; ")]
    exact = [v for v in visible if v["CellCount"] == record["CellCount"] and v["Dimension"] == record["Dimension"] and v["FormulaPatternHash"] == record["FormulaPatternHash"] and v["ValuePatternHash"] == record["ValuePatternHash"]]
    formula_match = [v for v in visible if v["CellCount"] == record["CellCount"] and v["Dimension"] == record["Dimension"] and v["FormulaPatternHash"] == record["FormulaPatternHash"]]
    candidate = exact[0] if exact else (formula_match[0] if formula_match else None)
    match_type = "EXACT_CONTENT_AND_FORMULA_PATTERN" if exact else ("FORMULA_PATTERN_AND_STRUCTURE_ONLY" if formula_match else "NONE")
    differences = []
    if candidate:
        for key, label in [("ValuePatternHash", "CELL_VALUES_DIFFER"), ("StylePatternHash", "STYLE_DISTRIBUTION_DIFFERS"), ("PrintAreaNames", "PRINT_AREA_NAMES_DIFFER"), ("PrintTitleNames", "PRINT_TITLE_NAMES_DIFFER"), ("HyperlinkCount", "HYPERLINK_COUNT_DIFFERS"), ("DataValidationCount", "DATA_VALIDATION_COUNT_DIFFERS"), ("ConditionalFormattingCount", "CONDITIONAL_FORMATTING_COUNT_DIFFERS"), ("DrawingCount", "DRAWING_COUNT_DIFFERS")]:
            if candidate[key] != record[key]:
                differences.append(label)
    risk = []
    if formula_sources: risk.append("INCOMING_FORMULA_REFERENCES")
    if name_sources: risk.append("DEFINED_NAME_REFERENCES")
    if record["FormulaExternalCount"]: risk.append("EXTERNAL_FORMULAS")
    if record["FormulaIndirectCount"]: risk.append("INDIRECT_OR_ADDRESS")
    if record["DataValidationCount"]: risk.append("DATA_VALIDATION")
    if record["ConditionalFormattingCount"]: risk.append("CONDITIONAL_FORMATTING")
    if record["DrawingCount"]: risk.append("DRAWINGS")
    if record["HyperlinkCount"]: risk.append("HYPERLINKS")
    hidden = record["Visibility"] != "VISIBLE"
    if not hidden:
        recommendation, confidence, reason = "KEEP", "HIGH", "Visible sheet remains the user-facing source candidate."
    elif match_type == "EXACT_CONTENT_AND_FORMULA_PATTERN" and not differences:
        recommendation, confidence, reason = "MERGE_AND_RELINK", "MEDIUM", "Exact content/formula/structure match to a Visible Sheet; relink dependencies before retirement."
    elif match_type == "FORMULA_PATTERN_AND_STRUCTURE_ONLY":
        recommendation, confidence, reason = "REVIEW_REQUIRED", "LOW", "Formula pattern and dimensions match, but values or other structures differ."
        risk.append("CONTENT_DIFFERENCE_REQUIRES_HUMAN_REVIEW")
    elif formula_sources or name_sources:
        recommendation, confidence, reason = "DELETE_AFTER_RELINK", "LOW", "Incoming references exist but no verified Visible replacement was found."
        risk.append("NO_VERIFIED_VISIBLE_REPLACEMENT")
    else:
        recommendation, confidence, reason = "REVIEW_REQUIRED", "LOW", "Hidden sheet contains content or structures without a verified Visible replacement."
    plan_rows.append({
        "SourceWorksheet": record["Worksheet"], "SourceVisibility": record["Visibility"], "SourceCategory": classify_name(record["Worksheet"]),
        "ProposedVisibleTarget": candidate["Worksheet"] if candidate else "", "MatchType": match_type,
        "SourceDimension": record["Dimension"], "TargetDimension": candidate["Dimension"] if candidate else "",
        "SourceCellCount": record["CellCount"], "TargetCellCount": candidate["CellCount"] if candidate else "",
        "SourceFormulaCount": record["FormulaCount"], "TargetFormulaCount": candidate["FormulaCount"] if candidate else "",
        "DataAndFormulaRange": record["Dimension"], "FormulaReferencesToRewrite": "; ".join(f"{f['Worksheet']}!{f['Cell']}" for f in formula_sources),
        "DefinedNamesToReview": "; ".join(n["Name"] for n in name_sources), "PrintAreasAffected": record["PrintAreaNames"], "PrintTitlesAffected": record["PrintTitleNames"],
        "HyperlinkCount": record["HyperlinkCount"], "DataValidationCount": record["DataValidationCount"], "ConditionalFormattingCount": record["ConditionalFormattingCount"],
        "DrawingCountAfterImageCleanup": record["DrawingCount"], "PivotTablePartCount": record["PivotTablePartCount"], "OutgoingWorksheetReferences": record["OutgoingWorksheets"],
        "Differences": "; ".join(dict.fromkeys(differences)), "RiskFlags": "; ".join(dict.fromkeys(risk)), "Recommendation": recommendation, "Confidence": confidence, "Reason": reason,
        "HumanReviewRequired": "YES" if hidden else "NO_FOR_THIS_PLAN",
    })

REPORT_DIR.mkdir(parents=True, exist_ok=True)
with PLAN_PATH.open("w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=list(plan_rows[0].keys()))
    writer.writeheader(); writer.writerows(plan_rows)

hidden_plan = [r for r in plan_rows if r["SourceVisibility"] != "VISIBLE"]
exact = [r for r in hidden_plan if r["MatchType"] == "EXACT_CONTENT_AND_FORMULA_PATTERN"]
low_risk = [r for r in exact if not re.search(r"INCOMING_FORMULA_REFERENCES|EXTERNAL|INDIRECT|DATA_VALIDATION|CONDITIONAL|HYPERLINK", r["RiskFlags"])]
counts = Counter(r["Recommendation"] for r in plan_rows)
lines = [
    "# Sheet Consolidation Plan", "", f"Source workbook: {WORKBOOK}",
    "Read-only Open XML analysis only. No formulas, links, values, formatting, worksheets, or external links were changed.", "",
    "## Recommendation counts", "", "| Recommendation | Sheets |", "|---|---:|",
]
for status in ("KEEP", "MERGE_AND_RELINK", "DELETE_AFTER_RELINK", "REVIEW_REQUIRED"):
    lines.append(f"| {status} | {counts[status]} |")
lines += ["", "## Lowest-risk trial candidates", ""]
if low_risk:
    lines.extend(f"- `{r['SourceWorksheet']}` -> `{r['ProposedVisibleTarget']}` ({r['MatchType']})" for r in low_risk[:20])
else:
    lines.append("- None identified by the read-only structural scan.")
lines += ["", "These are candidates for a separate test-copy experiment, not approval for deletion.", "", "## Limitations", "", "- Formula results were not recalculated.", "- External link targets were not opened or validated.", "- Pivot caches, chart series, query providers, connection credentials, and VBA runtime behavior require Excel-side validation.", "- A matching formula pattern does not prove matching business meaning; differences are retained in the CSV.", "- No data, formula, formatting, link, or worksheet was changed.", "", "Processing stops here pending approval of a specific trial group."]
SUMMARY_PATH.write_text("\n".join(lines), encoding="utf-8")
print({"TotalRows": len(plan_rows), "HiddenRows": len(hidden_plan), "VisibleRows": len(plan_rows)-len(hidden_plan), "ExactMatches": len(exact), "LowRiskTrialCandidates": len(low_risk), "PlanPath": str(PLAN_PATH), "SummaryPath": str(SUMMARY_PATH)})
