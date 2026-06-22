#!/usr/bin/env python3
"""Build a static visual QA report for sanitized Excel report workbooks.

The report inspects OpenXML only. It is designed for workbook-backed regression
fixtures and pre-delivery QA triage. It does not render Excel, recalculate
formulas, refresh data, or replace a human visual review.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "officeRel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

SEVERITY_ORDER = {"info": 0, "low": 1, "medium": 2, "high": 3}
CELL_RE = re.compile(r"^([A-Z]+)([0-9]+)$")


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def qname(local: str, namespace: str = "main") -> str:
    return f"{{{NS[namespace]}}}{local}"


def parse_xml(zf: zipfile.ZipFile, part: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(part))
    except (KeyError, ET.ParseError):
        return None


def read_relationships(zf: zipfile.ZipFile, part: str) -> dict[str, dict[str, str]]:
    root = parse_xml(zf, part)
    if root is None:
        return {}
    rels: dict[str, dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id", "")
        if rel_id:
            rels[rel_id] = {
                "type": rel.attrib.get("Type", ""),
                "target": rel.attrib.get("Target", ""),
                "targetMode": rel.attrib.get("TargetMode", ""),
            }
    return rels


def normalize_target(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_path = Path(base).parent
    parts: list[str] = []
    for part in (base_path / target).as_posix().split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        else:
            parts.append(part)
    return "/".join(parts)


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    root = parse_xml(zf, "xl/sharedStrings.xml")
    if root is None:
        return []
    values: list[str] = []
    for si in root.findall("main:si", NS):
        values.append("".join(node.text or "" for node in si.findall(".//main:t", NS)))
    return values


def col_to_number(column: str) -> int:
    value = 0
    for char in column:
        value = value * 26 + (ord(char.upper()) - ord("A") + 1)
    return value


def cell_parts(ref: str) -> tuple[str, int]:
    match = CELL_RE.match(ref)
    if not match:
        return "", 0
    return match.group(1), int(match.group(2))


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//main:t", NS))
    value = cell.find("main:v", NS)
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def workbook_sheets(zf: zipfile.ZipFile) -> tuple[list[dict[str, str]], dict[str, str]]:
    workbook_part = "xl/workbook.xml"
    workbook = parse_xml(zf, workbook_part)
    rels = read_relationships(zf, "xl/_rels/workbook.xml.rels")
    sheets: list[dict[str, str]] = []
    print_areas: dict[str, str] = {}
    if workbook is None:
        return sheets, print_areas
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rel_id = sheet.attrib.get(f"{{{NS['officeRel']}}}id", "")
        target = rels.get(rel_id, {}).get("target", "")
        sheets.append(
            {
                "name": sheet.attrib.get("name", ""),
                "state": sheet.attrib.get("state", "visible") or "visible",
                "part": normalize_target(workbook_part, target) if target else "",
            }
        )
    for defined in workbook.findall("main:definedNames/main:definedName", NS):
        if defined.attrib.get("name") != "_xlnm.Print_Area":
            continue
        try:
            sheet_index = int(defined.attrib.get("localSheetId", ""))
        except ValueError:
            continue
        if 0 <= sheet_index < len(sheets):
            print_areas[sheets[sheet_index]["name"]] = defined.text or ""
    return sheets, print_areas


def column_widths(root: ET.Element) -> dict[int, float]:
    widths: dict[int, float] = {}
    for col in root.findall("main:cols/main:col", NS):
        try:
            start = int(col.attrib.get("min", "0"))
            end = int(col.attrib.get("max", "0"))
            width = float(col.attrib.get("width", "8.43"))
        except ValueError:
            continue
        for index in range(start, end + 1):
            widths[index] = width
    return widths


def row_heights(root: ET.Element) -> dict[int, float]:
    heights: dict[int, float] = {}
    for row in root.findall("main:sheetData/main:row", NS):
        try:
            row_index = int(row.attrib.get("r", "0"))
            height = float(row.attrib.get("ht", "15"))
        except ValueError:
            continue
        heights[row_index] = height
    return heights


def sheet_drawing_chart_count(zf: zipfile.ZipFile, sheet_part: str, root: ET.Element) -> tuple[bool, int]:
    drawing = root.find("main:drawing", NS)
    if drawing is None:
        return False, 0
    rel_id = drawing.attrib.get(f"{{{NS['officeRel']}}}id", "")
    sheet_rel_part = str(Path(sheet_part).parent / "_rels" / f"{Path(sheet_part).name}.rels").replace("\\", "/")
    rels = read_relationships(zf, sheet_rel_part)
    drawing_part = normalize_target(sheet_part, rels.get(rel_id, {}).get("target", "")) if rel_id else ""
    drawing_rel_part = str(Path(drawing_part).parent / "_rels" / f"{Path(drawing_part).name}.rels").replace("\\", "/")
    drawing_rels = read_relationships(zf, drawing_rel_part)
    chart_count = sum(1 for rel in drawing_rels.values() if "chart" in rel.get("type", "").lower())
    return True, chart_count


def add_finding(findings: list[dict[str, Any]], *, code: str, severity: str, title: str, sheet: str, evidence: dict[str, Any], action: str) -> None:
    findings.append(
        {
            "code": code,
            "severity": severity,
            "title": title,
            "sheet": sheet,
            "evidence": evidence,
            "recommendedAction": action,
        }
    )


def inspect_sheet(zf: zipfile.ZipFile, sheet: dict[str, str], shared_strings: list[str], print_area: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    findings: list[dict[str, Any]] = []
    root = parse_xml(zf, sheet.get("part", ""))
    if root is None:
        return {
            "name": sheet.get("name", ""),
            "state": sheet.get("state", ""),
            "isReportSheet": False,
            "error": "sheet XML not found",
        }, findings

    name = sheet.get("name", "")
    is_visible = sheet.get("state", "visible") == "visible"
    is_report = is_visible and "report" in name.lower()
    dimension = root.find("main:dimension", NS)
    dimension_ref = dimension.attrib.get("ref", "") if dimension is not None else ""
    widths = column_widths(root)
    heights = row_heights(root)
    has_frozen_pane = any(
        pane.attrib.get("state") in {"frozen", "frozenSplit"}
        for pane in root.findall("main:sheetViews/main:sheetView/main:pane", NS)
    )
    has_drawing, chart_count = sheet_drawing_chart_count(zf, sheet.get("part", ""), root)
    merge_count = len(root.findall("main:mergeCells/main:mergeCell", NS))

    text_count = 0
    formula_count = 0
    meaningful_cell_count = 0
    error_cell_count = 0
    long_text_cells: list[dict[str, Any]] = []
    for cell in root.findall(".//main:c", NS):
        ref = cell.attrib.get("r", "")
        text = cell_text(cell, shared_strings)
        formula = cell.find("main:f", NS)
        if formula is not None:
            formula_count += 1
        if cell.attrib.get("t") == "e" or text.startswith("#"):
            error_cell_count += 1
        if text.strip() or formula is not None:
            meaningful_cell_count += 1
        if text.strip():
            text_count += 1
        column, row = cell_parts(ref)
        width = widths.get(col_to_number(column), 8.43) if column else 8.43
        height = heights.get(row, 15)
        if text.strip() and len(text) >= 45 and len(text) > width * 4:
            severity = "high" if width <= 8 and len(text) >= 80 else "medium"
            long_text_cells.append({"cell": ref, "length": len(text), "columnWidth": width, "rowHeight": height, "severity": severity})

    if is_report and meaningful_cell_count < 2 and not has_drawing:
        add_finding(
            findings,
            code="blank-report-sheet",
            severity="high",
            title="Report sheet is blank or nearly blank",
            sheet=name,
            evidence={"meaningfulCellCount": meaningful_cell_count, "dimension": dimension_ref},
            action="Confirm whether this is an intentional placeholder; otherwise remove it or populate the report surface before delivery.",
        )
    if is_report and not print_area:
        add_finding(
            findings,
            code="missing-print-area",
            severity="medium",
            title="Report sheet has no print area",
            sheet=name,
            evidence={"sheet": name},
            action="Set a print/output area when the workbook is intended for client review or export.",
        )
    if is_report and not has_frozen_pane and meaningful_cell_count >= 3:
        add_finding(
            findings,
            code="missing-frozen-pane",
            severity="low",
            title="Report sheet has no frozen header pane",
            sheet=name,
            evidence={"sheet": name},
            action="Consider freezing header rows for repeat review workflows.",
        )
    for item in long_text_cells:
        add_finding(
            findings,
            code="long-text-narrow-column",
            severity=item["severity"],
            title="Long text appears in a narrow visible cell",
            sheet=name,
            evidence=item,
            action="Increase column width, enable wrapping, reduce text length, or move the text to a note area before delivery.",
        )
    if is_report and error_cell_count:
        add_finding(
            findings,
            code="visible-report-error-cell",
            severity="high",
            title="Visible report sheet contains cached error values",
            sheet=name,
            evidence={"errorCellCount": error_cell_count},
            action="Recalculate and fix visible errors before delivery.",
        )

    return {
        "name": name,
        "state": sheet.get("state", ""),
        "isReportSheet": is_report,
        "dimension": dimension_ref,
        "meaningfulCellCount": meaningful_cell_count,
        "textCellCount": text_count,
        "formulaCount": formula_count,
        "errorCellCount": error_cell_count,
        "hasFrozenPane": has_frozen_pane,
        "hasDrawing": has_drawing,
        "chartCount": chart_count,
        "mergeCellCount": merge_count,
        "printArea": print_area,
        "longTextRiskCount": len(long_text_cells),
    }, findings


def build_report(workbook: Path) -> dict[str, Any]:
    with zipfile.ZipFile(workbook) as zf:
        shared_strings = load_shared_strings(zf)
        sheets, print_areas = workbook_sheets(zf)
        sheet_summaries: list[dict[str, Any]] = []
        findings: list[dict[str, Any]] = []
        for sheet in sheets:
            summary, sheet_findings = inspect_sheet(zf, sheet, shared_strings, print_areas.get(sheet.get("name", ""), ""))
            sheet_summaries.append(summary)
            findings.extend(sheet_findings)

    max_severity = "info"
    severity_counts = {severity: 0 for severity in SEVERITY_ORDER}
    for finding in findings:
        severity = str(finding.get("severity", "info"))
        severity_counts[severity] = severity_counts.get(severity, 0) + 1
        if SEVERITY_ORDER.get(severity, 0) > SEVERITY_ORDER.get(max_severity, 0):
            max_severity = severity
    if not findings:
        readiness = "clean"
    elif max_severity == "high":
        readiness = "blocked-for-delivery"
    elif max_severity == "medium":
        readiness = "review-required"
    else:
        readiness = "low-risk"

    report_sheets = [sheet for sheet in sheet_summaries if sheet.get("isReportSheet")]
    return {
        "generatedAt": now_iso(),
        "workbookPath": str(workbook),
        "source": "OpenXML static visual QA",
        "summary": {
            "readiness": readiness,
            "maxSeverity": max_severity,
            "sheetCount": len(sheet_summaries),
            "reportSheetCount": len(report_sheets),
            "findingCount": len(findings),
            "highFindingCount": severity_counts.get("high", 0),
            "mediumFindingCount": severity_counts.get("medium", 0),
            "lowFindingCount": severity_counts.get("low", 0),
            "blankReportSheetCount": sum(1 for sheet in report_sheets if int(sheet.get("meaningfulCellCount") or 0) < 2),
            "reportSheetWithChartCount": sum(1 for sheet in report_sheets if int(sheet.get("chartCount") or 0) > 0),
            "reportSheetWithPrintAreaCount": sum(1 for sheet in report_sheets if sheet.get("printArea")),
        },
        "sheetSummaries": sheet_summaries,
        "findings": findings,
        "limitations": [
            "Static OpenXML inspection only; this report does not render Excel pixels.",
            "Column width and text-length checks are heuristics for review, not a replacement for screenshot QA.",
            "Use Windows Excel COM or manual Excel review for final visible rendering, zoom, page breaks, and print/PDF output.",
        ],
    }


def clean_md(value: Any) -> str:
    return str(value if value is not None else "").replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    lines = [
        "# Visual QA Report",
        "",
        f"- workbook: `{report.get('workbookPath', '')}`",
        f"- readiness: **{summary.get('readiness', '')}**",
        f"- report sheets: `{summary.get('reportSheetCount', 0)}`",
        f"- findings: `{summary.get('findingCount', 0)}`",
        f"- max severity: `{summary.get('maxSeverity', '')}`",
        "",
        "## Findings",
        "",
        "| Code | Severity | Sheet | Evidence | Recommended action |",
        "|---|---:|---|---|---|",
    ]
    if report.get("findings"):
        for finding in report["findings"]:
            evidence = json.dumps(finding.get("evidence", {}), ensure_ascii=False, sort_keys=True)
            lines.append(
                "| "
                + " | ".join(
                    [
                        clean_md(finding.get("code", "")),
                        clean_md(finding.get("severity", "")),
                        clean_md(finding.get("sheet", "")),
                        clean_md(evidence),
                        clean_md(finding.get("recommendedAction", "")),
                    ]
                )
                + " |"
            )
    else:
        lines.append("| none | info | n/a | n/a | No reviewed static visual QA risks detected. |")

    lines.extend(["", "## Sheet Summary", ""])
    lines.extend(["| Sheet | Report | Cells | Formulas | Frozen | Charts | Print Area | Long Text Risks |", "|---|---:|---:|---:|---:|---:|---|---:|"])
    for sheet in report.get("sheetSummaries", []):
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_md(sheet.get("name", "")),
                    clean_md(sheet.get("isReportSheet", False)),
                    clean_md(sheet.get("meaningfulCellCount", 0)),
                    clean_md(sheet.get("formulaCount", 0)),
                    clean_md(sheet.get("hasFrozenPane", False)),
                    clean_md(sheet.get("chartCount", 0)),
                    clean_md(sheet.get("printArea", "")),
                    clean_md(sheet.get("longTextRiskCount", 0)),
                ]
            )
            + " |"
        )
    lines.extend(["", "## Limitations", ""])
    for item in report.get("limitations", []):
        lines.append(f"- {item}")
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", required=True, type=Path, help="Input .xlsx/.xlsm OpenXML workbook")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--fail-on-high-risk", action="store_true", help="Exit non-zero when high-severity findings exist")
    parser.add_argument("--fail-on-review", action="store_true", help="Exit non-zero when any finding exists")
    args = parser.parse_args()

    workbook = args.workbook.expanduser().resolve()
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 2
    try:
        report = build_report(workbook)
    except zipfile.BadZipFile:
        print(f"Not an OpenXML workbook: {workbook}", file=sys.stderr)
        return 2

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    summary = report.get("summary", {})
    print(
        "Visual QA {readiness}: reports={reports}, findings={findings}".format(
            readiness=summary.get("readiness", ""),
            reports=summary.get("reportSheetCount", 0),
            findings=summary.get("findingCount", 0),
        )
    )
    if args.fail_on_review and int(summary.get("findingCount") or 0) > 0:
        return 1
    if args.fail_on_high_risk and int(summary.get("highFindingCount") or 0) > 0:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
