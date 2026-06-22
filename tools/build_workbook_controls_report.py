#!/usr/bin/env python3
"""Build a static workbook controls and visibility report from OpenXML JSON.

Input is the JSON produced by ``tools/inspect_excel_bi_workbook.py``. The report
does not open Excel, unprotect sheets, reveal hidden content, or modify the
workbook. It highlights delivery-surface states that usually need to be
intentional before handoff: hidden sheets, very hidden sheets, workbook/sheet
protection, active filters, frozen panes, and data validation rules.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


SEVERITY_ORDER = {"info": 0, "low": 1, "medium": 2, "high": 3}


def as_list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def add_finding(
    findings: list[dict[str, Any]],
    code: str,
    severity: str,
    title: str,
    evidence: dict[str, Any],
    action: str,
) -> None:
    findings.append(
        {
            "code": code,
            "severity": severity,
            "title": title,
            "evidence": evidence,
            "recommendedAction": action,
        }
    )


def build_report(openxml_report: dict[str, Any]) -> dict[str, Any]:
    findings: list[dict[str, Any]] = []
    sheet_controls = [item for item in as_list(openxml_report.get("sheetControls")) if isinstance(item, dict)]
    hidden_sheets = [item for item in sheet_controls if str(item.get("state", "visible")) == "hidden"]
    very_hidden_sheets = [item for item in sheet_controls if str(item.get("state", "visible")) == "veryHidden"]
    protected_sheets = [item for item in sheet_controls if item.get("hasSheetProtection")]
    filtered_sheets = [item for item in sheet_controls if item.get("hasAutoFilter")]
    frozen_sheets = [item for item in sheet_controls if item.get("hasFrozenPane")]
    validation_sheets = [item for item in sheet_controls if int(item.get("dataValidationCount") or 0) > 0]

    if openxml_report.get("hasWorkbookProtection"):
        add_finding(
            findings,
            "workbook-protection",
            "medium",
            "Workbook structure protection is enabled",
            {"workbookProtection": openxml_report.get("workbookProtection", {})},
            "Confirm the protection is intentional and document whether users need a password or unlocked delivery copy.",
        )

    if hidden_sheets:
        add_finding(
            findings,
            "hidden-sheets",
            "low",
            "Workbook contains hidden worksheets",
            {"sheets": [item.get("name", "") for item in hidden_sheets]},
            "Confirm hidden support sheets are intentional and no required user-facing output is hidden.",
        )

    if very_hidden_sheets:
        add_finding(
            findings,
            "very-hidden-sheets",
            "medium",
            "Workbook contains very hidden worksheets",
            {"sheets": [item.get("name", "") for item in very_hidden_sheets]},
            "Review very hidden sheets in the VBA editor or OpenXML before delivery; they are not visible through normal Excel UI.",
        )

    if protected_sheets:
        add_finding(
            findings,
            "sheet-protection",
            "medium",
            "One or more worksheets are protected",
            {"sheets": [item.get("name", "") for item in protected_sheets]},
            "Verify locked/unlocked cells, allowed actions, and password expectations before handoff.",
        )

    if filtered_sheets:
        add_finding(
            findings,
            "active-auto-filter",
            "low",
            "One or more worksheets contain AutoFilter settings",
            {
                "sheets": [
                    {"name": item.get("name", ""), "autoFilterRef": item.get("autoFilterRef", "")}
                    for item in filtered_sheets
                ]
            },
            "Confirm filters do not hide required rows in the delivered view, especially before value-freezing outputs.",
        )

    if validation_sheets:
        add_finding(
            findings,
            "data-validation-rules",
            "low",
            "One or more worksheets contain data validation rules",
            {
                "sheets": [
                    {"name": item.get("name", ""), "dataValidationCount": item.get("dataValidationCount", 0)}
                    for item in validation_sheets
                ]
            },
            "Confirm validation rules still target the intended input ranges after sheet/column edits.",
        )

    max_severity = "info"
    for finding in findings:
        severity = str(finding.get("severity", "info"))
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

    severity_counts = {severity: 0 for severity in SEVERITY_ORDER}
    for finding in findings:
        severity = str(finding.get("severity", "info"))
        severity_counts[severity] = severity_counts.get(severity, 0) + 1

    summary = {
        "readiness": readiness,
        "maxSeverity": max_severity,
        "sheetCount": len(sheet_controls),
        "findingCount": len(findings),
        "mediumFindingCount": severity_counts.get("medium", 0),
        "lowFindingCount": severity_counts.get("low", 0),
        "hiddenSheetCount": len(hidden_sheets),
        "veryHiddenSheetCount": len(very_hidden_sheets),
        "protectedSheetCount": len(protected_sheets),
        "filteredSheetCount": len(filtered_sheets),
        "frozenPaneSheetCount": len(frozen_sheets),
        "dataValidationSheetCount": len(validation_sheets),
        "hasWorkbookProtection": bool(openxml_report.get("hasWorkbookProtection")),
    }
    return {
        "workbookPath": openxml_report.get("workbookPath", ""),
        "sourceInspector": "tools/inspect_excel_bi_workbook.py",
        "summary": summary,
        "findings": findings,
        "sheetControls": sheet_controls,
        "limitations": [
            "Static OpenXML inspection only; the report does not open Excel or test passwords.",
            "Hidden or protected sheets can be intentional; findings mean review is required, not automatic failure.",
            "Use Excel desktop validation for final user interaction, protection behavior, and visible-layout checks.",
        ],
    }


def render_markdown(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    lines = [
        "# Workbook Controls Report",
        "",
        f"- workbook: `{report.get('workbookPath', '')}`",
        f"- readiness: **{summary.get('readiness', '')}**",
        f"- max severity: `{summary.get('maxSeverity', '')}`",
        f"- sheets: `{summary.get('sheetCount', 0)}`",
        f"- findings: `{summary.get('findingCount', 0)}`",
        "",
        "| Code | Severity | Evidence | Recommended action |",
        "|---|---:|---|---|",
    ]
    for finding in report.get("findings", []):
        evidence = json.dumps(finding.get("evidence", {}), ensure_ascii=False, sort_keys=True)
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(finding.get("code", "")),
                    clean_markdown(finding.get("severity", "")),
                    clean_markdown(evidence),
                    clean_markdown(finding.get("recommendedAction", "")),
                ]
            )
            + " |"
        )
    if not report.get("findings"):
        lines.append("| none | info | n/a | no reviewed workbook control risks detected | No action required by this static check. |")
    lines.extend(["", "## Sheet Controls", ""])
    lines.extend(["| Sheet | State | Protected | AutoFilter | Frozen Pane | Data Validations |", "|---|---|---:|---|---:|---:|"])
    for sheet in report.get("sheetControls", []):
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(sheet.get("name", "")),
                    clean_markdown(sheet.get("state", "")),
                    clean_markdown(sheet.get("hasSheetProtection", False)),
                    clean_markdown(sheet.get("autoFilterRef", "")),
                    clean_markdown(sheet.get("hasFrozenPane", False)),
                    clean_markdown(sheet.get("dataValidationCount", 0)),
                ]
            )
            + " |"
        )
    lines.extend(["", "## Limitations", ""])
    for item in report.get("limitations", []):
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--openxml-json", required=True, type=Path, help="JSON from inspect_excel_bi_workbook.py")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--fail-on-review", action="store_true", help="Exit with code 1 when any review finding exists")
    args = parser.parse_args()

    openxml_report = json.loads(args.openxml_json.expanduser().read_text(encoding="utf-8-sig"))
    report = build_report(openxml_report)
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(render_markdown(report), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    summary = report.get("summary", {})
    print(
        "Workbook controls {readiness}: sheets={sheets}, findings={findings}".format(
            readiness=summary.get("readiness", ""),
            sheets=summary.get("sheetCount", 0),
            findings=summary.get("findingCount", 0),
        )
    )
    if args.fail_on_review and int(summary.get("findingCount") or 0) > 0:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
