#!/usr/bin/env python3
"""Build a first-pass triage report for an Excel BI workbook.

The report aggregates the OpenXML workbook inventory and any specialized static
reports already produced by this package. It is intentionally read-only: it does
not open Excel, refresh Power Query, calculate formulas, compile VBA, query a
Power Pivot model, or validate credentials.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any


SEVERITY_ORDER = {"info": 0, "low": 1, "medium": 2, "high": 3}
STATUS_ORDER = {"pass": 0, "review": 1, "blocked": 2}


def as_list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def as_dict(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_md(value: Any) -> str:
    return str(value if value is not None else "").replace("|", "\\|").replace("\n", " ")


def max_severity(*values: str) -> str:
    current = "info"
    for value in values:
        severity = str(value or "info").lower()
        if SEVERITY_ORDER.get(severity, 0) > SEVERITY_ORDER.get(current, 0):
            current = severity
    return current


def merge_status(*values: str) -> str:
    current = "pass"
    for value in values:
        status = str(value or "pass").lower()
        if STATUS_ORDER.get(status, 0) > STATUS_ORDER.get(current, 0):
            current = status
    return current


def severity_to_status(severity: str, finding_count: int = 0, readiness: str = "") -> str:
    severity = str(severity or "info").lower()
    readiness_l = str(readiness or "").lower()
    if "blocked" in readiness_l or "fail" in readiness_l:
        return "blocked"
    if severity == "high":
        return "blocked"
    if finding_count or severity in {"medium", "low"}:
        return "review"
    return "pass"


def inspect_workbook(workbook: Path) -> dict[str, Any]:
    inspector = Path(__file__).resolve().parent / "inspect_excel_bi_workbook.py"
    if not inspector.is_file():
        raise FileNotFoundError(f"inspector not found: {inspector}")
    with tempfile.TemporaryDirectory(prefix="excel_bi_triage_inspect_") as tmp:
        out_json = Path(tmp) / "openxml.json"
        result = subprocess.run(
            [sys.executable, str(inspector), str(workbook), "--out-json", str(out_json)],
            text=True,
            capture_output=True,
        )
        if result.returncode != 0:
            raise RuntimeError(
                "inspect_excel_bi_workbook.py failed: "
                + (result.stderr.strip() or result.stdout.strip() or f"exit {result.returncode}")
            )
        return load_json(out_json)


def function_counts(openxml: dict[str, Any]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in as_list(openxml.get("formulas")):
        if not isinstance(item, dict):
            continue
        formula = str(item.get("formula", ""))
        for raw in formula.replace("=", " ").replace("+", " ").replace("-", " ").replace("*", " ").replace("/", " ").split("("):
            token = raw.split()[-1].strip().upper() if raw.split() else ""
            if token and token.replace(".", "").replace("_", "").isalnum():
                counts[token] = counts.get(token, 0) + 1
    return counts


def workbook_surfaces(openxml: dict[str, Any]) -> dict[str, Any]:
    connections = as_list(openxml.get("connections"))
    external_links = as_list(openxml.get("externalLinks"))
    formulas = as_list(openxml.get("formulas"))
    defined_names = as_list(openxml.get("definedNames"))
    sheet_controls = as_list(openxml.get("sheetControls"))
    hidden_sheets = [
        item.get("name")
        for item in as_list(openxml.get("sheets"))
        if isinstance(item, dict) and str(item.get("state", "visible")).lower() != "visible"
    ]
    protected_sheets = [
        item.get("name")
        for item in sheet_controls
        if isinstance(item, dict) and item.get("hasSheetProtection")
    ]
    filtered_sheets = [
        item.get("name")
        for item in sheet_controls
        if isinstance(item, dict) and item.get("hasAutoFilter")
    ]
    frozen_pane_sheets = [
        item.get("name")
        for item in sheet_controls
        if isinstance(item, dict) and item.get("hasFrozenPane")
    ]
    data_validation_sheets = [
        item.get("name")
        for item in sheet_controls
        if isinstance(item, dict) and int(item.get("dataValidationCount") or 0) > 0
    ]
    cube_count = int(openxml.get("cubeFormulaCount") or len(as_list(openxml.get("cubeFormulas"))))
    has_connections = bool(connections)
    has_power_query = bool(openxml.get("hasMashupLikeParts")) or any(
        "mashup" in json.dumps(item, ensure_ascii=False).lower()
        or "power query" in json.dumps(item, ensure_ascii=False).lower()
        for item in connections
        if isinstance(item, dict)
    )
    has_power_pivot = bool(openxml.get("hasPowerPivotLikeParts")) or cube_count > 0 or any(
        "msolap" in json.dumps(item, ensure_ascii=False).lower()
        or "olap" in json.dumps(item, ensure_ascii=False).lower()
        for item in connections
        if isinstance(item, dict)
    )
    return {
        "sheetCount": len(as_list(openxml.get("sheets"))),
        "formulaCount": int(openxml.get("totalFormulaCount") or len(formulas)),
        "definedNameCount": len(defined_names),
        "tableCount": len(as_list(openxml.get("tables"))),
        "pivotCacheCount": len(as_list(openxml.get("pivotCaches"))),
        "connectionCount": len(connections),
        "externalLinkCount": len(external_links),
        "cubeFormulaCount": cube_count,
        "hasPowerQueryLikeParts": has_power_query,
        "hasPowerPivotLikeParts": has_power_pivot,
        "hasVbaProject": bool(openxml.get("hasVbaProject")),
        "chartPartCount": len(as_list(openxml.get("chartParts"))),
        "drawingPartCount": len(as_list(openxml.get("drawingParts"))),
        "hiddenSheets": hidden_sheets,
        "protectedSheets": protected_sheets,
        "filteredSheets": filtered_sheets,
        "frozenPaneSheets": frozen_pane_sheets,
        "dataValidationSheets": data_validation_sheets,
        "functionCounts": function_counts(openxml),
    }


def report_signal(kind: str, report: dict[str, Any]) -> dict[str, Any]:
    summary = as_dict(report.get("summary"))
    findings = as_list(report.get("findings"))
    readiness = str(summary.get("readiness", ""))
    max_sev = str(summary.get("maxSeverity", "info") or "info").lower()
    finding_count = int(summary.get("findingCount") or len(findings) or 0)

    if kind == "cube":
        flags = as_dict(report.get("byDiagnosticFlag"))
        missing = as_list(report.get("missingModelMeasures"))
        finding_count = len(missing) + sum(int(v or 0) for v in flags.values())
        max_sev = "high" if missing else ("medium" if finding_count else "info")
        readiness = "review-required" if finding_count else "clean"

    if kind == "model":
        missing = as_list(report.get("cubeFormulaReferencesMissingModelMeasure"))
        finding_count = len(missing)
        max_sev = "high" if missing else "info"
        readiness = "review-required" if missing else "clean"

    status = severity_to_status(max_sev, finding_count, readiness)
    return {
        "kind": kind,
        "status": status,
        "readiness": readiness or ("clean" if finding_count == 0 else "review-required"),
        "maxSeverity": max_sev,
        "findingCount": finding_count,
        "highFindingCount": int(summary.get("highFindingCount") or (finding_count if max_sev == "high" else 0)),
        "mediumFindingCount": int(summary.get("mediumFindingCount") or (finding_count if max_sev == "medium" else 0)),
        "sourceWorkbook": report.get("workbookPath") or report.get("sourceInspector") or "",
    }


def add_gap(gaps: list[dict[str, Any]], kind: str, severity: str, title: str, reason: str, command: str) -> None:
    gaps.append(
        {
            "kind": kind,
            "severity": severity,
            "title": title,
            "reason": reason,
            "recommendedCommand": command,
        }
    )


def build_coverage_gaps(surfaces: dict[str, Any], provided: set[str]) -> list[dict[str, Any]]:
    gaps: list[dict[str, Any]] = []
    if surfaces["formulaCount"] and "formula" not in provided:
        add_gap(
            gaps,
            "formula",
            "medium",
            "Formula quality report not supplied",
            "Workbook contains formulas, but static formula risk checks were not included in this triage.",
            "python tools/build_formula_quality_report.py --openxml-json openxml.json --out-json formula_quality.json --out-md formula_quality.md",
        )
    if "controls" not in provided:
        add_gap(
            gaps,
            "controls",
            "low",
            "Workbook controls report not supplied",
            "Sheet visibility, protection, filters, freeze panes, and data validation should be intentional before handoff.",
            "python tools/build_workbook_controls_report.py --openxml-json openxml.json --out-json controls.json --out-md controls.md",
        )
    if (
        surfaces["connectionCount"]
        or surfaces["externalLinkCount"]
        or surfaces["hasPowerQueryLikeParts"]
        or surfaces["hasPowerPivotLikeParts"]
        or surfaces["cubeFormulaCount"]
    ) and "external" not in provided:
        add_gap(
            gaps,
            "external",
            "high",
            "External dependency report not supplied",
            "The workbook has dependency surfaces that can block a pure-value or self-contained deliverable.",
            "python tools/build_external_dependency_report.py --openxml-json openxml.json --out-json external_dependencies.json --out-md external_dependencies.md",
        )
    if surfaces["cubeFormulaCount"] and "cube" not in provided:
        add_gap(
            gaps,
            "cube",
            "medium",
            "CUBE dependency report not supplied",
            "CUBE formulas need measure/member dependency checks before Data Model or OLAP delivery.",
            "python tools/build_cube_dependency_report.py --openxml-json openxml.json --model-summary-json model_report.json --out-json cube_dependencies.json --out-md cube_dependencies.md",
        )
    if surfaces["hasPowerQueryLikeParts"] and "powerQueryLineage" not in provided:
        add_gap(
            gaps,
            "powerQueryLineage",
            "medium",
            "Power Query lineage report not supplied",
            "Mashup-like parts or query connections are present; exported M should be checked for source, credential, native-query, and dependency risks.",
            "python tools/build_power_query_lineage_report.py exported_power_queries --out-json pq_lineage.json --out-md pq_lineage.md",
        )
    if surfaces["hasPowerPivotLikeParts"] and "model" not in provided:
        add_gap(
            gaps,
            "model",
            "medium",
            "Power Pivot/Data Model report not supplied",
            "Model tables, measures, relationships, and CUBE formula dependencies need Excel COM inspection when a live model is delivered.",
            "python tools/build_excel_bi_model_report.py --model-json model_com.json --openxml-json openxml.json --out-json model_report.json --out-md model_report.md",
        )
    if surfaces["hasVbaProject"] and "vbaButton" not in provided:
        add_gap(
            gaps,
            "vba",
            "medium",
            "VBA/button binding report not supplied",
            "Macro-enabled delivery needs exported VBA linting, compile/run validation, and button OnAction checks.",
            "python tools/build_vba_button_binding_report.py --openxml-json openxml.json --vba-source-dir exported_vba --out-json vba_buttons.json --out-md vba_buttons.md",
        )
    return gaps


def action(
    actions: list[dict[str, Any]],
    priority: str,
    area: str,
    title: str,
    why: str,
    command: str,
    boundary: str,
) -> None:
    actions.append(
        {
            "priority": priority,
            "area": area,
            "title": title,
            "why": why,
            "command": command,
            "boundary": boundary,
        }
    )


def build_actions(surfaces: dict[str, Any], gaps: list[dict[str, Any]], signals: list[dict[str, Any]]) -> list[dict[str, Any]]:
    actions: list[dict[str, Any]] = []
    for gap in gaps:
        action(
            actions,
            "high" if gap["severity"] == "high" else "medium",
            gap["kind"],
            gap["title"],
            gap["reason"],
            gap["recommendedCommand"],
            "Static triage can route this check, but the specialized report owns the detailed finding logic.",
        )

    blocked = [item for item in signals if item.get("status") == "blocked"]
    if blocked:
        action(
            actions,
            "high",
            "delivery-readiness",
            "Resolve blocked specialized reports before handoff",
            "At least one supplied report contains high-severity or blocked readiness findings.",
            "Review the supplied report Markdown, refresh/recalculate in Excel when needed, then rerun this triage.",
            "The triage report does not fix workbook content; it orders the follow-up work.",
        )

    if surfaces["hasPowerQueryLikeParts"]:
        action(
            actions,
            "medium",
            "power-query",
            "Validate Power Query refresh behavior in Excel",
            "OpenXML can detect query-like package parts, but only Excel can prove refresh order, privacy prompts, credentials, and loaded output correctness.",
            "pwsh -File .agents/skills/power-query-m-engineering/scripts/refresh_power_query_and_wait.ps1 -WorkbookPath workbook.xlsx -OutJson pq_refresh.json",
            "Live refresh validation requires desktop Excel on Windows.",
        )
    if surfaces["hasPowerPivotLikeParts"] or surfaces["cubeFormulaCount"]:
        action(
            actions,
            "medium",
            "power-pivot",
            "Validate Data Model/CUBE calculation in Excel",
            "CUBE formulas and model metadata are structural until Excel recalculates against a real model or OLAP connection.",
            "pwsh -File tools/inspect_excel_data_model_com.ps1 -WorkbookPath workbook.xlsx -OutJson model_com.json",
            "Data Model inspection requires desktop Excel COM; static OpenXML cannot execute DAX or MDX.",
        )
    if surfaces["hasVbaProject"]:
        action(
            actions,
            "medium",
            "vba",
            "Export, lint, compile, and smoke-run VBA",
            "Macro-enabled workbooks need runtime validation that OpenXML cannot provide.",
            "pwsh -File .agents/skills/excel-vba-workbook-engineering/scripts/export_vba.ps1 -WorkbookPath workbook.xlsm -OutDir exported_vba",
            "VBA compile/run validation requires Excel with trusted VBA project access.",
        )
    if not actions:
        action(
            actions,
            "low",
            "handoff",
            "No immediate static risks detected",
            "The supplied static reports are clean and no unreported high-risk workbook surfaces were found.",
            "For final delivery, open in Excel once, recalculate/refresh if applicable, and save the verified copy.",
            "A clean static triage is not a substitute for live Excel validation when the workbook contains live features.",
        )
    return actions


def build_report(
    openxml: dict[str, Any],
    specialized_reports: dict[str, dict[str, Any]],
    source: dict[str, Any],
) -> dict[str, Any]:
    surfaces = workbook_surfaces(openxml)
    provided = set(specialized_reports)
    signals = [report_signal(kind, report) for kind, report in sorted(specialized_reports.items())]
    gaps = build_coverage_gaps(surfaces, provided)

    status = "pass"
    max_sev = "info"
    for signal in signals:
        status = merge_status(status, str(signal.get("status", "pass")))
        max_sev = max_severity(max_sev, str(signal.get("maxSeverity", "info")))
    for gap in gaps:
        status = merge_status(status, severity_to_status(str(gap.get("severity", "info")), 1))
        max_sev = max_severity(max_sev, str(gap.get("severity", "info")))

    actions = build_actions(surfaces, gaps, signals)
    workbook = {
        "path": openxml.get("workbookPath") or source.get("workbook") or "",
        "fileType": openxml.get("fileType", ""),
        "source": source,
    }
    return {
        "status": status,
        "readiness": "ready-for-live-validation" if status == "pass" else ("blocked-before-handoff" if status == "blocked" else "review-required"),
        "maxSeverity": max_sev,
        "workbook": workbook,
        "surfaces": surfaces,
        "reportSignals": signals,
        "coverageGaps": gaps,
        "recommendedNextActions": actions,
        "limitations": [
            "Static triage does not refresh Power Query, execute SQL/MDX, evaluate CUBE formulas, calculate formulas, or compile/run VBA.",
            "Power Pivot/Data Model metadata and DAX behavior require Excel COM or another live model inspection path.",
            "A pure-value deliverable still needs an Excel-opened refresh/recalculate/value-freeze pass when the workbook has live sources.",
        ],
    }


def render_table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(clean_md(value) for value in row) + " |")
    return lines


def render_markdown(report: dict[str, Any]) -> str:
    surfaces = report["surfaces"]
    lines: list[str] = [
        "# Workbook Triage Report",
        "",
        f"- status: `{report.get('status')}`",
        f"- readiness: `{report.get('readiness')}`",
        f"- max severity: `{report.get('maxSeverity')}`",
        f"- workbook: `{report.get('workbook', {}).get('path', '')}`",
        "",
        "## Workbook Surfaces",
        "",
    ]
    lines.extend(
        render_table(
            ["Surface", "Count / State"],
            [
                ["worksheets", surfaces["sheetCount"]],
                ["formulas", surfaces["formulaCount"]],
                ["defined names", surfaces["definedNameCount"]],
                ["tables", surfaces["tableCount"]],
                ["connections", surfaces["connectionCount"]],
                ["external links", surfaces["externalLinkCount"]],
                ["CUBE formulas", surfaces["cubeFormulaCount"]],
                ["Power Query-like parts", surfaces["hasPowerQueryLikeParts"]],
                ["Power Pivot-like parts", surfaces["hasPowerPivotLikeParts"]],
                ["VBA project", surfaces["hasVbaProject"]],
                ["hidden/protected/filtered sheets", f"{len(surfaces['hiddenSheets'])}/{len(surfaces['protectedSheets'])}/{len(surfaces['filteredSheets'])}"],
            ],
        )
    )
    lines.extend(["", "## Supplied Specialized Reports", ""])
    signals = as_list(report.get("reportSignals"))
    if signals:
        lines.extend(
            render_table(
                ["Report", "Status", "Readiness", "Severity", "Findings"],
                [
                    [
                        item.get("kind"),
                        item.get("status"),
                        item.get("readiness"),
                        item.get("maxSeverity"),
                        item.get("findingCount"),
                    ]
                    for item in signals
                    if isinstance(item, dict)
                ],
            )
        )
    else:
        lines.append("_No specialized report JSON was supplied._")

    lines.extend(["", "## Coverage Gaps", ""])
    gaps = as_list(report.get("coverageGaps"))
    if gaps:
        lines.extend(
            render_table(
                ["Area", "Severity", "Reason", "Recommended command"],
                [
                    [
                        item.get("kind"),
                        item.get("severity"),
                        item.get("reason"),
                        f"`{item.get('recommendedCommand')}`",
                    ]
                    for item in gaps
                    if isinstance(item, dict)
                ],
            )
        )
    else:
        lines.append("_No coverage gaps were detected from the supplied static inputs._")

    lines.extend(["", "## Recommended Next Actions", ""])
    lines.extend(
        render_table(
            ["Priority", "Area", "Action", "Why", "Boundary"],
            [
                [
                    item.get("priority"),
                    item.get("area"),
                    item.get("title"),
                    item.get("why"),
                    item.get("boundary"),
                ]
                for item in as_list(report.get("recommendedNextActions"))
                if isinstance(item, dict)
            ],
        )
    )

    lines.extend(["", "## Boundary", ""])
    for item in as_list(report.get("limitations")):
        lines.append(f"- {item}")
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a workbook triage report from OpenXML inventory and optional "
            "specialized Excel BI static reports."
        )
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--workbook", type=Path, help="Workbook to inspect with tools/inspect_excel_bi_workbook.py")
    source.add_argument("--inspection-json", type=Path, help="JSON from tools/inspect_excel_bi_workbook.py")
    parser.add_argument("--external-report-json", type=Path)
    parser.add_argument("--formula-report-json", type=Path)
    parser.add_argument("--controls-report-json", type=Path)
    parser.add_argument("--cube-report-json", type=Path)
    parser.add_argument("--power-query-lineage-json", type=Path)
    parser.add_argument("--model-report-json", type=Path)
    parser.add_argument("--provider-report-json", type=Path)
    parser.add_argument("--vba-button-report-json", type=Path)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-md", type=Path)
    parser.add_argument("--print", action="store_true", help="Print Markdown report")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero unless triage status is pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.workbook:
        openxml = inspect_workbook(args.workbook)
        source = {"workbook": str(args.workbook), "inspection": "generated"}
    else:
        openxml = load_json(args.inspection_json)
        source = {"inspectionJson": str(args.inspection_json), "inspection": "supplied"}

    report_paths = {
        "external": args.external_report_json,
        "formula": args.formula_report_json,
        "controls": args.controls_report_json,
        "cube": args.cube_report_json,
        "powerQueryLineage": args.power_query_lineage_json,
        "model": args.model_report_json,
        "provider": args.provider_report_json,
        "vbaButton": args.vba_button_report_json,
    }
    specialized_reports = {
        kind: load_json(path)
        for kind, path in report_paths.items()
        if path is not None
    }
    report = build_report(openxml, specialized_reports, source)
    markdown = render_markdown(report)

    if args.out_json:
        write_json(args.out_json, report)
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(markdown, encoding="utf-8")
    if args.print:
        print(markdown)
    elif not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    if args.require_pass and report.get("status") != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
