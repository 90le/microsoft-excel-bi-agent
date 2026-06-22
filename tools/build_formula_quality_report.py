#!/usr/bin/env python3
"""Build a static formula quality report from OpenXML workbook inspection JSON.

Input is the JSON produced by ``tools/inspect_excel_bi_workbook.py``. The report
is read-only: it does not open Excel, recalculate formulas, refresh Power Query,
or decide whether a business result is correct. It highlights delivery risks
that are visible from formula text and cached formula values.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


ERROR_VALUES = {"#VALUE!", "#REF!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}
LOCAL_PATH_RE = re.compile(r"(?i)(?:[A-Z]:\\|file:///|\\\\|/Users/|/home/)")
REF_ERROR_RE = re.compile(r"#REF!", re.IGNORECASE)
FUNCTION_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
VOLATILE_FUNCTIONS = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"}
HIGH_RISK_DYNAMIC_FUNCTIONS = {"INDIRECT", "OFFSET"}
SEVERITY_ORDER = {"info": 0, "low": 1, "medium": 2, "high": 3}


def as_list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def normalize_function_name(name: str) -> str:
    text = name.upper()
    if text.startswith("_XLFN."):
        text = text[6:]
    if "." in text:
        text = text.rsplit(".", 1)[-1]
    return text


def extract_functions(formula: str) -> list[str]:
    seen: set[str] = set()
    functions: list[str] = []
    for match in FUNCTION_RE.finditer(formula):
        function = normalize_function_name(match.group(1))
        if function and function not in seen:
            seen.add(function)
            functions.append(function)
    return functions


def formula_ref(item: dict[str, Any]) -> str:
    return f"{item.get('sheet', '')}!{item.get('cell', '')}"


def add_finding(
    findings: list[dict[str, Any]],
    code: str,
    severity: str,
    title: str,
    formula: dict[str, Any],
    evidence: dict[str, Any],
    action: str,
) -> None:
    findings.append(
        {
            "code": code,
            "severity": severity,
            "title": title,
            "sheet": formula.get("sheet", ""),
            "cell": formula.get("cell", ""),
            "formula": formula.get("formula", ""),
            "cachedValue": formula.get("cachedValue"),
            "evidence": evidence,
            "recommendedAction": action,
        }
    )


def build_report(openxml_report: dict[str, Any]) -> dict[str, Any]:
    formulas = [item for item in as_list(openxml_report.get("formulas")) if isinstance(item, dict)]
    findings: list[dict[str, Any]] = []
    function_counts: dict[str, int] = {}
    formulas_with_no_cached_value = 0

    for formula in formulas:
        formula_text = str(formula.get("formula", "") or "")
        cached_value = formula.get("cachedValue")
        cached_text = str(cached_value or "").strip()
        if cached_value is None:
            formulas_with_no_cached_value += 1

        functions = extract_functions(formula_text)
        for function in functions:
            function_counts[function] = function_counts.get(function, 0) + 1

        if cached_text.upper() in ERROR_VALUES:
            add_finding(
                findings,
                "cached-formula-error",
                "high",
                "Formula has a cached Excel error value",
                formula,
                {"cachedValue": cached_text, "cell": formula_ref(formula)},
                "Open the workbook in Excel, refresh/recalculate, then fix the upstream formula or data source before handoff.",
            )

        if REF_ERROR_RE.search(formula_text):
            add_finding(
                findings,
                "formula-ref-error",
                "high",
                "Formula text contains #REF!",
                formula,
                {"cell": formula_ref(formula)},
                "Repair deleted/moved sheet, row, column, or name references before delivery.",
            )

        if LOCAL_PATH_RE.search(formula_text):
            add_finding(
                findings,
                "local-path-formula",
                "high",
                "Formula references a local filesystem path",
                formula,
                {"cell": formula_ref(formula)},
                "Replace the path-dependent formula with a portable workbook reference, a refreshed value, or a documented live dependency.",
            )

        risky_dynamic = sorted(function for function in functions if function in HIGH_RISK_DYNAMIC_FUNCTIONS)
        if risky_dynamic:
            add_finding(
                findings,
                "dynamic-reference-function",
                "medium",
                "Formula uses dynamic reference functions",
                formula,
                {"functions": risky_dynamic, "cell": formula_ref(formula)},
                "Review these formulas manually because INDIRECT/OFFSET can hide dependencies from static auditing and break after sheet/name changes.",
            )

        volatile = sorted(function for function in functions if function in VOLATILE_FUNCTIONS - HIGH_RISK_DYNAMIC_FUNCTIONS)
        if volatile:
            add_finding(
                findings,
                "volatile-function",
                "low",
                "Formula uses volatile functions",
                formula,
                {"functions": volatile, "cell": formula_ref(formula)},
                "Confirm volatility is intended and that cached values were recalculated before a static handoff.",
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
        "formulaCount": len(formulas),
        "findingCount": len(findings),
        "highFindingCount": severity_counts.get("high", 0),
        "mediumFindingCount": severity_counts.get("medium", 0),
        "lowFindingCount": severity_counts.get("low", 0),
        "formulasWithNoCachedValue": formulas_with_no_cached_value,
        "functionCounts": dict(sorted(function_counts.items())),
    }
    return {
        "workbookPath": openxml_report.get("workbookPath", ""),
        "sourceInspector": "tools/inspect_excel_bi_workbook.py",
        "summary": summary,
        "findings": findings,
        "limitations": [
            "Static OpenXML inspection only; formulas were not recalculated.",
            "A clean report means no selected static formula risks were found, not that numeric results are business-correct.",
            "Use Windows Excel COM or manual Excel validation for final calculation, refresh, and chart/rendering checks.",
        ],
    }


def render_markdown(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    lines = [
        "# Formula Quality Report",
        "",
        f"- workbook: `{report.get('workbookPath', '')}`",
        f"- readiness: **{summary.get('readiness', '')}**",
        f"- max severity: `{summary.get('maxSeverity', '')}`",
        f"- formulas: `{summary.get('formulaCount', 0)}`",
        f"- findings: `{summary.get('findingCount', 0)}`",
        "",
        "| Code | Severity | Cell | Evidence | Recommended action |",
        "|---|---:|---|---|---|",
    ]
    for finding in report.get("findings", []):
        evidence = json.dumps(finding.get("evidence", {}), ensure_ascii=False, sort_keys=True)
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(finding.get("code", "")),
                    clean_markdown(finding.get("severity", "")),
                    clean_markdown(f"{finding.get('sheet', '')}!{finding.get('cell', '')}"),
                    clean_markdown(evidence),
                    clean_markdown(finding.get("recommendedAction", "")),
                ]
            )
            + " |"
        )
    if not report.get("findings"):
        lines.append("| none | info | n/a | no selected static formula risks detected | No action required by this static check. |")
    lines.extend(["", "## Function Counts", ""])
    function_counts = summary.get("functionCounts", {})
    if isinstance(function_counts, dict) and function_counts:
        for function, count in function_counts.items():
            lines.append(f"- `{function}`: {count}")
    else:
        lines.append("- none")
    lines.extend(["", "## Limitations", ""])
    for item in report.get("limitations", []):
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--openxml-json", required=True, type=Path, help="JSON from inspect_excel_bi_workbook.py")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--fail-on-high-risk", action="store_true", help="Exit with code 1 when high-risk findings are present")
    parser.add_argument("--fail-on-review", action="store_true", help="Exit with code 1 when any finding is present")
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
        "Formula quality {readiness}: formulas={formulas}, findings={findings}, high={high}".format(
            readiness=summary.get("readiness", ""),
            formulas=summary.get("formulaCount", 0),
            findings=summary.get("findingCount", 0),
            high=summary.get("highFindingCount", 0),
        )
    )
    if args.fail_on_review and int(summary.get("findingCount") or 0) > 0:
        return 1
    if args.fail_on_high_risk and int(summary.get("highFindingCount") or 0) > 0:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
