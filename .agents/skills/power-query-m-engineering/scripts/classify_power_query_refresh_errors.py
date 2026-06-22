#!/usr/bin/env python3
"""Classify Power Query refresh errors into actionable diagnostic buckets.

The input can be a JSON report produced by refresh_power_queries_excel_com.ps1,
a JSON object from another wrapper, or plain copied error text. This script does
not refresh Excel, inspect credentials, or prove that a fix works; it turns raw
error text into a repeatable triage report.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Rule:
    code: str
    title: str
    severity: str
    patterns: tuple[str, ...]
    likely_cause: str
    next_steps: tuple[str, ...]


RULES: tuple[Rule, ...] = (
    Rule(
        code="credentials-or-permissions",
        title="Credentials or source permissions",
        severity="error",
        patterns=(
            r"\bcredential",
            r"\bauthentication\b",
            r"\bauthorization\b",
            r"\baccess\s+denied\b",
            r"\bpermission\b",
            r"\bsign\s*in\b",
        ),
        likely_cause="The source requires credentials, permissions, or an account session that VBA cannot repair safely.",
        next_steps=(
            "Open the workbook in Excel and refresh the failing query interactively once.",
            "Re-authenticate the source through Data Source Settings.",
            "Do not hide this class of error with retry loops.",
        ),
    ),
    Rule(
        code="privacy-firewall",
        title="Power Query privacy firewall",
        severity="error",
        patterns=(
            r"Formula\.Firewall",
            r"\bprivacy\s+level",
            r"\bdata\s+privacy\b",
            r"\bfirewall\b",
        ),
        likely_cause="Power Query is combining sources whose privacy levels or staging pattern block evaluation.",
        next_steps=(
            "Review Data Source Settings and privacy levels.",
            "Stage source queries separately, then combine cleaned staging queries.",
            "Avoid relying on VBA to bypass privacy evaluation.",
        ),
    ),
    Rule(
        code="missing-source",
        title="Missing file, folder, or source object",
        severity="error",
        patterns=(
            r"\bfile\s+not\s+found\b",
            r"\bcould\s+not\s+find\s+(?:file|path)",
            r"\bpath\s+.*not\s+found\b",
            r"\bfolder\s+.*not\s+found\b",
            r"\bthe\s+system\s+cannot\s+find\b",
            r"\bdoes\s+not\s+exist\b",
        ),
        likely_cause="A file path, folder path, sheet, table, or external source name changed or is unavailable.",
        next_steps=(
            "Verify the source path and source file list before opening Power Query.",
            "Filter hidden and temporary files when using Folder.Files.",
            "Parameterize environment-specific paths instead of hard-coding local paths.",
        ),
    ),
    Rule(
        code="query-not-found",
        title="Workbook query not found",
        severity="error",
        patterns=(r"\bquery\s+not\s+found\b", r"\bname\s+.*was\s+not\s+recognized\b"),
        likely_cause="The requested query name does not match Workbook.Queries.",
        next_steps=(
            "Run the query-list action and copy the exact query name.",
            "Check whitespace, localized names, and renamed queries before calling refresh.",
        ),
    ),
    Rule(
        code="missing-column",
        title="Missing column or schema drift",
        severity="error",
        patterns=(
            r"\bcolumn\b.*\bnot\s+found\b",
            r"\bcolumn\b.*\bwasn'?t\s+found\b",
            r"\bfield\b.*\bnot\s+found\b",
            r"\bfield\b.*\bwasn'?t\s+found\b",
        ),
        likely_cause="The source schema changed, headers were promoted differently, or a hard-coded column list is stale.",
        next_steps=(
            "Inspect the source columns at the first failing step.",
            "Use MissingField.UseNull where missing columns are acceptable.",
            "Normalize headers before joins, expands, and type conversions.",
        ),
    ),
    Rule(
        code="missing-workbook-item",
        title="Missing workbook item or navigation key",
        severity="error",
        patterns=(
            r"\bkey\s+did\s+not\s+match\s+any\s+rows\b",
            r"\bkey\s+didn'?t\s+match\s+any\s+rows\b",
            r"\bitem\s+.*not\s+found\b",
        ),
        likely_cause="The query navigates Excel.Workbook output by a hard-coded sheet/table/name that no longer exists.",
        next_steps=(
            "Inspect the Excel.Workbook navigation table.",
            "Select by Kind and normalized Name only after confirming the item exists.",
            "Add a clearer error step when a required sheet/table is absent.",
        ),
    ),
    Rule(
        code="type-conversion",
        title="Type conversion or dirty value",
        severity="error",
        patterns=(
            r"\bcannot\s+convert\b",
            r"\bcouldn'?t\s+convert\b",
            r"\bwe\s+cannot\s+convert\b",
            r"\bDataFormat\.Error\b",
            r"\binvalid\s+(?:cell\s+)?value\b",
            r"\bNumber\.From\b",
            r"\bDate\.From\b",
        ),
        likely_cause="A source value cannot be converted to the requested type, often because nulls, blanks, text, or localized values were not guarded.",
        next_steps=(
            "Move type conversion after null and blank handling.",
            "Use try ... otherwise null for dirty numeric/date fields.",
            "Inspect distinct bad values before changing the final data type.",
        ),
    ),
    Rule(
        code="syntax-or-formula",
        title="M syntax or formula error",
        severity="error",
        patterns=(
            r"\bExpression\.SyntaxError\b",
            r"\bToken\s+.*expected\b",
            r"\binvalid\s+identifier\b",
            r"\bcyclic\s+reference\b",
        ),
        likely_cause="The M formula is malformed or has a dependency cycle.",
        next_steps=(
            "Run lint_power_query_m.py against the changed source.",
            "Check step commas, quoted step names, final in expression, and cyclic references.",
        ),
    ),
    Rule(
        code="connector-provider",
        title="Connector, provider, or driver error",
        severity="error",
        patterns=(
            r"\bOLE\s*DB\b",
            r"\bODBC\b",
            r"\bprovider\b.*\bnot\s+found\b",
            r"\bdriver\b.*\bnot\s+found\b",
            r"Microsoft\.Mashup\.OleDb",
            r"\bDataSource\.Error\b",
        ),
        likely_cause="The required connector, OLE DB provider, driver, or mashup provider is missing or failing.",
        next_steps=(
            "Run the provider probe for Excel COM, ACE, MSOLAP, ADODB, and ADOMD evidence.",
            "Check 32-bit vs 64-bit Office/provider alignment.",
            "Separate provider installation issues from M formula issues.",
        ),
    ),
    Rule(
        code="timeout-or-background-refresh",
        title="Refresh timeout or background refresh still running",
        severity="error",
        patterns=(
            r"\btimed\s+out\b",
            r"\bstill\s+refreshing\b",
            r"\bbackground\s+refresh\b",
        ),
        likely_cause="Queries continued asynchronously, a source is slow, or refresh waited on a blocked connection.",
        next_steps=(
            "Disable background refresh before dependent macros continue.",
            "Refresh only the failing query and capture elapsed time.",
            "Check whether a prompt or credential dialog is blocking a headless refresh.",
        ),
    ),
    Rule(
        code="row-count-or-cardinality",
        title="Unexpected row count or join cardinality",
        severity="warning",
        patterns=(
            r"\brow\s+count\b.*\bincreased\b",
            r"\bduplicate\s+key\b",
            r"\bmany\s+matches\b",
            r"\bcardinality\b",
        ),
        likely_cause="A join or expand step may be multiplying rows because lookup keys are not unique.",
        next_steps=(
            "Pre-aggregate or de-duplicate the lookup side before Table.NestedJoin.",
            "Record row counts before and after joins and expands.",
            "Surface duplicate-source problems as diagnostics instead of expanding all matches.",
        ),
    ),
)


def read_input(path: Path) -> tuple[str, Any | None]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    try:
        return text, json.loads(text)
    except json.JSONDecodeError:
        return text, None


def collect_strings(value: Any, label: str = "$") -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    if isinstance(value, dict):
        for key, item in value.items():
            items.extend(collect_strings(item, f"{label}.{key}"))
    elif isinstance(value, list):
        for index, item in enumerate(value):
            items.extend(collect_strings(item, f"{label}[{index}]"))
    elif isinstance(value, str) and value.strip():
        items.append({"path": label, "text": value.strip()})
    return items


def compact_snippet(text: str, limit: int = 220) -> str:
    compact = re.sub(r"\s+", " ", text).strip()
    if len(compact) <= limit:
        return compact
    return compact[: limit - 3] + "..."


def classify_messages(messages: list[dict[str, str]]) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    for rule in RULES:
        evidence: list[dict[str, str]] = []
        compiled = [re.compile(pattern, re.IGNORECASE) for pattern in rule.patterns]
        for message in messages:
            if any(pattern.search(message["text"]) for pattern in compiled):
                evidence.append({"path": message["path"], "snippet": compact_snippet(message["text"])})
        if evidence:
            findings.append(
                {
                    "code": rule.code,
                    "title": rule.title,
                    "severity": rule.severity,
                    "likelyCause": rule.likely_cause,
                    "nextSteps": list(rule.next_steps),
                    "evidence": evidence[:5],
                    "evidenceCount": len(evidence),
                }
            )
    return findings


def infer_refresh_status(data: Any | None, findings: list[dict[str, Any]]) -> str:
    if findings:
        return "needs-diagnosis"
    if isinstance(data, dict):
        if data.get("error"):
            return "unclassified-error"
        errors = data.get("errors")
        if isinstance(errors, list) and errors:
            return "unclassified-error"
        if data.get("failedAt"):
            return "unclassified-error"
        if data.get("completedAt"):
            return "no-known-errors"
    return "no-known-errors"


def build_report(input_path: Path) -> dict[str, Any]:
    raw_text, data = read_input(input_path)
    messages = collect_strings(data) if data is not None else [{"path": "$", "text": raw_text}]
    findings = classify_messages(messages)
    return {
        "source": str(input_path),
        "inputKind": "json" if data is not None else "text",
        "messageCount": len(messages),
        "status": infer_refresh_status(data, findings),
        "findingCount": len(findings),
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "# Power Query Refresh Error Classification",
        "",
        f"- Source: `{report['source']}`",
        f"- Input kind: `{report['inputKind']}`",
        f"- Status: **{report['status']}**",
        f"- Findings: `{report['findingCount']}`",
        "",
    ]
    if not report["findings"]:
        lines.append("No known Power Query refresh error category was detected.")
        return "\n".join(lines) + "\n"

    for finding in report["findings"]:
        lines.extend(
            [
                f"## {finding['title']}",
                "",
                f"- Code: `{finding['code']}`",
                f"- Severity: `{finding['severity']}`",
                f"- Likely cause: {finding['likelyCause']}",
                "- Next steps:",
            ]
        )
        for step in finding["nextSteps"]:
            lines.append(f"  - {step}")
        lines.append("- Evidence:")
        for evidence in finding["evidence"]:
            lines.append(f"  - `{evidence['path']}`: {evidence['snippet']}")
        lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Refresh JSON report or plain text error log")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--out-md", type=Path, help="Optional Markdown report path")
    parser.add_argument("--fail-on-diagnosis", action="store_true", help="Exit with 1 when any known diagnosis is found")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input.expanduser().resolve()
    report = build_report(input_path)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_report(report), encoding="utf-8")

    print(
        f"Power Query refresh classification: {report['status']}, "
        f"{report['findingCount']} findings from {report['messageCount']} messages"
    )
    return 1 if args.fail_on_diagnosis and report["findingCount"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
