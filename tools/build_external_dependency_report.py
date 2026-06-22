#!/usr/bin/env python3
"""Build a delivery-readiness report for Excel external dependencies.

Input is the JSON produced by ``tools/inspect_excel_bi_workbook.py``. The report
is static and read-only: it does not refresh queries, open external files,
evaluate formulas, break links, or remove connections. Use it before creating a
pure-value deliverable or before editing workbook connections.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


EXTERNAL_FORMULA_RE = re.compile(r"\[[^\]]+\](?!\])")
SEVERITY_ORDER = {"info": 0, "low": 1, "medium": 2, "high": 3}
CONNECTION_SECRET_KEY_RE = re.compile(
    r"(?i)\b("
    r"password|pwd|user id|uid|accountkey|sharedaccesssignature|sas token|"
    r"api key|apikey|client secret|clientsecret|authorization|access token|accesstoken|"
    r"refresh token|refreshtoken|secret"
    r")\s*="
)


def as_list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def connection_kind(connection: dict[str, Any]) -> str:
    text = " ".join(
        [
            str(connection.get("name", "")),
            str(connection.get("type", "")),
            str(connection.get("dbPr", {}).get("connection", "")),
            str(connection.get("dbPr", {}).get("command", "")),
        ]
    ).lower()
    if "mashup" in text or "power query" in text or "$workbook$" in text:
        return "power-query-like"
    if "msolap" in text or "olap" in text or "analysis services" in text:
        return "olap-or-data-model"
    if "oledb" in text or "odbc" in text or "provider=" in text:
        return "external-data-connection"
    return "workbook-connection"


def connection_secret_indicators(connection: dict[str, Any]) -> list[str]:
    texts = [
        str(connection.get("dbPr", {}).get("connection", "")),
        str(connection.get("dbPr", {}).get("command", "")),
        str(connection.get("name", "")),
        str(connection.get("description", "")),
    ]
    indicators: set[str] = set()
    for text in texts:
        for match in CONNECTION_SECRET_KEY_RE.finditer(text):
            indicators.add(match.group(1).lower())
    return sorted(indicators)


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

    connections = [item for item in as_list(openxml_report.get("connections")) if isinstance(item, dict)]
    if connections:
        kinds: dict[str, int] = {}
        names: list[str] = []
        credential_like_connections: list[dict[str, Any]] = []
        for connection in connections:
            kind = connection_kind(connection)
            kinds[kind] = kinds.get(kind, 0) + 1
            name = str(connection.get("name", ""))
            names.append(name)
            indicators = connection_secret_indicators(connection)
            if indicators:
                credential_like_connections.append(
                    {
                        "name": name,
                        "kind": kind,
                        "indicators": indicators,
                    }
                )
        add_finding(
            findings,
            "workbook-connections",
            "high",
            "Workbook contains data connections",
            {"count": len(connections), "names": names, "kinds": kinds},
            "For a pure deliverable, refresh first, convert dependent outputs to values, then remove workbook connections and query loads. For a live model deliverable, validate credentials, privacy levels, refresh order, and provider availability.",
        )
        if credential_like_connections:
            add_finding(
                findings,
                "connection-credential-like-literal",
                "high",
                "Workbook connection metadata contains credential-like key names",
                {"count": len(credential_like_connections), "connections": credential_like_connections},
                "Remove embedded credentials from workbook connection strings before delivery. Recreate credentials through Excel/Power Query Data Source Settings, managed gateways, or environment-specific connection configuration.",
            )

    external_links = as_list(openxml_report.get("externalLinks"))
    if external_links:
        add_finding(
            findings,
            "external-link-parts",
            "high",
            "Workbook contains external-link package parts",
            {"count": len(external_links), "parts": external_links[:25]},
            "Inspect Excel external links and either break/remove them after value conversion or document the source workbook requirement before delivery.",
        )

    formulas = [item for item in as_list(openxml_report.get("formulas")) if isinstance(item, dict)]
    external_formulas = [
        item
        for item in formulas
        if EXTERNAL_FORMULA_RE.search(str(item.get("formula", "")))
    ]
    if external_formulas:
        add_finding(
            findings,
            "external-formula-references",
            "high",
            "Worksheet formulas reference an external workbook",
            {
                "count": len(external_formulas),
                "cells": [f"{item.get('sheet')}!{item.get('cell')}" for item in external_formulas[:25]],
            },
            "Replace external-reference formulas with values or local formulas before creating a self-contained workbook.",
        )

    defined_names = [item for item in as_list(openxml_report.get("definedNames")) if isinstance(item, dict)]
    external_names = [
        item
        for item in defined_names
        if EXTERNAL_FORMULA_RE.search(str(item.get("refersTo", "")))
    ]
    if external_names:
        add_finding(
            findings,
            "external-defined-names",
            "high",
            "Defined names reference an external workbook",
            {"count": len(external_names), "names": [item.get("name", "") for item in external_names[:25]]},
            "Delete unused external names or replace them with local references after confirming no formula depends on them.",
        )

    if openxml_report.get("hasMashupLikeParts"):
        mashup_parts = as_list(openxml_report.get("mashupLikeParts"))
        add_finding(
            findings,
            "mashup-like-parts",
            "medium",
            "Workbook contains Power Query/mashup-like package parts",
            {"count": len(mashup_parts), "parts": mashup_parts[:25]},
            "Export and review Power Query M. For a pure-value deliverable, remove queries, loads, custom XML/mashup parts, and stale connections after refreshing and value-freezing outputs.",
        )

    if openxml_report.get("hasPowerPivotLikeParts"):
        power_pivot_parts = as_list(openxml_report.get("powerPivotLikeParts"))
        add_finding(
            findings,
            "power-pivot-like-parts",
            "medium",
            "Workbook contains Power Pivot/Data Model-like package parts",
            {"count": len(power_pivot_parts), "parts": power_pivot_parts[:25]},
            "Inspect the Data Model through Excel COM when available. For a pure workbook, confirm report cells no longer depend on model measures before removing the model or distributing as values.",
        )

    cube_count = int(openxml_report.get("cubeFormulaCount") or 0)
    if cube_count:
        add_finding(
            findings,
            "cube-formulas",
            "medium",
            "Workbook contains CUBE formulas",
            {"count": cube_count},
            "CUBE formulas depend on a workbook model or OLAP connection. Recalculate and value-freeze them for a static deliverable, or validate connection strings and measure names for a live deliverable.",
        )

    if openxml_report.get("hasVbaProject"):
        add_finding(
            findings,
            "vba-project",
            "low",
            "Workbook contains a VBA project",
            {"parts": as_list(openxml_report.get("vbaParts"))},
            "Export, lint, and compile VBA before macro-enabled delivery. For a pure xlsx deliverable, save a separate non-macro copy after validating formulas and values.",
        )

    max_severity = "info"
    for finding in findings:
        severity = str(finding.get("severity", "info"))
        if SEVERITY_ORDER.get(severity, 0) > SEVERITY_ORDER.get(max_severity, 0):
            max_severity = severity

    if not findings:
        readiness = "clean"
    elif max_severity == "high":
        readiness = "blocked-for-pure-deliverable"
    elif max_severity == "medium":
        readiness = "review-required"
    else:
        readiness = "low-risk"

    summary = {
        "readiness": readiness,
        "maxSeverity": max_severity,
        "findingCount": len(findings),
        "connectionCount": len(connections),
        "credentialLikeConnectionCount": len(
            [
                connection
                for connection in connections
                if connection_secret_indicators(connection)
            ]
        ),
        "externalLinkPartCount": len(external_links),
        "externalFormulaCount": len(external_formulas),
        "externalDefinedNameCount": len(external_names),
        "cubeFormulaCount": cube_count,
        "hasMashupLikeParts": bool(openxml_report.get("hasMashupLikeParts")),
        "hasPowerPivotLikeParts": bool(openxml_report.get("hasPowerPivotLikeParts")),
        "hasVbaProject": bool(openxml_report.get("hasVbaProject")),
    }
    return {
        "workbookPath": openxml_report.get("workbookPath", ""),
        "sourceInspector": "tools/inspect_excel_bi_workbook.py",
        "summary": summary,
        "findings": findings,
        "limitations": [
            "Static OpenXML inspection only; no queries were refreshed and no formulas were calculated.",
            "Connection reachability, credentials, privacy levels, and provider availability require runtime validation.",
            "Use this report to decide whether to value-freeze, remove links/connections, or keep a live workbook with documented prerequisites.",
        ],
    }


def clean_markdown(value: Any) -> str:
    text = str(value)
    return text.replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    lines = [
        "# External Dependency Readiness Report",
        "",
        f"- workbook: `{report.get('workbookPath', '')}`",
        f"- readiness: **{summary.get('readiness', '')}**",
        f"- max severity: `{summary.get('maxSeverity', '')}`",
        f"- findings: {summary.get('findingCount', 0)}",
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
        lines.append("| none | info | no external dependency structures detected | No action required by this static check. |")
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
    args = parser.parse_args()

    openxml_report = json.loads(args.openxml_json.expanduser().read_text(encoding="utf-8"))
    report = build_report(openxml_report)
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(render_markdown(report), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.fail_on_high_risk and report.get("summary", {}).get("maxSeverity") == "high":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
