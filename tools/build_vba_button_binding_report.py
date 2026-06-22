#!/usr/bin/env python3
"""Build a report that maps Excel shape/button OnAction values to VBA macros.

Inputs are JSON reports from:

- .agents/skills/excel-vba-workbook-engineering/scripts/inspect_workbook.ps1
- .agents/skills/excel-vba-workbook-engineering/scripts/lint_vba_source.py

The report is static. It does not click controls, compile VBA, or run macros.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


WORKBOOK_MACRO_RE = re.compile(r"^'[^']+'!(.+)$")
BARE_WORKBOOK_MACRO_RE = re.compile(r"^[^!]+!(.+)$")
IDENTIFIER_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*$")


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def normalize_on_action(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    match = WORKBOOK_MACRO_RE.match(text)
    if match:
        text = match.group(1).strip()
    else:
        match = BARE_WORKBOOK_MACRO_RE.match(text)
        if match:
            text = match.group(1).strip()
    if " " in text:
        text = text.split(" ", 1)[0].strip()
    if "." in text:
        text = text.rsplit(".", 1)[-1].strip()
    identifier_match = IDENTIFIER_RE.search(text)
    return identifier_match.group(0) if identifier_match else text


def public_macro_index(vba_lint: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = {}
    for entry in vba_lint.get("publicEntries", []):
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("name", "")).strip()
        if not name:
            continue
        result.setdefault(name.lower(), []).append(entry)
    return result


def iter_shape_actions(workbook_inventory: dict[str, Any]) -> list[dict[str, Any]]:
    actions: list[dict[str, Any]] = []
    for sheet in workbook_inventory.get("worksheets", []):
        if not isinstance(sheet, dict):
            continue
        sheet_name = str(sheet.get("name", ""))
        for shape in sheet.get("shapes", []):
            if not isinstance(shape, dict):
                continue
            on_action = str(shape.get("onAction", "") or "").strip()
            if not on_action:
                continue
            normalized = normalize_on_action(on_action)
            actions.append(
                {
                    "sheet": sheet_name,
                    "shape": str(shape.get("name", "")),
                    "text": str(shape.get("text", "")),
                    "rawOnAction": on_action,
                    "normalizedMacro": normalized,
                }
            )
    return actions


def build_report(workbook_inventory: dict[str, Any], vba_lint: dict[str, Any]) -> dict[str, Any]:
    macros = public_macro_index(vba_lint)
    bindings: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []
    for action in iter_shape_actions(workbook_inventory):
        normalized = str(action.get("normalizedMacro", ""))
        matches = macros.get(normalized.lower(), []) if normalized else []
        status = "resolved" if matches else "missing-macro"
        binding = {
            **action,
            "status": status,
            "matchedEntries": matches,
        }
        bindings.append(binding)
        if status != "resolved":
            findings.append(
                {
                    "severity": "error",
                    "code": "missing-onaction-macro",
                    "sheet": action.get("sheet"),
                    "shape": action.get("shape"),
                    "rawOnAction": action.get("rawOnAction"),
                    "normalizedMacro": normalized,
                    "message": "Shape OnAction does not match a public standard-module Sub in the VBA lint report.",
                }
            )

    return {
        "status": "pass" if not findings else "fail",
        "workbook": {
            "path": workbook_inventory.get("workbookPath", ""),
            "name": workbook_inventory.get("name", ""),
            "hasVBProject": workbook_inventory.get("hasVBProject"),
        },
        "summary": {
            "shapeActionCount": len(bindings),
            "resolvedCount": sum(1 for item in bindings if item.get("status") == "resolved"),
            "missingMacroCount": sum(1 for item in bindings if item.get("status") == "missing-macro"),
            "publicEntryCount": len(vba_lint.get("publicEntries", [])),
        },
        "bindings": bindings,
        "findings": findings,
        "boundaries": [
            "This report is static and uses workbook inventory plus VBA lint JSON.",
            "It does not click buttons, compile VBA, evaluate workbook references, or run macros.",
            "Use Windows Excel COM for final macro execution validation when a workbook deliverable depends on buttons.",
        ],
    }


def render_markdown(report: dict[str, Any]) -> str:
    summary = report["summary"]
    lines = [
        "# VBA Button Binding Report",
        "",
        f"- status: **{report['status']}**",
        f"- shape actions: `{summary['shapeActionCount']}`",
        f"- resolved: `{summary['resolvedCount']}`",
        f"- missing macros: `{summary['missingMacroCount']}`",
        f"- public entry macros: `{summary['publicEntryCount']}`",
        "",
        "| Sheet | Shape | OnAction | Macro | Status |",
        "|---|---|---|---|---:|",
    ]
    for item in report["bindings"]:
        lines.append(
            "| {sheet} | {shape} | `{raw}` | `{macro}` | {status} |".format(
                sheet=item.get("sheet", ""),
                shape=item.get("shape", ""),
                raw=str(item.get("rawOnAction", "")).replace("|", "\\|"),
                macro=str(item.get("normalizedMacro", "")).replace("|", "\\|"),
                status=item.get("status", ""),
            )
        )
    lines.extend(["", "## Findings", ""])
    if report["findings"]:
        for finding in report["findings"]:
            lines.append(
                "- `{code}` on `{sheet}!{shape}`: `{macro}`".format(
                    code=finding.get("code"),
                    sheet=finding.get("sheet"),
                    shape=finding.get("shape"),
                    macro=finding.get("normalizedMacro"),
                )
            )
    else:
        lines.append("No unresolved button bindings found.")
    lines.extend(["", "## Boundaries", ""])
    for item in report["boundaries"]:
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook-inventory-json", required=True, type=Path, help="inspect_workbook.ps1 JSON report")
    parser.add_argument("--vba-lint-json", required=True, type=Path, help="lint_vba_source.py JSON report")
    parser.add_argument("--out-json", type=Path, help="Write JSON report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--fail-on-unresolved", action="store_true", help="Exit non-zero when unresolved bindings exist")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = build_report(load_json(args.workbook_inventory_json), load_json(args.vba_lint_json))
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(render_markdown(report), encoding="utf-8")
    summary = report["summary"]
    print(
        "VBA button binding {status}: actions={actions}, resolved={resolved}, missing={missing}".format(
            status=report["status"],
            actions=summary["shapeActionCount"],
            resolved=summary["resolvedCount"],
            missing=summary["missingMacroCount"],
        )
    )
    if args.fail_on_unresolved and report["status"] != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
