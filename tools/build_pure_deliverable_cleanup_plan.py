#!/usr/bin/env python3
"""Build a pure-deliverable cleanup plan from an external dependency report.

Input is the JSON produced by ``tools/build_external_dependency_report.py``.
This script is intentionally non-destructive. It does not modify a workbook,
break links, remove connections, or convert formulas. It creates an auditable
plan that an agent can execute on a workbook copy and then verify with OpenXML
and, where available, Windows Excel COM.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


TARGETS = {"pure-xlsx", "pure-xlsm", "live-model"}


def findings_by_code(readiness_report: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for finding in readiness_report.get("findings", []):
        if isinstance(finding, dict) and finding.get("code"):
            result[str(finding["code"])] = finding
    return result


def add_step(
    steps: list[dict[str, Any]],
    action: str,
    phase: str,
    title: str,
    reason: str,
    method: str,
    verification: str,
    required: bool = True,
) -> None:
    steps.append(
        {
            "order": len(steps) + 1,
            "action": action,
            "phase": phase,
            "title": title,
            "required": required,
            "reason": reason,
            "method": method,
            "verification": verification,
        }
    )


def build_cleanup_plan(readiness_report: dict[str, Any], target: str) -> dict[str, Any]:
    if target not in TARGETS:
        raise ValueError(f"Unsupported target: {target}")

    codes = findings_by_code(readiness_report)
    summary = readiness_report.get("summary", {})
    steps: list[dict[str, Any]] = []

    add_step(
        steps,
        "create-working-copy",
        "prepare",
        "Create an isolated delivery copy",
        "Cleanup can remove links, connections, formulas, and package parts; never do this on the source workbook.",
        "Save or copy the source workbook to a deliverable path before editing.",
        "Source workbook remains unchanged and deliverable path is distinct.",
    )

    if target in {"pure-xlsx", "pure-xlsm"} and summary.get("findingCount", 0):
        add_step(
            steps,
            "refresh-and-freeze-values",
            "freeze",
            "Refresh live dependencies and convert delivery outputs to values",
            "Static delivery is only meaningful after formulas, queries, and model-backed cells have their final values.",
            "Use Excel desktop runtime where available: refresh Power Query/Data Model, calculate workbook, then paste report/output ranges as values.",
            "A representative value/range check proves expected outputs survived value conversion.",
        )

    if "external-formula-references" in codes:
        add_step(
            steps,
            "replace-external-formulas",
            "cleanup",
            "Replace formulas that reference external workbooks",
            "External workbook references make the deliverable dependent on unavailable source files.",
            "After value-freezing, replace external-reference formulas with values or local formulas.",
            "Re-run OpenXML inspection and confirm external formula reference count is 0.",
        )

    if "external-defined-names" in codes:
        add_step(
            steps,
            "remove-external-defined-names",
            "cleanup",
            "Remove or localize external defined names",
            "Defined names can keep stale workbook links even when visible formulas look clean.",
            "Delete unused external names or replace them with local ranges after dependency review.",
            "Re-run readiness report and confirm externalDefinedNameCount is 0.",
        )

    if "external-link-parts" in codes:
        add_step(
            steps,
            "remove-external-links",
            "cleanup",
            "Remove Excel external-link package parts",
            "External-link parts can trigger update-link prompts and break self-contained delivery.",
            "Use Excel link management or an OpenXML-aware cleanup flow on the copied workbook after value-freezing.",
            "Re-run OpenXML inspection and confirm externalLinkPartCount is 0.",
        )

    if "workbook-connections" in codes and target in {"pure-xlsx", "pure-xlsm"}:
        add_step(
            steps,
            "remove-workbook-connections",
            "cleanup",
            "Remove workbook data connections",
            "Pure deliverables should not require credentials, providers, privacy levels, or source availability.",
            "Remove WorkbookConnection objects, query loads, and stale connection metadata only after outputs are frozen.",
            "Re-run OpenXML inspection and confirm connectionCount is 0.",
        )

    if "mashup-like-parts" in codes and target in {"pure-xlsx", "pure-xlsm"}:
        add_step(
            steps,
            "remove-power-query-mashup-parts",
            "cleanup",
            "Remove Power Query and mashup-like package remnants",
            "A pure deliverable should not carry query definitions or mashup metadata unless it is intentionally live.",
            "Remove WorkbookQuery objects and query load tables or save a clean value-only copy.",
            "Re-run readiness report and confirm hasMashupLikeParts is false.",
        )

    if "power-pivot-like-parts" in codes and target in {"pure-xlsx", "pure-xlsm"}:
        add_step(
            steps,
            "remove-or-isolate-data-model",
            "cleanup",
            "Remove or isolate Power Pivot/Data Model dependencies",
            "Model parts can keep hidden semantic dependencies even after report cells are frozen.",
            "Confirm no report cells require measures, then save a value-only workbook or remove model artifacts through Excel-supported flows.",
            "Re-run readiness report and confirm hasPowerPivotLikeParts is false for pure delivery.",
        )

    if "cube-formulas" in codes and target in {"pure-xlsx", "pure-xlsm"}:
        add_step(
            steps,
            "replace-cube-formulas",
            "cleanup",
            "Replace CUBE formulas with calculated values",
            "CUBE formulas require a model or OLAP connection and can recalculate differently on another machine.",
            "Calculate in Excel, preserve final displayed values, then replace CUBE formulas in delivery ranges.",
            "Re-run OpenXML inspection and confirm cubeFormulaCount is 0 for the pure deliverable ranges.",
        )

    if "vba-project" in codes:
        if target == "pure-xlsx":
            add_step(
                steps,
                "save-non-macro-copy",
                "cleanup",
                "Save a non-macro `.xlsx` copy",
                "A pure `.xlsx` deliverable must not contain `vbaProject.bin`.",
                "After validating macro-generated outputs, save as `.xlsx` from a copied workbook.",
                "OpenXML inspection confirms hasVbaProject is false and the file extension is `.xlsx`.",
            )
        elif target == "pure-xlsm":
            add_step(
                steps,
                "validate-retained-vba",
                "validate",
                "Validate retained VBA project",
                "Macro-enabled deliverables can keep VBA, but the macros must compile and remain intentional.",
                "Export, lint, import into a copy, compile/run representative macros, and document the macro entry points.",
                "VBA source lint and representative Excel macro run pass.",
            )

    add_step(
        steps,
        "post-clean-openxml-audit",
        "validate",
        "Re-run static workbook inspection and readiness report",
        "Cleanup is not complete until the copied deliverable proves the original risk markers are gone or intentionally retained.",
        "Run inspect_excel_bi_workbook.py and build_external_dependency_report.py on the cleaned deliverable.",
        "For pure targets, readiness is `clean` or only contains documented retained low-risk items.",
    )

    if target == "live-model":
        add_step(
            steps,
            "document-live-prerequisites",
            "validate",
            "Document live workbook prerequisites",
            "Live workbooks can keep connections, but delivery needs reproducible environment requirements.",
            "Record required providers, credentials, privacy levels, source paths, refresh order, and expected row/value checks.",
            "Provider probe and refresh/runtime validation evidence are attached to the delivery package.",
        )

    blockers = [
        step["action"]
        for step in steps
        if step["required"] and step["phase"] in {"freeze", "cleanup", "validate"}
    ]
    status = "ready" if len(steps) == 2 and summary.get("findingCount", 0) == 0 else "cleanup-required"
    if target == "live-model":
        status = "live-validation-required" if summary.get("findingCount", 0) else "ready"

    return {
        "workbookPath": readiness_report.get("workbookPath", ""),
        "target": target,
        "sourceReadiness": summary,
        "status": status,
        "stepCount": len(steps),
        "requiredActionCount": sum(1 for step in steps if step["required"]),
        "blockingActions": blockers,
        "steps": steps,
        "postCleanupAssertions": build_assertions(target),
        "limitations": [
            "This is a cleanup plan only; it does not modify the workbook.",
            "Run destructive cleanup on a copied workbook and verify outputs before replacing a deliverable.",
            "Excel runtime validation is required for refresh, calculation, VBA, Data Model, and provider behavior.",
        ],
    }


def build_assertions(target: str) -> list[dict[str, str]]:
    if target == "live-model":
        return [
            {"check": "providersValidated", "expected": "true"},
            {"check": "refreshEvidenceAttached", "expected": "true"},
            {"check": "livePrerequisitesDocumented", "expected": "true"},
        ]
    assertions = [
        {"check": "connectionCount", "expected": "0"},
        {"check": "externalLinkPartCount", "expected": "0"},
        {"check": "externalFormulaCount", "expected": "0"},
        {"check": "externalDefinedNameCount", "expected": "0"},
        {"check": "hasMashupLikeParts", "expected": "false"},
        {"check": "hasPowerPivotLikeParts", "expected": "false or documented as intentionally retained"},
        {"check": "cubeFormulaCount", "expected": "0 for pure delivery ranges"},
    ]
    if target == "pure-xlsx":
        assertions.append({"check": "hasVbaProject", "expected": "false"})
    return assertions


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def render_markdown(plan: dict[str, Any]) -> str:
    lines = [
        "# Pure Deliverable Cleanup Plan",
        "",
        f"- workbook: `{plan.get('workbookPath', '')}`",
        f"- target: `{plan.get('target', '')}`",
        f"- status: **{plan.get('status', '')}**",
        f"- required actions: {plan.get('requiredActionCount', 0)}",
        "",
        "| # | Phase | Action | Required | Verification |",
        "|---:|---|---|---:|---|",
    ]
    for step in plan.get("steps", []):
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(step.get("order", "")),
                    clean_markdown(step.get("phase", "")),
                    clean_markdown(step.get("action", "")),
                    clean_markdown(step.get("required", "")),
                    clean_markdown(step.get("verification", "")),
                ]
            )
            + " |"
        )
    lines.extend(["", "## Post-Cleanup Assertions", ""])
    for assertion in plan.get("postCleanupAssertions", []):
        lines.append(f"- `{assertion.get('check')}`: {assertion.get('expected')}")
    lines.extend(["", "## Limitations", ""])
    for item in plan.get("limitations", []):
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--readiness-json", required=True, type=Path, help="JSON from build_external_dependency_report.py")
    parser.add_argument("--target", choices=sorted(TARGETS), default="pure-xlsx", help="Delivery target")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable cleanup plan")
    parser.add_argument("--out-md", type=Path, help="Write Markdown cleanup plan")
    parser.add_argument("--fail-if-cleanup-required", action="store_true", help="Exit with code 1 when cleanup is required")
    args = parser.parse_args()

    readiness_report = json.loads(args.readiness_json.expanduser().read_text(encoding="utf-8"))
    plan = build_cleanup_plan(readiness_report, args.target)
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(render_markdown(plan), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(plan, ensure_ascii=False, indent=2))
    if args.fail_if_cleanup_required and plan.get("status") == "cleanup-required":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
