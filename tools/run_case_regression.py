#!/usr/bin/env python3
"""Validate the real/sanitized Excel BI regression case library.

The case library is intentionally manifest-driven. It records repeatable
problem shapes learned from real work without shipping customer workbooks.
This runner validates the schema, coverage, boundaries, and referenced package
tools. It does not open private workbooks or claim live Excel runtime behavior.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REQUIRED_LAYERS = {
    "power-query",
    "dax",
    "cube-mdx",
    "vba",
    "deliverable",
    "visual-qa",
}

REQUIRED_CASE_FIELDS = {
    "id",
    "title",
    "layer",
    "evidenceMode",
    "originSignal",
    "riskPattern",
    "regressionChecks",
    "packageTools",
    "expectedEvidence",
    "boundaries",
    "nextLiveWorkbookRequirement",
}

ALLOWED_EVIDENCE_MODES = {"sanitized-spec", "fixture-backed", "live-workbook-required"}

FORBIDDEN_MARKERS = [
    "WX" + "Work",
    "codex" + "-" + "clipboard",
    "App" + "Data" + "\\Local\\Temp",
    "Temp" + "State" + "\\ScreenClip",
]


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_string_list(value: Any, *, min_items: int = 1) -> bool:
    return isinstance(value, list) and len(value) >= min_items and all(isinstance(item, str) and item.strip() for item in value)


def contains_forbidden_marker(value: Any) -> str:
    text = json.dumps(value, ensure_ascii=False)
    for marker in FORBIDDEN_MARKERS:
        if marker.lower() in text.lower():
            return marker
    return ""


def validate_case(project_root: Path, case_root: Path, case_ref: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    errors: list[str] = []
    case_id = str(case_ref.get("id", "")).strip()
    spec_path_text = str(case_ref.get("specPath", "")).strip()
    if not case_id:
        errors.append("case reference missing id")
    if not spec_path_text:
        errors.append(f"{case_id or '<unknown>'}: case reference missing specPath")
        return {"id": case_id}, errors

    spec_path = (case_root / spec_path_text).resolve()
    if not spec_path.is_file():
        errors.append(f"{case_id}: missing spec file {spec_path_text}")
        return {"id": case_id, "specPath": spec_path_text}, errors

    try:
        spec = read_json(spec_path)
    except Exception as exc:
        errors.append(f"{case_id}: cannot read spec JSON: {exc}")
        return {"id": case_id, "specPath": spec_path_text}, errors

    missing = sorted(REQUIRED_CASE_FIELDS - set(spec))
    if missing:
        errors.append(f"{case_id}: missing required fields: {', '.join(missing)}")
    if spec.get("id") != case_id:
        errors.append(f"{case_id}: spec id does not match manifest id {spec.get('id')!r}")

    layer = str(spec.get("layer", "")).strip()
    if layer not in REQUIRED_LAYERS:
        errors.append(f"{case_id}: unsupported layer {layer!r}")

    evidence_mode = str(spec.get("evidenceMode", "")).strip()
    if evidence_mode not in ALLOWED_EVIDENCE_MODES:
        errors.append(f"{case_id}: unsupported evidenceMode {evidence_mode!r}")

    for field_name in ["regressionChecks", "packageTools", "expectedEvidence", "boundaries"]:
        if not validate_string_list(spec.get(field_name)):
            errors.append(f"{case_id}: {field_name} must be a non-empty string list")

    for tool in spec.get("packageTools", []) if isinstance(spec.get("packageTools"), list) else []:
        tool_path = (project_root / tool).resolve()
        if not tool_path.exists():
            errors.append(f"{case_id}: referenced package tool does not exist: {tool}")

    forbidden = contains_forbidden_marker(spec)
    if forbidden:
        errors.append(f"{case_id}: forbidden local/customer marker found: {forbidden}")

    return spec, errors


def validate(project_root: Path, case_root: Path) -> dict[str, Any]:
    manifest_path = case_root / "manifest.json"
    errors: list[str] = []
    warnings: list[str] = []
    if not manifest_path.is_file():
        return {
            "status": "fail",
            "errors": [f"missing manifest: {manifest_path}"],
            "warnings": [],
            "cases": [],
        }

    manifest = read_json(manifest_path)
    cases_ref = manifest.get("cases")
    if not isinstance(cases_ref, list) or not cases_ref:
        errors.append("manifest.cases must be a non-empty list")
        cases_ref = []

    for required in ["schemaVersion", "purpose", "boundaries", "requiredLayers"]:
        if required not in manifest:
            errors.append(f"manifest missing {required}")

    if not validate_string_list(manifest.get("boundaries")):
        errors.append("manifest.boundaries must be a non-empty string list")

    manifest_layers = set(manifest.get("requiredLayers") or [])
    missing_manifest_layers = REQUIRED_LAYERS - manifest_layers
    if missing_manifest_layers:
        errors.append(f"manifest.requiredLayers missing: {', '.join(sorted(missing_manifest_layers))}")

    specs: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    for case_ref in cases_ref:
        if not isinstance(case_ref, dict):
            errors.append("manifest case reference must be an object")
            continue
        case_id = str(case_ref.get("id", "")).strip()
        if case_id in seen_ids:
            errors.append(f"duplicate case id: {case_id}")
        seen_ids.add(case_id)
        spec, case_errors = validate_case(project_root, case_root, case_ref)
        specs.append(spec)
        errors.extend(case_errors)

    covered_layers = {str(spec.get("layer", "")).strip() for spec in specs}
    missing_layers = REQUIRED_LAYERS - covered_layers
    if missing_layers:
        errors.append(f"case coverage missing layers: {', '.join(sorted(missing_layers))}")

    evidence_modes = sorted({str(spec.get("evidenceMode", "")).strip() for spec in specs if spec.get("evidenceMode")})
    if "live-workbook-required" not in evidence_modes:
        warnings.append("no live-workbook-required cases yet; add this when a real sanitized workbook is available")

    return {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "status": "pass" if not errors else "fail",
        "projectRoot": str(project_root),
        "caseRoot": str(case_root),
        "caseCount": len(specs),
        "requiredLayers": sorted(REQUIRED_LAYERS),
        "coveredLayers": sorted(covered_layers),
        "evidenceModes": evidence_modes,
        "cases": [
            {
                "id": spec.get("id", ""),
                "title": spec.get("title", ""),
                "layer": spec.get("layer", ""),
                "evidenceMode": spec.get("evidenceMode", ""),
                "checkCount": len(spec.get("regressionChecks", []) if isinstance(spec.get("regressionChecks"), list) else []),
                "toolCount": len(spec.get("packageTools", []) if isinstance(spec.get("packageTools"), list) else []),
            }
            for spec in specs
        ],
        "errors": errors,
        "warnings": warnings,
    }


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Excel BI Real/Sanitized Case Regression Report",
        "",
        f"- status: **{report.get('status')}**",
        f"- case root: `{report.get('caseRoot')}`",
        f"- case count: `{report.get('caseCount')}`",
        f"- covered layers: `{', '.join(report.get('coveredLayers', []))}`",
        f"- evidence modes: `{', '.join(report.get('evidenceModes', []))}`",
        "",
        "## Cases",
        "",
        "| Case | Layer | Evidence Mode | Checks | Tools |",
        "|---|---|---|---:|---:|",
    ]
    for case in report.get("cases", []):
        lines.append(
            f"| `{case.get('id')}` | {case.get('layer')} | {case.get('evidenceMode')} | {case.get('checkCount')} | {case.get('toolCount')} |"
        )
    lines.extend(["", "## Boundaries", ""])
    lines.append("- This validates regression definitions and package-tool references, not private customer workbooks.")
    lines.append("- Live refresh, VBA execution, Data Model evaluation, and visual rendering require separate runtime evidence.")
    if report.get("warnings"):
        lines.extend(["", "## Warnings", ""])
        for warning in report["warnings"]:
            lines.append(f"- {warning}")
    if report.get("errors"):
        lines.extend(["", "## Errors", ""])
        for error in report["errors"]:
            lines.append(f"- {error}")
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--case-root", default="fixtures/real-sanitized-cases", help="Regression case library root")
    parser.add_argument("--out-json", default="", help="Write JSON report")
    parser.add_argument("--out-md", default="", help="Write Markdown report")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero unless the report status is pass")
    args = parser.parse_args()

    project_root = Path(args.project_root).expanduser().resolve()
    case_root = (project_root / args.case_root).resolve() if not Path(args.case_root).is_absolute() else Path(args.case_root).resolve()
    report = validate(project_root, case_root)

    if args.out_json:
        write_json(Path(args.out_json).expanduser().resolve(), report)
    if args.out_md:
        out_md = Path(args.out_md).expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")

    if report["status"] == "pass":
        print(
            "Case regression validation OK: "
            f"{report['caseCount']} cases, layers={','.join(report['coveredLayers'])}"
        )
        return 0

    print("Case regression validation failed:", file=sys.stderr)
    for error in report["errors"]:
        print(f"- {error}", file=sys.stderr)
    return 1 if args.require_pass else 0


if __name__ == "__main__":
    raise SystemExit(main())
