#!/usr/bin/env python3
"""Build a cross-platform compatibility report from an Excel capability probe."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"
VALID_CAPABILITY_STATUSES = {PASS, FAIL, "skip", "error"}
VALID_EVIDENCE_LEVELS = {"registration", "activation", "smoke", "not-tested"}
RFC3339_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$"
)

CAPABILITY_IDS = (
    "excel.com.activation",
    "excel.workbook.roundtrip",
    "excel.vba.project-access",
    "excel.power-query.object-model",
    "excel.power-query.async-wait",
    "excel.data-model.object-model",
    "excel.pdf-export",
    "ado.com.activation",
    "ace.workbook-sql",
    "msolap.registration",
    "adomd.com.activation",
)

OPERATION_REQUIREMENTS = {
    "workbook-automation": ["excel.com.activation", "excel.workbook.roundtrip"],
    "vba-engineering": ["excel.com.activation", "excel.workbook.roundtrip", "excel.vba.project-access"],
    "power-query-automation": [
        "excel.com.activation",
        "excel.workbook.roundtrip",
        "excel.power-query.object-model",
        "excel.power-query.async-wait",
    ],
    "data-model-inspection": ["excel.com.activation", "excel.data-model.object-model"],
    "ado-workbook-sql": ["ado.com.activation", "ace.workbook-sql"],
    "adomd-endpoint-query": ["msolap.registration", "adomd.com.activation"],
    "rendered-pdf-evidence": ["excel.com.activation", "excel.workbook.roundtrip", "excel.pdf-export"],
}

TOP_LEVEL_FIELDS = {
    "schemaVersion",
    "kind",
    "generatedAt",
    "probe",
    "environment",
    "capabilities",
    "boundaries",
    "errors",
}
PROBE_FIELDS = {"profile", "platform", "syntheticFixture"}
ENVIRONMENT_FIELDS = {
    "osVersion",
    "is64BitOperatingSystem",
    "is64BitProcess",
    "powershellVersion",
    "excelVersion",
    "excelBuild",
}
CAPABILITY_FIELDS = {"status", "evidenceLevel", "detail", "errorCategory", "error"}


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def missing_capability(capability_id: str) -> dict[str, str]:
    return {
        "status": "skip",
        "evidenceLevel": "not-tested",
        "detail": f"Capability evidence is absent from the probe: {capability_id}",
        "errorCategory": "missing-evidence",
        "error": "",
    }


def normalize_capability(capability_id: str, value: Any, errors: list[str]) -> dict[str, str]:
    if value is None:
        return missing_capability(capability_id)
    if not isinstance(value, dict):
        errors.append(f"capabilities.{capability_id} must be an object")
        return {
            "status": "error",
            "evidenceLevel": "not-tested",
            "detail": "Malformed capability evidence.",
            "errorCategory": "invalid-contract",
            "error": "capability row is not an object",
        }

    status = str(value.get("status", ""))
    evidence_level = str(value.get("evidenceLevel", ""))
    if status not in VALID_CAPABILITY_STATUSES:
        errors.append(f"capabilities.{capability_id}.status is invalid: {status!r}")
        status = "error"
    if evidence_level not in VALID_EVIDENCE_LEVELS:
        errors.append(f"capabilities.{capability_id}.evidenceLevel is invalid: {evidence_level!r}")
        evidence_level = "not-tested"
    detail = str(value.get("detail", "") or "")
    if not detail:
        detail = "Capability detail is unavailable."
    return {
        "status": status,
        "evidenceLevel": evidence_level,
        "detail": detail,
        "errorCategory": str(value.get("errorCategory", "") or ""),
        "error": str(value.get("error", "") or ""),
    }


def validate_exact_fields(
    value: dict[str, Any],
    *,
    path: str,
    required: set[str],
    unknown_label: str,
) -> list[str]:
    errors: list[str] = []
    for field in sorted(required - set(value)):
        errors.append(f"{path}.{field} is required" if path else f"{field} is required")
    for field in sorted(set(value) - required):
        errors.append(f"unknown {unknown_label} field: {field}")
    return errors


def is_rfc3339(value: Any) -> bool:
    if not isinstance(value, str) or not RFC3339_RE.fullmatch(value):
        return False
    normalized = value[:-1] + "+00:00" if value.endswith("Z") else value
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return False
    return parsed.tzinfo is not None


def validate_probe_contract(probe: Any) -> list[str]:
    if not isinstance(probe, dict):
        return ["probe JSON root must be an object"]
    errors = validate_exact_fields(
        probe,
        path="",
        required=TOP_LEVEL_FIELDS,
        unknown_label="top-level",
    )
    if probe.get("schemaVersion") != "1.0":
        errors.append("schemaVersion must be '1.0'")
    if probe.get("kind") != "excel-capability-probe":
        errors.append("kind must be 'excel-capability-probe'")
    if not is_rfc3339(probe.get("generatedAt")):
        errors.append("generatedAt must be RFC3339 with an explicit timezone")

    probe_meta = probe.get("probe")
    if not isinstance(probe_meta, dict):
        errors.append("probe must be an object")
    else:
        errors.extend(
            validate_exact_fields(
                probe_meta,
                path="probe",
                required=PROBE_FIELDS,
                unknown_label="probe",
            )
        )
        if probe_meta.get("profile") not in {"inventory", "runtime"}:
            errors.append("probe.profile must be 'inventory' or 'runtime'")
        if probe_meta.get("platform") != "windows":
            errors.append("probe.platform must be 'windows'")
        if type(probe_meta.get("syntheticFixture")) is not bool:
            errors.append("probe.syntheticFixture must be a boolean")

    environment = probe.get("environment")
    if not isinstance(environment, dict):
        errors.append("environment must be an object")
    else:
        errors.extend(
            validate_exact_fields(
                environment,
                path="environment",
                required=ENVIRONMENT_FIELDS,
                unknown_label="environment",
            )
        )
        for field in ["osVersion", "powershellVersion", "excelVersion", "excelBuild"]:
            if field in environment and not isinstance(environment[field], str):
                errors.append(f"environment.{field} must be a string")
        for field in ["is64BitOperatingSystem", "is64BitProcess"]:
            if field in environment and type(environment[field]) is not bool:
                errors.append(f"environment.{field} must be a boolean")

    raw_capabilities = probe.get("capabilities")
    if not isinstance(raw_capabilities, dict):
        errors.append("capabilities must be an object")
    else:
        for capability_id in sorted(set(raw_capabilities) - set(CAPABILITY_IDS)):
            errors.append(f"unknown capability ID: {capability_id}")
        for capability_id, row in raw_capabilities.items():
            if capability_id not in CAPABILITY_IDS:
                continue
            if not isinstance(row, dict):
                errors.append(f"capabilities.{capability_id} must be an object")
                continue
            for field in sorted(CAPABILITY_FIELDS - set(row)):
                errors.append(f"capabilities.{capability_id}.{field} is required")
            for field in sorted(set(row) - CAPABILITY_FIELDS):
                errors.append(f"unknown capability field: {capability_id}.{field}")
            for field in sorted(CAPABILITY_FIELDS & set(row)):
                if not isinstance(row[field], str):
                    errors.append(f"capabilities.{capability_id}.{field} must be a string")
            if row.get("status") not in VALID_CAPABILITY_STATUSES:
                errors.append(f"capabilities.{capability_id}.status is invalid: {row.get('status')!r}")
            if row.get("evidenceLevel") not in VALID_EVIDENCE_LEVELS:
                errors.append(
                    f"capabilities.{capability_id}.evidenceLevel is invalid: {row.get('evidenceLevel')!r}"
                )
            if isinstance(row.get("detail"), str) and not row["detail"]:
                errors.append(f"capabilities.{capability_id}.detail must not be empty")

    boundaries = probe.get("boundaries")
    if not isinstance(boundaries, list):
        errors.append("boundaries must be a list")
    elif not boundaries:
        errors.append("boundaries must not be empty")
    elif not all(isinstance(item, str) and item for item in boundaries):
        errors.append("boundaries items must be non-empty strings")

    probe_errors = probe.get("errors")
    if not isinstance(probe_errors, list):
        errors.append("errors must be a list")
    elif not all(isinstance(item, str) and item for item in probe_errors):
        errors.append("errors items must be non-empty strings")
    return errors


def operation_readiness(operation_id: str, required: list[str], capabilities: dict[str, dict[str, str]]) -> dict[str, Any]:
    missing = [item for item in required if capabilities[item]["status"] != PASS]
    statuses = {capabilities[item]["status"] for item in required}
    if not missing:
        readiness = "requires-user-input" if operation_id == "adomd-endpoint-query" else "ready"
    elif FAIL in statuses:
        readiness = "blocked"
    else:
        readiness = "unknown"
    return {"readiness": readiness, "requires": required, "missing": missing}


def build_report(
    probe: Any,
    required_capabilities: list[str] | None = None,
    *,
    input_errors: list[str] | None = None,
) -> dict[str, Any]:
    required = list(dict.fromkeys(required_capabilities or []))
    errors = list(input_errors or [])
    errors.extend(validate_probe_contract(probe))
    probe_dict = probe if isinstance(probe, dict) else {}
    raw_capabilities = probe_dict.get("capabilities", {})
    if not isinstance(raw_capabilities, dict):
        raw_capabilities = {}

    capabilities = {
        capability_id: normalize_capability(capability_id, raw_capabilities.get(capability_id), errors)
        for capability_id in CAPABILITY_IDS
    }

    unknown_required = [item for item in required if item not in CAPABILITY_IDS]
    for capability_id in unknown_required:
        errors.append(f"unknown required capability: {capability_id}")
    unmet = [
        item
        for item in required
        if item not in capabilities or capabilities[item]["status"] != PASS
    ]

    counts = {
        status: sum(1 for item in capabilities.values() if item["status"] == status)
        for status in [PASS, FAIL, "skip", "error"]
    }
    operations = {
        operation_id: operation_readiness(operation_id, operation_required, capabilities)
        for operation_id, operation_required in OPERATION_REQUIREMENTS.items()
    }
    probe_meta = probe_dict.get("probe", {}) if isinstance(probe_dict.get("probe"), dict) else {}
    source_profile = str(probe_meta.get("profile", "") or "")
    if source_profile not in {"inventory", "runtime"}:
        source_profile = ""
    source_errors = probe_dict.get("errors", []) if isinstance(probe_dict.get("errors"), list) else []
    if source_errors:
        errors.append(
            f"incomplete probe evidence: probe reported {len(source_errors)} execution or cleanup error(s)"
        )

    return {
        "schemaVersion": "1.0",
        "kind": "excel-compatibility-report",
        "status": PASS if not errors and not unmet else FAIL,
        "generatedAt": now_iso(),
        "source": {
            "probeJson": "",
            "probeSchemaVersion": str(probe_dict.get("schemaVersion", "") or ""),
            "profile": source_profile,
        },
        "environment": probe_dict.get("environment", {}) if isinstance(probe_dict.get("environment"), dict) else {},
        "summary": {
            "passCount": counts[PASS],
            "failCount": counts[FAIL],
            "skipCount": counts["skip"],
            "errorCount": counts["error"],
            "requiredCount": len(required),
            "unmetRequiredCount": len(unmet),
        },
        "capabilities": capabilities,
        "operations": operations,
        "requirements": {"requested": required, "unmet": unmet},
        "boundaries": [
            "Missing or skipped capability evidence is unknown, not proof of support.",
            "Synthetic fixture evidence validates report behavior only, not local Office readiness.",
            "COM activation does not prove a real workbook, Data Model, or external endpoint query.",
            "ADOMD endpoint queries still require an explicit connection string and query.",
        ],
        "probeErrors": [str(item) for item in source_errors],
        "errors": errors,
    }


def clean_markdown(value: Any) -> str:
    return str(value if value is not None else "").replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict[str, Any]) -> str:
    summary = report["summary"]
    lines = [
        "# Excel Compatibility Report",
        "",
        f"- status: **{report['status']}**",
        f"- probe profile: `{report['source'].get('profile', '')}`",
        f"- capability evidence: pass `{summary['passCount']}`, fail `{summary['failCount']}`, skip `{summary['skipCount']}`, error `{summary['errorCount']}`",
        f"- required capabilities: `{summary['requiredCount']}`; unmet: `{summary['unmetRequiredCount']}`",
        "",
        "## Capabilities",
        "",
        "| Capability | Status | Evidence | Detail |",
        "|---|---:|---|---|",
    ]
    for capability_id in CAPABILITY_IDS:
        item = report["capabilities"][capability_id]
        lines.append(
            f"| `{capability_id}` | {item['status']} | {item['evidenceLevel']} | {clean_markdown(item['detail'])} |"
        )
    lines.extend(["", "## Operations", "", "| Operation | Readiness | Missing evidence |", "|---|---:|---|"])
    for operation_id, item in report["operations"].items():
        lines.append(f"| `{operation_id}` | {item['readiness']} | {clean_markdown(', '.join(item['missing']))} |")
    lines.extend(["", "## Boundaries", ""])
    lines.extend(f"- {item}" for item in report["boundaries"])
    lines.extend(["", "## Errors", ""])
    if report["errors"]:
        lines.extend(f"- {item}" for item in report["errors"])
    else:
        lines.append("- none")
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--probe-json", required=True, type=Path, help="JSON from probe_excel_capabilities.ps1")
    parser.add_argument("--require-capability", action="append", default=[], help="Capability ID that must have pass evidence")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--print", action="store_true", dest="print_report", help="Print Markdown report")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero unless report status is pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    probe_path = args.probe_json.expanduser().resolve()
    input_errors: list[str] = []
    try:
        probe: Any = json.loads(probe_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        probe = {}
        input_errors.append(f"cannot read probe JSON: {exc}")
    report = build_report(probe, args.require_capability, input_errors=input_errors)
    report["source"]["probeJson"] = str(probe_path)
    markdown = render_markdown(report)
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), report)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown, encoding="utf-8")
    if args.print_report:
        print(markdown)
    elif not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["status"] == PASS:
        print(
            f"Excel compatibility pass: capabilities={report['summary']['passCount']}/{len(CAPABILITY_IDS)}, "
            f"unmetRequired={report['summary']['unmetRequiredCount']}"
        )
    else:
        print("Excel compatibility report failed", file=sys.stderr)
    return 1 if args.require_pass and report["status"] != PASS else 0


if __name__ == "__main__":
    raise SystemExit(main())
