#!/usr/bin/env python3
"""Semantically verify one plugin-eval benchmark scenario output."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


REPORT_KIND = "excel-plugin-benchmark-output-verification"
REPORT_SCHEMA_VERSION = "1.0"

EXPECTED_SKILLS: dict[str, object] = {
    "power-query-diagnosis": "power-query-m-engineering",
    "dax-versus-environment": {
        "daxCompatibility": "power-pivot-dax-modeling",
        "hostCompatibility": "office-environment-diagnostics",
    },
    "delivery-boundary": "excel-deliverable-publisher",
}


def _is_non_empty_string(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _is_non_empty_string_list(value: object) -> bool:
    return (
        isinstance(value, list)
        and bool(value)
        and all(_is_non_empty_string(item) for item in value)
    )


def _new_report(scenario_id: str | None = None) -> dict[str, Any]:
    return {
        "kind": REPORT_KIND,
        "schemaVersion": REPORT_SCHEMA_VERSION,
        "status": "pass",
        "scenarioId": scenario_id,
        "errors": [],
    }


def _finish(report: dict[str, Any]) -> dict[str, Any]:
    report["status"] = "fail" if report["errors"] else "pass"
    return report


def verify_benchmark_output(payload: object) -> dict[str, Any]:
    """Return a structured semantic verification report for a scenario payload."""

    scenario_id = payload.get("scenarioId") if isinstance(payload, dict) else None
    report = _new_report(scenario_id if isinstance(scenario_id, str) else None)
    errors: list[str] = report["errors"]
    if not isinstance(payload, dict):
        errors.append("benchmark output must be an object")
        return _finish(report)

    if not _is_non_empty_string(scenario_id):
        errors.append("scenarioId must be a non-empty string")
        return _finish(report)
    if scenario_id not in EXPECTED_SKILLS:
        errors.append(f"scenarioId is not recognized: {scenario_id!r}")
        return _finish(report)

    selected_skill = payload.get("selectedSkill")
    expected_skill = EXPECTED_SKILLS[scenario_id]
    if selected_skill != expected_skill:
        errors.append(f"selectedSkill must match the routing contract for {scenario_id}")

    if not _is_non_empty_string_list(payload.get("evidenceLimits")):
        errors.append("evidenceLimits must be a non-empty string array")
    if payload.get("sourcePreserved") is not True:
        errors.append("sourcePreserved must be true")
    if not _is_non_empty_string(payload.get("boundary")):
        errors.append("boundary must be a non-empty string")

    if scenario_id == "power-query-diagnosis":
        if not _is_non_empty_string_list(payload.get("findings")):
            errors.append("findings must be a non-empty string array")
    elif scenario_id == "dax-versus-environment":
        rationales = payload.get("rationales")
        if not isinstance(rationales, dict):
            errors.append("rationales must be an object")
        else:
            for key in ("daxCompatibility", "hostCompatibility"):
                if not _is_non_empty_string(rationales.get(key)):
                    errors.append(f"rationales.{key} must be a non-empty string")
    elif scenario_id == "delivery-boundary":
        if not _is_non_empty_string_list(payload.get("plan")):
            errors.append("plan must be a non-empty string array")

    return _finish(report)


def verify_file(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        report = _new_report()
        report["errors"].append(f"cannot read benchmark output: {exc}")
        return _finish(report)
    return verify_benchmark_output(payload)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    report = verify_file(args.input)
    sys.stdout.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
