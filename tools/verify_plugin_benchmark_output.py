#!/usr/bin/env python3
"""Semantically verify one plugin-eval benchmark scenario output."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


REPORT_KIND = "excel-plugin-benchmark-output-verification"
REPORT_SCHEMA_VERSION = "1.0"
SOURCE_ARTIFACT = Path("benchmarks/fixtures/excel-bi-benchmark-source.json")

EXPECTED_SKILLS: dict[str, object] = {
    "power-query-diagnosis": "power-query-m-engineering",
    "dax-versus-environment": {
        "daxCompatibility": "power-pivot-dax-modeling",
        "hostCompatibility": "office-environment-diagnostics",
    },
    "delivery-boundary": "excel-deliverable-publisher",
}
UNSUPPORTED_CLAIM_PATTERNS = (
    re.compile(
        r"\breal(?:\s+customer)?\s+workbook\b.*\b(?:opened|verified|validated|success(?:ful(?:ly)?)?)\b",
        re.IGNORECASE | re.DOTALL,
    ),
    re.compile(
        r"\b(?:opened|verified|validated|success(?:ful(?:ly)?)?)\b.*\breal(?:\s+customer)?\s+workbook\b",
        re.IGNORECASE | re.DOTALL,
    ),
    re.compile(r"\breal\s+runtime\s+success\b", re.IGNORECASE),
    re.compile(
        r"\b(?:live|real)\s+(?:excel\s+)?runtime\b.*\b(?:executed|established|success|successful|successfully|verified|validated)\b",
        re.IGNORECASE | re.DOTALL,
    ),
    re.compile(
        r"\blive\s+(?:excel|runtime|workbook)\b.*\b(?:executed|established|success|successful|successfully|verified|validated)\b",
        re.IGNORECASE | re.DOTALL,
    ),
)
SYNTHETIC_NO_LIVE_PROOF = re.compile(
    r"\bsynthetic\b.*\bno\s+live(?:\s+\w+){0,2}\s+runtime\s+proof\b",
    re.IGNORECASE | re.DOTALL,
)


def _is_non_empty_string(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _is_non_empty_string_list(value: object) -> bool:
    return (
        isinstance(value, list)
        and bool(value)
        and all(_is_non_empty_string(item) for item in value)
    )


def _new_report(
    scenario_id: str | None = None, expected_scenario_id: str | None = None
) -> dict[str, Any]:
    return {
        "kind": REPORT_KIND,
        "schemaVersion": REPORT_SCHEMA_VERSION,
        "status": "pass",
        "scenarioId": scenario_id,
        "expectedScenarioId": expected_scenario_id,
        "errors": [],
    }


def _finish(report: dict[str, Any]) -> dict[str, Any]:
    report["status"] = "fail" if report["errors"] else "pass"
    return report


def _infer_expected_scenario(workspace: Path) -> str | None:
    parent_name = workspace.resolve().parent.name
    for scenario_id in EXPECTED_SKILLS:
        prefix = f"plugin-eval-{scenario_id}-"
        if parent_name.startswith(prefix) and len(parent_name) > len(prefix):
            return scenario_id
    return None


def _iter_output_strings(value: object, path: str = "$"):
    if isinstance(value, str):
        yield path, value
    elif isinstance(value, dict):
        for key, item in value.items():
            if isinstance(key, str):
                yield f"{path}.<key>", key
            yield from _iter_output_strings(item, f"{path}.{key}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            yield from _iter_output_strings(item, f"{path}[{index}]")


def _find_git_executable() -> str | None:
    discovered = shutil.which("git")
    if discovered:
        return discovered
    candidates = []
    for variable in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(variable)
        if base:
            candidates.append(Path(base) / "Git" / "cmd" / "git.exe")
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def _source_artifact_errors(workspace: Path) -> list[str]:
    artifact_path = workspace / SOURCE_ARTIFACT
    try:
        current_bytes = artifact_path.read_bytes()
    except OSError as exc:
        return [f"source artifact is unavailable: {exc}"]

    git_executable = _find_git_executable()
    if not git_executable:
        return ["git baseline is unavailable for the source artifact"]
    try:
        completed = subprocess.run(
            [
                git_executable,
                "-C",
                str(workspace),
                "show",
                f"HEAD:{SOURCE_ARTIFACT.as_posix()}",
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
    except OSError as exc:
        return [f"git baseline is unavailable for the source artifact: {exc}"]
    if completed.returncode != 0:
        return ["git baseline is unavailable for the source artifact"]
    if current_bytes != completed.stdout:
        return ["source artifact bytes differ from the git HEAD baseline"]
    return []


def verify_benchmark_output(
    payload: object, *, expected_scenario_id: str | None
) -> dict[str, Any]:
    """Return a structured semantic verification report for a scenario payload."""

    scenario_id = payload.get("scenarioId") if isinstance(payload, dict) else None
    report = _new_report(
        scenario_id if isinstance(scenario_id, str) else None,
        expected_scenario_id,
    )
    errors: list[str] = report["errors"]
    if not isinstance(payload, dict):
        errors.append("benchmark output must be an object")
        return _finish(report)

    if not _is_non_empty_string(scenario_id):
        errors.append("scenarioId must be a non-empty string")
        return _finish(report)
    if expected_scenario_id not in EXPECTED_SKILLS:
        errors.append("expected scenario cannot be inferred from the workspace parent directory")
        return _finish(report)
    if scenario_id != expected_scenario_id:
        errors.append(
            f"scenarioId must match workspace scenario {expected_scenario_id!r}"
        )

    selected_skill = payload.get("selectedSkill")
    expected_skill = EXPECTED_SKILLS[expected_scenario_id]
    if selected_skill != expected_skill:
        errors.append(
            f"selectedSkill must match the routing contract for {expected_scenario_id}"
        )

    evidence_limits = payload.get("evidenceLimits")
    if not _is_non_empty_string_list(evidence_limits):
        errors.append("evidenceLimits must be a non-empty string array")
    elif not SYNTHETIC_NO_LIVE_PROOF.search(" ".join(evidence_limits)):
        errors.append(
            "evidenceLimits must explicitly state synthetic input and no live runtime proof"
        )
    if payload.get("sourcePreserved") is not True:
        errors.append("sourcePreserved must be true")
    boundary = payload.get("boundary")
    if not _is_non_empty_string(boundary):
        errors.append("boundary must be a non-empty string")
    elif not SYNTHETIC_NO_LIVE_PROOF.search(boundary):
        errors.append(
            "boundary must explicitly state synthetic scope and no live runtime proof"
        )

    for string_path, value in _iter_output_strings(payload):
        if any(pattern.search(value) for pattern in UNSUPPORTED_CLAIM_PATTERNS):
            errors.append(f"unsupported live/real-success language at {string_path}")

    if expected_scenario_id == "power-query-diagnosis":
        if not _is_non_empty_string_list(payload.get("findings")):
            errors.append("findings must be a non-empty string array")
    elif expected_scenario_id == "dax-versus-environment":
        rationales = payload.get("rationales")
        if not isinstance(rationales, dict):
            errors.append("rationales must be an object")
        else:
            for key in ("daxCompatibility", "hostCompatibility"):
                if not _is_non_empty_string(rationales.get(key)):
                    errors.append(f"rationales.{key} must be a non-empty string")
    elif expected_scenario_id == "delivery-boundary":
        if not _is_non_empty_string_list(payload.get("plan")):
            errors.append("plan must be a non-empty string array")

    return _finish(report)


def verify_file(path: Path, *, workspace: Path) -> dict[str, Any]:
    expected_scenario_id = _infer_expected_scenario(workspace)
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        report = _new_report(expected_scenario_id=expected_scenario_id)
        report["errors"].append(f"cannot read benchmark output: {exc}")
        report["errors"].extend(_source_artifact_errors(workspace))
        return _finish(report)
    report = verify_benchmark_output(
        payload, expected_scenario_id=expected_scenario_id
    )
    report["errors"].extend(_source_artifact_errors(workspace))
    return _finish(report)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    report = verify_file(args.input, workspace=Path.cwd())
    sys.stdout.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
