#!/usr/bin/env python3
"""Semantically verify one plugin-eval benchmark scenario output."""

from __future__ import annotations

import argparse
import hashlib
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
DAX_CHANGE_ALLOWLIST = Path(
    "benchmarks/fixtures/dax-workspace-change-allowlist.json"
)
DEFAULT_SOURCE_SHA256 = (
    "1610e349e03bcb5e7bfab93c84f25c5add842fd35b01884d319c365a3138373b"
)
DEFAULT_DAX_ALLOWLIST_SHA256 = (
    "3372305a7de1948289a972cda845fb1ee0b9d335ad1b7d5dd65025d8458dee7f"
)

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
EXECUTION_SUCCESS_CLAIM = re.compile(
    r"\b(?:excel|power\s+query|power\s+pivot|refresh|runtime|live|automation)\b"
    r".{0,200}?\b(?P<verb>executed|ran|completed|succeeded|successfully|passed)\b",
    re.IGNORECASE | re.DOTALL,
)
NEGATED_EXECUTION = re.compile(
    r"\b(?:no|not|never)\b(?:\W+\w+){0,3}\W*$", re.IGNORECASE
)
SHA256_PATTERN = re.compile(r"[0-9a-f]{64}", re.IGNORECASE)


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


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _fixed_file_errors(
    path: Path, expected_sha256: str, *, label: str
) -> tuple[bytes | None, list[str]]:
    if not SHA256_PATTERN.fullmatch(expected_sha256):
        return None, [f"fixed {label} SHA-256 is invalid"]
    try:
        current_bytes = path.read_bytes()
    except OSError as exc:
        return None, [f"{label} is unavailable: {exc}"]
    actual_sha256 = _sha256(current_bytes)
    if actual_sha256 != expected_sha256.lower():
        return None, [
            f"{label} SHA-256 differs from the fixed benchmark baseline"
        ]
    return current_bytes, []


def _source_artifact_errors(
    workspace: Path, expected_sha256: str
) -> list[str]:
    artifact_path = workspace / SOURCE_ARTIFACT
    _, errors = _fixed_file_errors(
        artifact_path,
        expected_sha256,
        label="source artifact",
    )
    return errors


def _load_dax_allowlist(
    workspace: Path, expected_sha256: str
) -> tuple[dict[str, Any] | None, list[str]]:
    content, errors = _fixed_file_errors(
        workspace / DAX_CHANGE_ALLOWLIST,
        expected_sha256,
        label="DAX workspace change allowlist",
    )
    if errors or content is None:
        return None, errors
    try:
        allowlist = json.loads(content.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        return None, [f"DAX workspace change allowlist is invalid: {exc}"]
    if not isinstance(allowlist, dict):
        return None, ["DAX workspace change allowlist must be an object"]
    if allowlist.get("scenarioId") != "dax-versus-environment":
        return None, ["DAX workspace change allowlist scenarioId is invalid"]
    for field in ("allowedChangedPaths", "allowedChangedPrefixes"):
        value = allowlist.get(field)
        if not isinstance(value, list) or not all(
            isinstance(item, str) and item for item in value
        ):
            return None, [f"DAX workspace change allowlist {field} is invalid"]
    return allowlist, []


def _status_paths(output: str) -> list[str]:
    paths: list[str] = []
    for line in output.splitlines():
        if len(line) < 4:
            continue
        value = line[3:].replace("\\", "/")
        if " -> " in value:
            paths.extend(part.strip('"') for part in value.split(" -> ", 1))
        else:
            paths.append(value.strip('"'))
    return paths


def _dax_workspace_change_errors(
    workspace: Path, expected_allowlist_sha256: str
) -> list[str]:
    allowlist, errors = _load_dax_allowlist(
        workspace, expected_allowlist_sha256
    )
    if errors or allowlist is None:
        return errors

    git_executable = _find_git_executable()
    if not git_executable:
        return ["git is unavailable for DAX workspace change verification"]
    try:
        completed = subprocess.run(
            [
                git_executable,
                "-C",
                str(workspace),
                "status",
                "--porcelain=v1",
                "--untracked-files=all",
                "--",
            ],
            text=True,
            capture_output=True,
            check=False,
        )
    except OSError as exc:
        return [f"DAX workspace change verification is unavailable: {exc}"]
    if completed.returncode != 0:
        return ["DAX workspace change verification failed closed"]

    allowed_paths = set(allowlist["allowedChangedPaths"])
    allowed_prefixes = tuple(allowlist["allowedChangedPrefixes"])
    return [
        f"DAX workspace change is not allowed: {path}"
        for path in _status_paths(completed.stdout)
        if path not in allowed_paths and not path.startswith(allowed_prefixes)
    ]


def _has_unsupported_claim(value: str) -> bool:
    if any(pattern.search(value) for pattern in UNSUPPORTED_CLAIM_PATTERNS):
        return True
    for match in EXECUTION_SUCCESS_CLAIM.finditer(value):
        before_verb = value[match.start() : match.start("verb")]
        if NEGATED_EXECUTION.search(before_verb):
            continue
        return True
    return False


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
        if _has_unsupported_claim(value):
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


def verify_file(
    path: Path,
    *,
    workspace: Path,
    source_sha256: str,
    dax_allowlist_sha256: str,
) -> dict[str, Any]:
    expected_scenario_id = _infer_expected_scenario(workspace)
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        report = _new_report(expected_scenario_id=expected_scenario_id)
        report["errors"].append(f"cannot read benchmark output: {exc}")
        report["errors"].extend(
            _source_artifact_errors(workspace, source_sha256)
        )
        return _finish(report)
    report = verify_benchmark_output(
        payload, expected_scenario_id=expected_scenario_id
    )
    report["errors"].extend(
        _source_artifact_errors(workspace, source_sha256)
    )
    if expected_scenario_id == "dax-versus-environment":
        report["errors"].extend(
            _dax_workspace_change_errors(workspace, dax_allowlist_sha256)
        )
    return _finish(report)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument(
        "--source-sha256", default=DEFAULT_SOURCE_SHA256
    )
    parser.add_argument(
        "--dax-allowlist-sha256", default=DEFAULT_DAX_ALLOWLIST_SHA256
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    report = verify_file(
        args.input,
        workspace=Path.cwd(),
        source_sha256=args.source_sha256,
        dax_allowlist_sha256=args.dax_allowlist_sha256,
    )
    sys.stdout.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
