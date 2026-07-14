#!/usr/bin/env python3
"""Semantically verify one plugin-eval benchmark scenario output."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any


REPORT_KIND = "excel-plugin-benchmark-output-verification"
REPORT_SCHEMA_VERSION = "1.0"
SOURCE_ARTIFACT = Path("benchmarks/fixtures/excel-bi-benchmark-source.json")
DAX_BASELINE_MANIFEST = Path(
    "benchmarks/fixtures/dax-workspace-baseline.json"
)
DEFAULT_SOURCE_SHA256 = (
    "1610e349e03bcb5e7bfab93c84f25c5add842fd35b01884d319c365a3138373b"
)
BASELINE_EXCLUDED_PATHS = frozenset(
    {
        DAX_BASELINE_MANIFEST.as_posix(),
        "benchmarks/plugin-eval-v0.2.1.json",
    }
)
BASELINE_EXCLUDED_PREFIXES = (
    ".git/",
    ".plugin-eval/",
    ".superpowers/",
)
BASELINE_EXCLUDED_PARTS = frozenset(
    {"__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache"}
)
BASELINE_ALLOWED_EXTRA_PATHS = frozenset({"benchmark-output.json"})

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
    r"\b(?:no|not|never|wasn['’]t|weren['’]t|isn['’]t|aren['’]t|didn['’]t|"
    r"doesn['’]t|don['’]t|hasn['’]t|haven['’]t|hadn['’]t|can['’]t|couldn['’]t)"
    r"\b(?:\W+\w+){0,3}\W*$",
    re.IGNORECASE,
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


def _is_baseline_excluded(relative_path: str) -> bool:
    path = Path(relative_path)
    return (
        relative_path in BASELINE_EXCLUDED_PATHS
        or relative_path in BASELINE_ALLOWED_EXTRA_PATHS
        or relative_path.startswith(BASELINE_EXCLUDED_PREFIXES)
        or any(part in BASELINE_EXCLUDED_PARTS for part in path.parts)
        or path.suffix in {".pyc", ".pyo"}
    )


def _workspace_file_hashes(workspace: Path) -> dict[str, str]:
    hashes: dict[str, str] = {}
    for path in workspace.rglob("*"):
        if not path.is_file():
            continue
        relative_path = path.relative_to(workspace).as_posix()
        if _is_baseline_excluded(relative_path):
            continue
        hashes[relative_path] = _sha256(path.read_bytes())
    return dict(sorted(hashes.items()))


def build_workspace_baseline_manifest(workspace: Path) -> dict[str, Any]:
    return {
        "kind": "excel-bi-dax-workspace-baseline",
        "schemaVersion": "1.0",
        "algorithm": "sha256",
        "scenarioId": "dax-versus-environment",
        "allowedExtraPaths": sorted(BASELINE_ALLOWED_EXTRA_PATHS),
        "excludedPaths": sorted(BASELINE_EXCLUDED_PATHS),
        "excludedPrefixes": list(BASELINE_EXCLUDED_PREFIXES),
        "files": [
            {"path": path, "sha256": digest}
            for path, digest in _workspace_file_hashes(workspace).items()
        ],
    }


def write_workspace_baseline_manifest(
    workspace: Path, output_path: Path
) -> str:
    manifest = build_workspace_baseline_manifest(workspace)
    payload = json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload_bytes = payload.encode("utf-8")
    output_path.write_bytes(payload_bytes)
    return _sha256(payload_bytes)


def _load_dax_baseline_manifest(
    workspace: Path, expected_sha256: str
) -> tuple[dict[str, str] | None, list[str]]:
    content, errors = _fixed_file_errors(
        workspace / DAX_BASELINE_MANIFEST,
        expected_sha256,
        label="DAX workspace baseline manifest",
    )
    if errors or content is None:
        return None, errors
    try:
        manifest = json.loads(content.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        return None, [f"DAX workspace baseline manifest is invalid: {exc}"]
    if not isinstance(manifest, dict):
        return None, ["DAX workspace baseline manifest must be an object"]
    if (
        manifest.get("kind") != "excel-bi-dax-workspace-baseline"
        or manifest.get("schemaVersion") != "1.0"
        or manifest.get("algorithm") != "sha256"
        or manifest.get("scenarioId") != "dax-versus-environment"
    ):
        return None, ["DAX workspace baseline manifest metadata is invalid"]
    files = manifest.get("files")
    if not isinstance(files, list):
        return None, ["DAX workspace baseline manifest files must be an array"]
    expected: dict[str, str] = {}
    for index, item in enumerate(files):
        if not isinstance(item, dict):
            return None, [f"DAX workspace baseline files[{index}] is invalid"]
        path = item.get("path")
        digest = item.get("sha256")
        if (
            not isinstance(path, str)
            or not path
            or path in expected
            or not isinstance(digest, str)
            or not SHA256_PATTERN.fullmatch(digest)
        ):
            return None, [f"DAX workspace baseline files[{index}] is invalid"]
        expected[path] = digest.lower()
    return expected, []


def _dax_workspace_change_errors(
    workspace: Path, expected_manifest_sha256: str
) -> list[str]:
    expected, errors = _load_dax_baseline_manifest(
        workspace, expected_manifest_sha256
    )
    if errors or expected is None:
        return errors
    actual = _workspace_file_hashes(workspace)
    errors = []
    for path in sorted(expected.keys() - actual.keys()):
        errors.append(f"DAX workspace change removed baseline file: {path}")
    for path in sorted(actual.keys() - expected.keys()):
        errors.append(f"DAX workspace change added unexpected file: {path}")
    for path in sorted(expected.keys() & actual.keys()):
        if expected[path] != actual[path]:
            errors.append(f"DAX workspace change modified baseline file: {path}")
    return errors


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
    dax_baseline_manifest_sha256: str,
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
            _dax_workspace_change_errors(
                workspace, dax_baseline_manifest_sha256
            )
        )
    return _finish(report)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--input", type=Path)
    mode.add_argument("--write-baseline-manifest", type=Path)
    parser.add_argument(
        "--source-sha256", default=DEFAULT_SOURCE_SHA256
    )
    parser.add_argument(
        "--dax-baseline-manifest-sha256", default=""
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    if args.write_baseline_manifest:
        manifest_sha256 = write_workspace_baseline_manifest(
            Path.cwd(), args.write_baseline_manifest
        )
        payload = {
            "kind": "excel-bi-dax-workspace-baseline-generation",
            "schemaVersion": "1.0",
            "status": "pass",
            "fileCount": len(
                build_workspace_baseline_manifest(Path.cwd())["files"]
            ),
            "manifestSha256": manifest_sha256,
        }
        sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
        return 0
    report = verify_file(
        args.input,
        workspace=Path.cwd(),
        source_sha256=args.source_sha256,
        dax_baseline_manifest_sha256=args.dax_baseline_manifest_sha256,
    )
    sys.stdout.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
