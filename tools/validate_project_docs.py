#!/usr/bin/env python3
"""Validate project documentation consistency for the Excel BI plugin."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


REQUIRED_DOCS = [
    "docs/master-goal.md",
    "docs/current-status.md",
    "docs/goal-tracking.md",
    "docs/iteration-protocol.md",
    "docs/progress.md",
    "docs/validation.md",
    "docs/completion-evidence.md",
    "docs/task-recipes.md",
    "docs/project.md",
    "docs/real-case-regression.md",
]

VERSION_RE = re.compile(r"0\.1\.0\+codex(?:\.local-|\.)\d{14}")
CURRENT_STATUS_RE = re.compile(r"## \d{4}-\d{2}-\d{2} Current Active Status")
THREAD_GOAL_ID = "019e96b6-393d-77b1-bc12-456d6083d4d6"
ACTIVE_OBJECTIVE_FRAGMENT = "rendered Visual QA evidence chain V1"
MOJIBAKE_MARKERS = (
    "閹",
    "閵",
    "閿",
    "閻",
    "妤",
    "瀵",
    "鐠",
    "閼",
    "瀹屾",
    "鑳藉",
)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def plugin_version(project_root: Path) -> str:
    manifest = project_root / ".codex-plugin" / "plugin.json"
    data = json.loads(read_text(manifest))
    value = str(data.get("version", "")).strip()
    if not value:
        raise ValueError(f"missing version in {manifest}")
    return value


def latest_current_status(progress_text: str) -> str:
    matches = list(CURRENT_STATUS_RE.finditer(progress_text))
    if not matches:
        return ""
    start = matches[-1].start()
    next_header = progress_text.find("\n## ", start + 1)
    if next_header == -1:
        return progress_text[start:]
    return progress_text[start:next_header]


def find_mojibake(project_root: Path) -> list[dict[str, object]]:
    findings: list[dict[str, object]] = []
    for doc_path in REQUIRED_DOCS:
        path = project_root / doc_path
        if not path.is_file():
            continue
        text = read_text(path)
        markers = sorted(marker for marker in MOJIBAKE_MARKERS if marker in text)
        if markers:
            findings.append({"path": doc_path, "markers": markers})
    return findings


def validate(project_root: Path) -> dict[str, object]:
    errors: list[str] = []
    warnings: list[str] = []

    try:
        version = plugin_version(project_root)
    except Exception as exc:
        version = ""
        errors.append(str(exc))

    missing_docs = [path for path in REQUIRED_DOCS if not (project_root / path).is_file()]
    for path in missing_docs:
        errors.append(f"missing required project document: {path}")

    validation_text = read_text(project_root / "docs" / "validation.md") if (project_root / "docs" / "validation.md").is_file() else ""
    evidence_text = read_text(project_root / "docs" / "completion-evidence.md") if (project_root / "docs" / "completion-evidence.md").is_file() else ""
    progress_text = read_text(project_root / "docs" / "progress.md") if (project_root / "docs" / "progress.md").is_file() else ""
    project_text = read_text(project_root / "docs" / "project.md") if (project_root / "docs" / "project.md").is_file() else ""
    current_status_text = read_text(project_root / "docs" / "current-status.md") if (project_root / "docs" / "current-status.md").is_file() else ""
    master_goal_text = read_text(project_root / "docs" / "master-goal.md") if (project_root / "docs" / "master-goal.md").is_file() else ""
    goal_tracking_text = read_text(project_root / "docs" / "goal-tracking.md") if (project_root / "docs" / "goal-tracking.md").is_file() else ""

    mojibake_findings = find_mojibake(project_root)
    for finding in mojibake_findings:
        errors.append(
            f"{finding['path']} contains possible Chinese mojibake markers: "
            f"{', '.join(str(marker) for marker in finding['markers'])}"
        )

    if version:
        for label, text in [
            ("docs/validation.md", validation_text),
            ("docs/completion-evidence.md", evidence_text),
            ("docs/progress.md", progress_text),
        ]:
            if version not in text:
                errors.append(f"{label} does not contain current plugin version {version}")

        cache_path = f"microsoft-excel-bi-agent-pack\\{version}"
        cache_path_alt = f"microsoft-excel-bi-agent-pack/{version}"
        if cache_path not in validation_text and cache_path_alt not in validation_text:
            errors.append(f"docs/validation.md does not mention installed cache path for {version}")

    current_status = latest_current_status(progress_text)
    if not current_status:
        errors.append("docs/progress.md has no Current Active Status section")
    elif version and version not in current_status:
        errors.append("latest Current Active Status does not contain current plugin version")

    if "docs/task-recipes.md" not in project_text:
        errors.append("docs/project.md does not link docs/task-recipes.md")
    if "docs/current-status.md" not in project_text:
        errors.append("docs/project.md does not link docs/current-status.md")
    if "docs/completion-evidence.md" not in project_text:
        errors.append("docs/project.md does not link docs/completion-evidence.md")
    if "docs/real-case-regression.md" not in project_text:
        errors.append("docs/project.md does not link docs/real-case-regression.md")
    if version and version not in current_status_text:
        errors.append(f"docs/current-status.md does not contain current plugin version {version}")
    if "docs/goal-tracking.md" not in master_goal_text:
        errors.append("docs/master-goal.md does not link docs/goal-tracking.md")
    if THREAD_GOAL_ID not in master_goal_text or THREAD_GOAL_ID not in goal_tracking_text:
        errors.append(f"runtime goal id {THREAD_GOAL_ID} is not recorded in both master-goal and goal-tracking docs")
    if ACTIVE_OBJECTIVE_FRAGMENT not in master_goal_text or ACTIVE_OBJECTIVE_FRAGMENT not in goal_tracking_text:
        errors.append("active runtime maintenance objective is not recorded in both master-goal and goal-tracking docs")
    if "Goal Control Checklist" not in goal_tracking_text:
        errors.append("docs/goal-tracking.md does not contain the Goal Control Checklist")
    for required in [
        "docs/progress.md",
        "docs/validation.md",
        "docs/completion-evidence.md",
        "docs/real-case-regression.md",
        "tools/run_case_regression.py",
        "tools/build_goal_coverage_report.py",
        "tools/run_release_gate.py",
    ]:
        if required not in goal_tracking_text:
            errors.append(f"docs/goal-tracking.md does not bind required goal-control item: {required}")

    historical_versions = sorted(set(VERSION_RE.findall(progress_text)))
    if len(historical_versions) < 5:
        warnings.append("progress history contains fewer version markers than expected for this project")

    result = {
        "status": "pass" if not errors else "fail",
        "version": version,
        "requiredDocs": REQUIRED_DOCS,
        "historicalVersionCount": len(historical_versions),
        "latestCurrentStatus": current_status.strip().splitlines()[:6],
        "mojibakeFindings": mojibake_findings,
        "errors": errors,
        "warnings": warnings,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--out-json", default="", help="Write validation report JSON")
    args = parser.parse_args()

    project_root = Path(args.project_root).expanduser().resolve()
    result = validate(project_root)

    if args.out_json:
        out_path = Path(args.out_json).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    if result["status"] == "pass":
        print(
            "Project docs validation OK: "
            f"version={result['version']}, "
            f"requiredDocs={len(result['requiredDocs'])}, "
            f"historicalVersions={result['historicalVersionCount']}"
        )
        return 0

    print("Project docs validation failed:", file=sys.stderr)
    for error in result["errors"]:
        print(f"- {error}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
