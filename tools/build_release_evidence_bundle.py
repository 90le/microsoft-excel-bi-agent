#!/usr/bin/env python3
"""Build a consolidated release evidence bundle for this plugin.

The bundle is a handoff artifact for agents and reviewers. It collects the
current manifest, project documentation consistency, task recipes, official
documentation index validation, goal coverage, and optionally a previously
generated release-gate report into one JSON/Markdown pair.

It intentionally does not run the full release gate by default. Pass
``--release-gate-json`` when a full or structural gate has already been run and
should be attached as evidence.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def run_json_command(command: list[str], cwd: Path, out_json: Path) -> dict[str, Any]:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    completed = subprocess.run(command, cwd=str(cwd), text=True, capture_output=True, env=env, timeout=180)
    report: dict[str, Any]
    if out_json.is_file():
        try:
            report = read_json(out_json)
        except (OSError, json.JSONDecodeError) as exc:
            report = {"status": FAIL, "errors": [f"could not read {out_json}: {exc}"]}
    else:
        report = {"status": FAIL, "errors": [f"expected report missing: {out_json}"]}
    report["_command"] = command
    report["_exitCode"] = completed.returncode
    report["_stdout"] = completed.stdout.strip()
    report["_stderr"] = completed.stderr.strip()
    if completed.returncode != 0 and str(report.get("status", "")).lower() == PASS:
        report["status"] = FAIL
        report.setdefault("errors", []).append(f"command returned {completed.returncode}")
    return report


def plugin_manifest(project_root: Path) -> dict[str, Any]:
    manifest_path = project_root / ".codex-plugin" / "plugin.json"
    data = read_json(manifest_path)
    return {
        "name": data.get("name", ""),
        "version": data.get("version", ""),
        "description": data.get("description", ""),
        "manifestPath": str(manifest_path),
        "skillRoot": data.get("skills", ""),
    }


def expected_cache_path(plugin: dict[str, Any]) -> str:
    name = str(plugin.get("name", ""))
    version = str(plugin.get("version", ""))
    if not name or not version:
        return ""
    return str(Path.home() / ".codex" / "plugins" / "cache" / "personal" / name / version)


def release_gate_summary(path: Path | None) -> dict[str, Any]:
    if not path:
        return {
            "included": False,
            "status": "not-supplied",
            "note": "Pass --release-gate-json to attach a full or structural release gate report.",
        }
    report = read_json(path.expanduser().resolve())
    checks = report.get("checks", [])
    status_counts: dict[str, int] = {}
    if isinstance(checks, list):
        for check in checks:
            if isinstance(check, dict):
                status = str(check.get("status", "")).lower() or "unknown"
                status_counts[status] = status_counts.get(status, 0) + 1
    return {
        "included": True,
        "path": str(path.expanduser().resolve()),
        "profile": report.get("profile", ""),
        "overall": report.get("overallStatus", report.get("overall", report.get("status", ""))),
        "checkCount": len(checks) if isinstance(checks, list) else 0,
        "statusCounts": status_counts,
    }


def normalize_status(value: Any) -> str:
    text = str(value).strip().lower()
    return text or "unknown"


def report_statuses(reports: dict[str, dict[str, Any]], release_gate: dict[str, Any]) -> dict[str, str]:
    statuses = {
        "projectDocs": normalize_status(reports["projectDocs"].get("status")),
        "taskRecipes": normalize_status(reports["taskRecipes"].get("status")),
        "officialDocs": PASS if not reports["officialDocs"].get("errors") else FAIL,
        "goalCoverage": normalize_status(reports["goalCoverage"].get("status")),
    }
    if release_gate.get("included"):
        statuses["releaseGate"] = normalize_status(release_gate.get("overall"))
    else:
        statuses["releaseGate"] = "not-supplied"
    return statuses


def overall_status(statuses: dict[str, str]) -> str:
    required = {key: value for key, value in statuses.items() if key != "releaseGate" or value != "not-supplied"}
    return PASS if required and all(value == PASS for value in required.values()) else FAIL


def collect_reports(project_root: Path, tmp_dir: Path) -> dict[str, dict[str, Any]]:
    reports: dict[str, dict[str, Any]] = {}
    reports["projectDocs"] = run_json_command(
        [
            sys.executable,
            str(project_root / "tools" / "validate_project_docs.py"),
            "--project-root",
            str(project_root),
            "--out-json",
            str(tmp_dir / "project-docs.json"),
        ],
        project_root,
        tmp_dir / "project-docs.json",
    )
    reports["taskRecipes"] = run_json_command(
        [
            sys.executable,
            str(project_root / "tools" / "validate_task_recipes.py"),
            "--project-root",
            str(project_root),
            "--out-json",
            str(tmp_dir / "task-recipes.json"),
        ],
        project_root,
        tmp_dir / "task-recipes.json",
    )
    reports["officialDocs"] = run_json_command(
        [
            sys.executable,
            str(project_root / "tools" / "validate_official_docs_index.py"),
            "--project-root",
            str(project_root),
            "--out-json",
            str(tmp_dir / "official-docs.json"),
        ],
        project_root,
        tmp_dir / "official-docs.json",
    )
    reports["goalCoverage"] = run_json_command(
        [
            sys.executable,
            str(project_root / "tools" / "build_goal_coverage_report.py"),
            "--project-root",
            str(project_root),
            "--out-json",
            str(tmp_dir / "goal-coverage.json"),
        ],
        project_root,
        tmp_dir / "goal-coverage.json",
    )
    return reports


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def render_markdown(bundle: dict[str, Any]) -> str:
    plugin = bundle.get("plugin", {})
    statuses = bundle.get("statuses", {})
    release_gate = bundle.get("releaseGate", {})
    lines = [
        "# Release Evidence Bundle",
        "",
        f"- plugin: `{plugin.get('name', '')}`",
        f"- version: `{plugin.get('version', '')}`",
        f"- generated: `{bundle.get('generatedAt', '')}`",
        f"- overall: **{bundle.get('status', '')}**",
        f"- expected cache: `{bundle.get('expectedCachePath', '')}`",
        "",
        "## Evidence Status",
        "",
        "| Evidence | Status | Detail |",
        "|---|---:|---|",
    ]
    detail_by_key = {
        "projectDocs": "manifest/document version alignment and required tracking docs",
        "taskRecipes": "required recipe sections, canonical skills, and package paths",
        "officialDocs": "official Microsoft documentation index structure and search routing",
        "goalCoverage": "core goal areas mapped to files, validation evidence, and completion evidence",
        "releaseGate": "attached full/structural release gate report, if supplied",
    }
    for key in ["projectDocs", "taskRecipes", "officialDocs", "goalCoverage", "releaseGate"]:
        lines.append(
            "| "
            + " | ".join([clean_markdown(key), clean_markdown(statuses.get(key, "")), clean_markdown(detail_by_key[key])])
            + " |"
        )
    lines.extend(["", "## Release Gate Attachment", ""])
    if release_gate.get("included"):
        lines.extend(
            [
                f"- path: `{release_gate.get('path', '')}`",
                f"- profile: `{release_gate.get('profile', '')}`",
                f"- overall: `{release_gate.get('overall', '')}`",
                f"- checks: {release_gate.get('checkCount', 0)}",
                f"- status counts: `{release_gate.get('statusCounts', {})}`",
            ]
        )
    else:
        lines.append(f"- {release_gate.get('note', '')}")
    lines.extend(
        [
            "",
            "## Boundaries",
            "",
            "- This bundle summarizes evidence; it does not replace the full release gate.",
            "- Generated reports should be written to temporary or deliverable folders, not committed with customer paths.",
            "- Windows Excel COM runtime behavior remains proven only by the full release gate on a machine with Excel.",
            "",
        ]
    )
    return "\n".join(lines)


def build_bundle(project_root: Path, release_gate_json: Path | None = None) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="excel_bi_release_evidence_") as tmp:
        tmp_dir = Path(tmp)
        reports = collect_reports(project_root, tmp_dir)
    plugin = plugin_manifest(project_root)
    release_gate = release_gate_summary(release_gate_json)
    statuses = report_statuses(reports, release_gate)
    bundle = {
        "generatedAt": now_iso(),
        "projectRoot": str(project_root),
        "plugin": plugin,
        "expectedCachePath": expected_cache_path(plugin),
        "reports": reports,
        "releaseGate": release_gate,
        "statuses": statuses,
        "status": overall_status(statuses),
        "boundaries": [
            "This evidence bundle summarizes current validation evidence and handoff state.",
            "It does not run full Excel COM runtime checks unless a release gate report is supplied.",
            "Use tools/run_release_gate.py for authoritative package validation.",
        ],
    }
    return bundle


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--release-gate-json", type=Path, help="Optional existing release gate JSON report to attach")
    parser.add_argument("--out-json", type=Path, help="Output bundle JSON")
    parser.add_argument("--out-md", type=Path, help="Output bundle Markdown")
    parser.add_argument("--require-pass", action="store_true", help="Exit with code 1 unless included evidence passes")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    bundle = build_bundle(project_root, release_gate_json=args.release_gate_json)
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), bundle)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(bundle), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(bundle, ensure_ascii=False, indent=2))
    if args.require_pass and bundle.get("status") != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
