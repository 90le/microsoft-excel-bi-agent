#!/usr/bin/env python3
"""Build a public maintenance readiness audit.

This audit checks whether the public maintenance goal coverage is readable and
passing. It deliberately does not require private completion ledgers or
machine-specific Excel runtime evidence.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"
READY = "ready"
IN_PROGRESS = "in-progress"


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def read_text(path: Path) -> str:
    if not path.is_file():
        return ""
    return path.read_text(encoding="utf-8-sig", errors="replace")


def read_json(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def plugin_manifest(project_root: Path) -> dict[str, Any]:
    data = read_json(project_root / ".codex-plugin" / "plugin.json")
    return {
        "name": data.get("name", ""),
        "version": data.get("version", ""),
        "description": data.get("description", ""),
    }


def goal_coverage(project_root: Path) -> dict[str, Any]:
    tools_dir = project_root / "tools"
    if str(tools_dir) not in sys.path:
        sys.path.insert(0, str(tools_dir))
    sys.dont_write_bytecode = True
    try:
        import build_goal_coverage_report  # type: ignore
    except Exception as exc:
        return {"status": FAIL, "errors": [f"could not import build_goal_coverage_report: {exc}"]}
    try:
        return build_goal_coverage_report.build_report(project_root)
    except Exception as exc:
        return {"status": FAIL, "errors": [f"could not build goal coverage report: {exc}"]}


def build_audit(project_root: Path) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    plugin = plugin_manifest(project_root)
    goals_en = read_text(project_root / "docs" / "maintenance-goals.en-US.md")
    goals_zh = read_text(project_root / "docs" / "maintenance-goals.zh-CN.md")
    distribution = read_text(project_root / "docs" / "distribution-checklist.md")
    coverage = goal_coverage(project_root)

    blockers: list[dict[str, Any]] = []
    if coverage.get("status") != PASS:
        blockers.append({"code": "goal-coverage-not-passing", "detail": coverage.get("status", "")})
    if "Optimization Backlog" in goals_en or "优化 Backlog" in goals_zh:
        blockers.append(
            {
                "code": "completion-evidence-first-pass",
                "detail": "public maintenance goals intentionally keep a live optimization backlog",
            }
        )
    if "node tools\\install.mjs --check" not in distribution and "node tools/install.mjs --check" not in distribution:
        blockers.append(
            {
                "code": "validation-version-mismatch",
                "detail": "distribution checklist is missing the public installer check",
            }
        )

    completion_ready = not blockers
    return {
        "generatedAt": now_iso(),
        "projectRoot": str(project_root),
        "plugin": plugin,
        "status": PASS if coverage.get("status") == PASS else FAIL,
        "readinessStatus": READY if completion_ready else IN_PROGRESS,
        "completionReady": completion_ready,
        "coverage": {
            "status": coverage.get("status", ""),
            "areaCount": coverage.get("areaCount", 0),
            "passedAreaCount": coverage.get("passedAreaCount", 0),
            "failedAreaCount": coverage.get("failedAreaCount", 0),
        },
        "completionEvidence": {
            "rowCount": 0,
            "incompleteRowCount": 0,
            "incompleteRows": [],
        },
        "masterGoal": {
            "rowCount": 0,
            "incompleteRowCount": 0,
            "incompleteRows": [],
        },
        "blockers": blockers,
        "latestCurrentStatus": [
            "Public maintenance goals are documented in docs/maintenance-goals.en-US.md and docs/maintenance-goals.zh-CN.md.",
            "Raw runtime evidence and maintainer-only ledgers are intentionally outside the public repository.",
        ],
        "boundaries": [
            "A passing audit status means public maintenance goal coverage is readable and complete.",
            "`completionReady=false` is acceptable while the public optimization backlog remains active.",
            "`completionReady=true` should be used only when every public backlog item is closed or explicitly accepted.",
            "Windows Excel COM runtime proof still requires task-specific evidence outside this public package.",
        ],
    }


def clean_md(value: Any) -> str:
    return str(value if value is not None else "").replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Completion Readiness Audit",
        "",
        f"- plugin: `{report.get('plugin', {}).get('name', '')}`",
        f"- version: `{report.get('plugin', {}).get('version', '')}`",
        f"- audit status: `{report.get('status', '')}`",
        f"- readiness: **{report.get('readinessStatus', '')}**",
        f"- completion ready: `{report.get('completionReady')}`",
        "",
        "## Coverage",
        "",
        f"- status: `{report.get('coverage', {}).get('status', '')}`",
        f"- areas: `{report.get('coverage', {}).get('passedAreaCount', 0)}/{report.get('coverage', {}).get('areaCount', 0)}`",
        "",
        "## Completion Blockers",
        "",
    ]
    blockers = report.get("blockers", [])
    if blockers:
        lines.extend(["| Code | Detail | Count |", "|---|---|---:|"])
        for blocker in blockers:
            lines.append(
                "| "
                + " | ".join(
                    [
                        clean_md(blocker.get("code", "")),
                        clean_md(blocker.get("detail", "")),
                        clean_md(blocker.get("count", "")),
                    ]
                )
                + " |"
            )
    else:
        lines.append("- none")

    lines.extend(["", "## Incomplete Completion Evidence Rows", "", "- none", "", "## Boundaries", ""])
    for boundary in report.get("boundaries", []):
        lines.append(f"- {boundary}")
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, default=Path("."), help="Plugin project root")
    parser.add_argument("--out-json", type=Path, help="Write JSON audit report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown audit report")
    parser.add_argument("--print", action="store_true", help="Print Markdown audit")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero if the audit itself failed")
    parser.add_argument("--require-complete", action="store_true", help="Exit non-zero unless completionReady is true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = build_audit(args.project_root)
    markdown = render_markdown(report)
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), report)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown, encoding="utf-8")
    if args.print:
        print(markdown)
    elif not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.require_complete and not report.get("completionReady"):
        return 1
    if args.require_pass and report.get("status") != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
