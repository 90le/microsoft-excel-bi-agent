#!/usr/bin/env python3
"""Build a completion-readiness audit for the active Excel BI plugin goal.

This audit is stricter than coverage. A coverage report can prove that every
named area has files and validation evidence. Completion readiness additionally
checks whether the project documents themselves say the goal is complete or
still first-pass/in-progress.

Use `--require-complete` for the final closure audit before calling
`update_goal(status=complete)`.
"""

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
READY = "ready"
IN_PROGRESS = "in-progress"
ACCEPTED = "accepted-boundary"


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


def parse_markdown_table(text: str, header_prefix: str) -> list[dict[str, str]]:
    lines = text.splitlines()
    rows: list[dict[str, str]] = []
    headers: list[str] = []
    in_table = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("|") and header_prefix in stripped:
            headers = [part.strip() for part in stripped.strip("|").split("|")]
            in_table = True
            continue
        if not in_table:
            continue
        if not stripped.startswith("|"):
            break
        cells = [part.strip() for part in stripped.strip("|").split("|")]
        if all(set(cell.replace(":", "").strip()) <= {"-"} for cell in cells):
            continue
        if len(cells) != len(headers):
            continue
        rows.append(dict(zip(headers, cells)))
    return rows


def parse_completion_rows(completion_text: str) -> list[dict[str, str]]:
    return parse_markdown_table(completion_text, "Goal requirement")


def parse_master_rows(master_text: str) -> list[dict[str, str]]:
    return parse_markdown_table(master_text, "Area")


def goal_coverage(project_root: Path) -> dict[str, Any]:
    tools_dir = project_root / "tools"
    if str(tools_dir) not in sys.path:
        sys.path.insert(0, str(tools_dir))
    try:
        import build_goal_coverage_report  # type: ignore
    except Exception as exc:
        return {"status": FAIL, "errors": [f"could not import build_goal_coverage_report: {exc}"]}
    try:
        return build_goal_coverage_report.build_report(project_root)
    except Exception as exc:
        return {"status": FAIL, "errors": [f"could not build goal coverage report: {exc}"]}


def classify_evidence_status(value: str) -> str:
    text = value.strip().lower()
    if text in {"complete", "completed", "done"}:
        return "complete"
    if "accepted boundary" in text or "accepted-boundary" in text or "accepted" == text:
        return ACCEPTED
    if "first pass" in text or "in progress" in text or "partial" in text:
        return IN_PROGRESS
    if not text:
        return "missing"
    return text


def latest_current_status(progress_text: str) -> str:
    matches = list(re.finditer(r"^## \d{4}-\d{2}-\d{2} Current Active Status", progress_text, flags=re.MULTILINE))
    if not matches:
        return ""
    start = matches[-1].start()
    next_match = re.search(r"^## ", progress_text[start + 1 :], flags=re.MULTILINE)
    if not next_match:
        return progress_text[start:]
    return progress_text[start : start + 1 + next_match.start()]


def build_audit(project_root: Path) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    plugin = plugin_manifest(project_root)
    validation_text = read_text(project_root / "docs" / "validation.md")
    progress_text = read_text(project_root / "docs" / "progress.md")
    completion_text = read_text(project_root / "docs" / "completion-evidence.md")
    master_text = read_text(project_root / "docs" / "master-goal.md")
    goal_tracking_text = read_text(project_root / "docs" / "goal-tracking.md")

    completion_rows = parse_completion_rows(completion_text)
    master_rows = parse_master_rows(master_text)
    coverage = goal_coverage(project_root)
    latest_status = latest_current_status(progress_text)

    incomplete_completion_rows: list[dict[str, str]] = []
    for row in completion_rows:
        status = classify_evidence_status(row.get("Status", ""))
        if status not in {"complete", ACCEPTED}:
            incomplete_completion_rows.append(
                {
                    "requirement": row.get("Goal requirement", ""),
                    "status": row.get("Status", ""),
                    "remaining": row.get("Remaining hardening", ""),
                }
            )

    incomplete_master_rows: list[dict[str, str]] = []
    for row in master_rows:
        status = classify_evidence_status(row.get("Status", ""))
        if status not in {"complete", ACCEPTED}:
            incomplete_master_rows.append(
                {
                    "area": row.get("Area", ""),
                    "status": row.get("Status", ""),
                    "criteria": row.get("Completion Criteria", ""),
                }
            )

    blockers: list[dict[str, Any]] = []
    if coverage.get("status") != PASS:
        blockers.append({"code": "goal-coverage-not-passing", "detail": coverage.get("status", "")})
    if incomplete_completion_rows:
        blockers.append(
            {
                "code": "completion-evidence-first-pass",
                "count": len(incomplete_completion_rows),
                "detail": "completion-evidence.md still records non-complete requirement rows",
            }
        )
    if incomplete_master_rows:
        blockers.append(
            {
                "code": "master-goal-not-complete",
                "count": len(incomplete_master_rows),
                "detail": "master-goal.md still records non-complete area rows",
            }
        )
    if "Goal status remains active" not in latest_status and "Goal status is complete" not in latest_status:
        blockers.append({"code": "latest-progress-missing-goal-state", "detail": "latest Current Active Status does not state active or complete goal state"})
    if "Only after all audit points are satisfied may a future agent call `update_goal(status=complete)`." not in goal_tracking_text:
        blockers.append({"code": "completion-audit-guard-missing", "detail": "goal-tracking.md is missing explicit complete guard text"})
    if plugin.get("version") and plugin["version"] not in validation_text:
        blockers.append({"code": "validation-version-mismatch", "detail": f"docs/validation.md does not contain {plugin['version']}"})

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
            "rowCount": len(completion_rows),
            "incompleteRowCount": len(incomplete_completion_rows),
            "incompleteRows": incomplete_completion_rows,
        },
        "masterGoal": {
            "rowCount": len(master_rows),
            "incompleteRowCount": len(incomplete_master_rows),
            "incompleteRows": incomplete_master_rows,
        },
        "blockers": blockers,
        "latestCurrentStatus": latest_status.strip().splitlines()[:8],
        "boundaries": [
            "A passing audit status means the readiness audit ran and coverage evidence is readable.",
            "`completionReady=false` means the active thread goal must remain open.",
            "`completionReady=true` means every audited row is complete or has an accepted boundary.",
            "Use --require-complete only when the project is ready for a final update_goal(status=complete) decision.",
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

    lines.extend(["", "## Incomplete Completion Evidence Rows", ""])
    incomplete = report.get("completionEvidence", {}).get("incompleteRows", [])
    if incomplete:
        lines.extend(["| Requirement | Status | Remaining hardening |", "|---|---|---|"])
        for row in incomplete:
            lines.append(
                "| "
                + " | ".join([clean_md(row.get("requirement", "")), clean_md(row.get("status", "")), clean_md(row.get("remaining", ""))])
                + " |"
            )
    else:
        lines.append("- none")

    lines.extend(["", "## Boundaries", ""])
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
