#!/usr/bin/env python3
"""Build a runbook for collecting cross-agent forward-test responses."""

from __future__ import annotations

import argparse
import csv
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path


PASS = "pass"
FAIL = "fail"
EXPECTED_AGENT_IDS = {"codex", "claude", "opencode", "generic"}
EXPECTED_SKILL_IDS = {
    "excel-ado-sql-data-access",
    "excel-bi-router",
    "excel-deliverable-publisher",
    "excel-report-builder",
    "excel-testing-fixtures",
    "excel-vba-workbook-engineering",
    "excel-workbook-qa-auditor",
    "mdx-cubevalue-extraction",
    "office-environment-diagnostics",
    "power-bi-semantic-model",
    "power-pivot-dax-modeling",
    "power-query-m-engineering",
}


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def safe_clean_dir(path: Path) -> None:
    resolved = path.expanduser().resolve()
    home = Path.home().resolve()
    if str(resolved) == resolved.anchor:
        raise ValueError(f"refusing to remove filesystem root: {resolved}")
    if resolved == home:
        raise ValueError(f"refusing to remove user home directory: {resolved}")
    if len(resolved.parts) < 3:
        raise ValueError(f"refusing to remove shallow directory: {resolved}")
    if resolved.exists():
        shutil.rmtree(resolved)


def read_json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def prompt_absolute_path(manifest_json: Path, entry: dict[str, object]) -> Path:
    raw = str(entry.get("path", "")).strip()
    return (manifest_json.parent / raw).resolve()


def response_relative_path(entry: dict[str, object]) -> Path:
    agent = str(entry.get("agent", "")).strip()
    skill = str(entry.get("skill", "")).strip()
    return Path(agent) / f"{skill}.md"


def response_absolute_path(responses_dir: Path, entry: dict[str, object]) -> Path:
    return (responses_dir / response_relative_path(entry)).resolve()


def validate_manifest(manifest: dict[str, object], manifest_json: Path) -> list[str]:
    failures: list[str] = []
    prompts = [item for item in manifest.get("prompts", []) if isinstance(item, dict)]
    agent_ids = {str(item.get("id")) for item in manifest.get("agentTargets", []) if isinstance(item, dict)}
    skill_ids = {str(item.get("name")) for item in manifest.get("skills", []) if isinstance(item, dict)}
    if manifest.get("status") != PASS:
        failures.append(f"manifest status={manifest.get('status')}")
    if agent_ids != EXPECTED_AGENT_IDS:
        failures.append(f"agent ids={sorted(agent_ids)}")
    if skill_ids != EXPECTED_SKILL_IDS:
        failures.append(f"skill ids={sorted(skill_ids)}")
    if manifest.get("promptCount") != len(EXPECTED_AGENT_IDS) * len(EXPECTED_SKILL_IDS):
        failures.append(f"promptCount={manifest.get('promptCount')}")
    if len(prompts) != len(EXPECTED_AGENT_IDS) * len(EXPECTED_SKILL_IDS):
        failures.append(f"prompt entries={len(prompts)}")
    for entry in prompts:
        path = prompt_absolute_path(manifest_json, entry)
        if not path.is_file():
            failures.append(f"missing prompt file={path}")
    return failures


def build_assignments(manifest: dict[str, object], manifest_json: Path, responses_dir: Path) -> list[dict[str, object]]:
    assignments: list[dict[str, object]] = []
    for entry in manifest.get("prompts", []):
        if not isinstance(entry, dict):
            continue
        prompt_path = prompt_absolute_path(manifest_json, entry)
        response_path = response_absolute_path(responses_dir, entry)
        assignments.append(
            {
                "agent": str(entry.get("agent", "")),
                "skill": str(entry.get("skill", "")),
                "title": str(entry.get("title", "")),
                "promptPath": str(prompt_path),
                "responsePath": str(response_path),
                "responseExists": response_path.is_file(),
            }
        )
    return assignments


def write_assignment_csv(path: Path, assignments: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["agent", "skill", "title", "promptPath", "responsePath", "responseExists"],
        )
        writer.writeheader()
        for row in assignments:
            writer.writerow(row)


def write_response_stubs(responses_dir: Path, assignments: list[dict[str, object]], overwrite: bool = False) -> list[str]:
    written: list[str] = []
    for item in assignments:
        path = Path(str(item["responsePath"]))
        if path.exists() and not overwrite:
            continue
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            "\n".join(
                [
                    f"# Fresh-agent response: {item['agent']} / {item['skill']}",
                    "",
                    "Paste the fresh-session response below this line before scoring.",
                    "Keep generated reports outside the plugin package.",
                    "State runtime boundaries instead of claiming unavailable Excel COM, Power Query refresh, VBA, or ADOMD evidence.",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        written.append(str(path))
    return written


def markdown_runbook(report: dict[str, object]) -> str:
    lines = [
        "# Cross-Agent Forward-Test Runbook",
        "",
        f"- Status: `{report.get('status')}`",
        f"- Assignments: `{report.get('assignmentCount')}`",
        f"- Existing responses: `{report.get('existingResponseCount')}`",
        f"- Missing responses: `{report.get('missingResponseCount')}`",
        "",
        "## Response Folder Contract",
        "",
        "Save each fresh-session response under:",
        "",
        "```text",
        "responses/<agent>/<skill>.md",
        "```",
        "",
        "Do not place customer workbooks or generated machine reports inside the plugin package.",
        "",
        "## Score Command",
        "",
        "```bash",
        str(report.get("scoreCommand", "")),
        "```",
        "",
        "## Assignment Matrix",
        "",
        "| Agent | Skill | Response Exists | Prompt | Response |",
        "|---|---|---:|---|---|",
    ]
    for item in report.get("assignments", []):
        if not isinstance(item, dict):
            continue
        lines.append(
            f"| {item.get('agent', '')} | {item.get('skill', '')} | "
            f"{item.get('responseExists', False)} | `{item.get('promptPath', '')}` | `{item.get('responsePath', '')}` |"
        )
    lines.extend(
        [
            "",
            "## Boundary",
            "",
            "This runbook proves that prompt distribution and response collection paths are defined. It does not prove external-agent behavior until real fresh-session responses are saved and scored.",
            "",
        ]
    )
    return "\n".join(lines)


def build_runbook(
    manifest_json: Path,
    responses_dir: Path,
    out_dir: Path,
    clean: bool = False,
    write_stubs: bool = False,
    overwrite_stubs: bool = False,
) -> dict[str, object]:
    manifest_json = manifest_json.expanduser().resolve()
    responses_dir = responses_dir.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    if clean:
        safe_clean_dir(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = read_json(manifest_json)
    failures = validate_manifest(manifest, manifest_json)
    assignments = build_assignments(manifest, manifest_json, responses_dir)
    stub_paths = write_response_stubs(responses_dir, assignments, overwrite=overwrite_stubs) if write_stubs else []
    existing_count = sum(1 for item in assignments if item["responseExists"])
    missing_count = len(assignments) - existing_count
    assignment_csv = out_dir / "assignment-matrix.csv"
    runbook_md = out_dir / "RUNBOOK.md"
    score_json = out_dir / "forward-test-score.json"
    score_md = out_dir / "forward-test-score.md"
    score_command = (
        "python tools/score_cross_agent_forward_test_results.py "
        f"--manifest-json \"{manifest_json}\" "
        f"--responses-dir \"{responses_dir}\" "
        f"--out-json \"{score_json}\" "
        f"--out-md \"{score_md}\" "
        "--require-pass"
    )
    report: dict[str, object] = {
        "generatedAt": now_iso(),
        "status": PASS if not failures else FAIL,
        "manifestJson": str(manifest_json),
        "responsesDir": str(responses_dir),
        "outDir": str(out_dir),
        "assignmentCount": len(assignments),
        "existingResponseCount": existing_count,
        "missingResponseCount": missing_count,
        "assignmentCsv": str(assignment_csv),
        "runbookMarkdown": str(runbook_md),
        "scoreCommand": score_command,
        "stubCount": len(stub_paths),
        "stubPaths": stub_paths,
        "assignments": assignments,
        "failures": failures,
        "boundaries": [
            "Runbook generation is not proof of external-agent execution.",
            "Real proof requires fresh-session responses saved into the response folder and a passing scorer report.",
        ],
    }
    write_assignment_csv(assignment_csv, assignments)
    runbook_md.write_text(markdown_runbook(report), encoding="utf-8")
    return report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest-json", required=True, type=Path, help="Manifest JSON from build_cross_agent_forward_test_pack.py")
    parser.add_argument("--responses-dir", required=True, type=Path, help="Response collection directory")
    parser.add_argument("--out-dir", required=True, type=Path, help="Runbook output directory")
    parser.add_argument("--clean", action="store_true", help="Remove runbook output folder before generating")
    parser.add_argument("--write-response-stubs", action="store_true", help="Create response markdown stubs under responses-dir")
    parser.add_argument("--overwrite-response-stubs", action="store_true", help="Overwrite existing response markdown stubs")
    parser.add_argument("--out-json", type=Path, help="Optional runbook JSON path")
    parser.add_argument("--out-md", type=Path, help="Optional extra runbook Markdown path")
    parser.add_argument("--require-pass", action="store_true", help="Return non-zero when runbook validation fails")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = build_runbook(
        args.manifest_json,
        args.responses_dir,
        args.out_dir,
        clean=args.clean,
        write_stubs=args.write_response_stubs,
        overwrite_stubs=args.overwrite_response_stubs,
    )
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), report)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_runbook(report), encoding="utf-8")
    print(
        json.dumps(
            {
                "status": report["status"],
                "assignmentCount": report["assignmentCount"],
                "existingResponseCount": report["existingResponseCount"],
                "missingResponseCount": report["missingResponseCount"],
            },
            ensure_ascii=False,
        )
    )
    if args.require_pass and report["status"] != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
