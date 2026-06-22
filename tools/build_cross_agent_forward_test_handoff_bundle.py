#!/usr/bin/env python3
"""Build a complete cross-agent fresh-session forward-test handoff bundle."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"


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


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def run_python(args: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, *args],
        cwd=str(cwd),
        text=True,
        encoding="utf-8",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )


def command_text(parts: list[str]) -> str:
    return " ".join(f'"{part}"' if " " in part else part for part in parts)


def markdown_handoff(report: dict[str, Any]) -> str:
    lines = [
        "# Cross-Agent Fresh-Session Handoff Bundle",
        "",
        f"- Status: `{report['status']}`",
        f"- Prompt count: `{report['promptCount']}`",
        f"- Assignment count: `{report['assignmentCount']}`",
        f"- Response stubs: `{report['stubCount']}`",
        "",
        "## Use This Bundle",
        "",
        "1. Open `pack/prompts/<agent>/<skill>.md` in a fresh session for the named agent.",
        "2. Save the response to `responses/<agent>/<skill>.md`.",
        "3. Run the score command after all responses are saved.",
        "",
        "```bash",
        str(report["scoreCommand"]),
        "```",
        "",
        "## Included Paths",
        "",
        f"- Prompt pack manifest: `{report['packManifest']}`",
        f"- Collection runbook: `{report['runbookMarkdown']}`",
        f"- Assignment matrix: `{report['assignmentCsv']}`",
        f"- Response folder: `{report['responsesDir']}`",
        f"- Score JSON target: `{report['scoreJson']}`",
        f"- Score Markdown target: `{report['scoreMarkdown']}`",
        "",
        "## Boundary",
        "",
        "This bundle is collection infrastructure. It does not prove external-agent behavior until real fresh-session responses are saved and the scorer passes.",
        "Generated prompts must use generic fixtures or temp folders only; do not put customer workbooks or machine-specific reports inside the plugin package.",
        "",
    ]
    if report.get("zipPath"):
        lines.extend(["## Archive", "", f"- Zip archive: `{report['zipPath']}`", ""])
    if report.get("failures"):
        lines.extend(["## Failures", ""])
        for failure in report["failures"]:
            lines.append(f"- {failure}")
    return "\n".join(lines)


def create_zip(out_dir: Path, zip_path: Path) -> None:
    zip_path = zip_path.expanduser().resolve()
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(out_dir.rglob("*")):
            if path.is_file() and path.resolve() != zip_path:
                archive.write(path, path.relative_to(out_dir).as_posix())


def validate_bundle(report: dict[str, Any], out_dir: Path, zip_path: Path | None) -> list[str]:
    failures: list[str] = []
    expected_files = [
        Path(report["packManifest"]),
        Path(report["runbookMarkdown"]),
        Path(report["assignmentCsv"]),
        Path(report["handoffMarkdown"]),
        Path(report["handoffJson"]),
    ]
    for path in expected_files:
        if not path.is_file():
            failures.append(f"missing expected file: {path}")

    if report.get("promptCount") != 48:
        failures.append(f"promptCount={report.get('promptCount')}")
    if report.get("assignmentCount") != 48:
        failures.append(f"assignmentCount={report.get('assignmentCount')}")
    if report.get("stubCount") != 48:
        failures.append(f"stubCount={report.get('stubCount')}")
    if report.get("missingResponseCount") != 48:
        failures.append(f"missingResponseCount={report.get('missingResponseCount')}")
    if "score_cross_agent_forward_test_results.py" not in str(report.get("scoreCommand", "")):
        failures.append("score command missing scorer script")

    response_stub = Path(report["responsesDir"]) / "codex" / "excel-bi-router.md"
    if not response_stub.is_file():
        failures.append(f"missing response stub: {response_stub}")
    else:
        text = response_stub.read_text(encoding="utf-8")
        if "Fresh-agent response" not in text or "runtime boundaries" not in text:
            failures.append("response stub missing fresh-agent boundary text")

    handoff = Path(report["handoffMarkdown"])
    if handoff.is_file():
        text = handoff.read_text(encoding="utf-8")
        for required in [
            "Cross-Agent Fresh-Session Handoff Bundle",
            "fresh-session responses",
            "responses/<agent>/<skill>.md",
            "score_cross_agent_forward_test_results.py",
        ]:
            if required not in text:
                failures.append(f"handoff Markdown missing {required}")

    if zip_path is not None:
        if not zip_path.is_file():
            failures.append(f"zip archive missing: {zip_path}")
        else:
            with zipfile.ZipFile(zip_path) as archive:
                names = set(archive.namelist())
            for required in [
                "pack/forward-test-pack.json",
                "runbook/RUNBOOK.md",
                "responses/codex/excel-bi-router.md",
                "HANDOFF.md",
                "handoff-manifest.json",
            ]:
                if required not in names:
                    failures.append(f"zip archive missing {required}")

    if not out_dir.is_dir():
        failures.append(f"outDir missing: {out_dir}")
    return failures


def build_bundle(project_root: Path, out_dir: Path, clean: bool = False, zip_path: Path | None = None) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    if clean:
        safe_clean_dir(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pack_dir = out_dir / "pack"
    runbook_dir = out_dir / "runbook"
    responses_dir = out_dir / "responses"
    score_dir = out_dir / "score"
    pack_manifest = pack_dir / "forward-test-pack.json"
    pack_readme = pack_dir / "README.md"
    runbook_json = runbook_dir / "runbook.json"
    runbook_md_copy = runbook_dir / "RUNBOOK.copy.md"
    handoff_json = out_dir / "handoff-manifest.json"
    handoff_md = out_dir / "HANDOFF.md"
    score_json = score_dir / "forward-test-score.json"
    score_md = score_dir / "forward-test-score.md"

    pack_cmd = [
        str(project_root / "tools" / "build_cross_agent_forward_test_pack.py"),
        "--project-root",
        str(project_root),
        "--out-dir",
        str(pack_dir),
        "--clean",
        "--out-json",
        str(pack_manifest),
        "--out-md",
        str(pack_readme),
        "--require-pass",
    ]
    pack_result = run_python(pack_cmd, project_root)

    runbook_cmd = [
        str(project_root / "tools" / "build_cross_agent_forward_test_runbook.py"),
        "--manifest-json",
        str(pack_manifest),
        "--responses-dir",
        str(responses_dir),
        "--out-dir",
        str(runbook_dir),
        "--clean",
        "--write-response-stubs",
        "--out-json",
        str(runbook_json),
        "--out-md",
        str(runbook_md_copy),
        "--require-pass",
    ]
    runbook_result = run_python(runbook_cmd, project_root) if pack_result.returncode == 0 else None

    failures: list[str] = []
    if pack_result.returncode != 0:
        failures.append(f"prompt pack command failed: {pack_result.stderr.strip() or pack_result.stdout.strip()}")
    if runbook_result is None:
        failures.append("runbook command skipped because prompt pack failed")
        pack_report: dict[str, Any] = {}
        runbook_report: dict[str, Any] = {}
    else:
        if runbook_result.returncode != 0:
            failures.append(f"runbook command failed: {runbook_result.stderr.strip() or runbook_result.stdout.strip()}")
        pack_report = read_json(pack_manifest) if pack_manifest.is_file() else {}
        runbook_report = read_json(runbook_json) if runbook_json.is_file() else {}

    score_command = (
        "python tools/score_cross_agent_forward_test_results.py "
        f"--manifest-json \"{pack_manifest}\" "
        f"--responses-dir \"{responses_dir}\" "
        f"--out-json \"{score_json}\" "
        f"--out-md \"{score_md}\" "
        "--require-pass"
    )
    report: dict[str, Any] = {
        "generatedAt": now_iso(),
        "status": FAIL,
        "projectRoot": str(project_root),
        "outDir": str(out_dir),
        "packManifest": str(pack_manifest),
        "packReadme": str(pack_readme),
        "runbookJson": str(runbook_json),
        "runbookMarkdown": str(runbook_dir / "RUNBOOK.md"),
        "runbookMarkdownCopy": str(runbook_md_copy),
        "assignmentCsv": str(runbook_dir / "assignment-matrix.csv"),
        "responsesDir": str(responses_dir),
        "scoreJson": str(score_json),
        "scoreMarkdown": str(score_md),
        "scoreCommand": score_command,
        "handoffJson": str(handoff_json),
        "handoffMarkdown": str(handoff_md),
        "zipPath": str(zip_path.expanduser().resolve()) if zip_path else "",
        "promptCount": pack_report.get("promptCount", 0),
        "assignmentCount": runbook_report.get("assignmentCount", 0),
        "stubCount": runbook_report.get("stubCount", 0),
        "missingResponseCount": runbook_report.get("missingResponseCount", 0),
        "commands": {
            "pack": command_text([sys.executable, *pack_cmd]),
            "runbook": command_text([sys.executable, *runbook_cmd]),
            "score": score_command,
        },
        "boundaries": [
            "This is collection infrastructure, not proof of external-agent execution.",
            "Fresh-session responses must be saved and scored before claiming external-agent behavior.",
            "Keep customer workbooks, screenshots, and machine-specific reports outside the plugin package.",
        ],
        "failures": failures,
    }

    handoff_md.write_text(markdown_handoff(report), encoding="utf-8")
    write_json(handoff_json, report)
    validation_failures = validate_bundle(report, out_dir, None)
    if validation_failures:
        failures.extend(validation_failures)
    report["failures"] = failures
    report["status"] = PASS if not failures else FAIL
    write_json(handoff_json, report)
    handoff_md.write_text(markdown_handoff(report), encoding="utf-8")

    if zip_path is not None and report["status"] == PASS:
        create_zip(out_dir, zip_path)
        zip_failures = validate_bundle(report, out_dir, zip_path)
        if zip_failures:
            report["failures"] = zip_failures
            report["status"] = FAIL
        write_json(handoff_json, report)
        handoff_md.write_text(markdown_handoff(report), encoding="utf-8")

    return report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", type=Path, help="Plugin project root")
    parser.add_argument("--out-dir", required=True, type=Path, help="Output folder for the handoff bundle")
    parser.add_argument("--clean", action="store_true", help="Remove the output folder before generating")
    parser.add_argument("--zip-path", type=Path, help="Optional zip archive path for the generated bundle")
    parser.add_argument("--out-json", type=Path, help="Optional copy of handoff manifest JSON")
    parser.add_argument("--out-md", type=Path, help="Optional copy of handoff Markdown")
    parser.add_argument("--require-pass", action="store_true", help="Return non-zero when bundle validation fails")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = build_bundle(args.project_root, args.out_dir, clean=args.clean, zip_path=args.zip_path)
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), report)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_handoff(report), encoding="utf-8")
    print(
        json.dumps(
            {
                "status": report["status"],
                "promptCount": report["promptCount"],
                "assignmentCount": report["assignmentCount"],
                "stubCount": report["stubCount"],
                "zipPath": report.get("zipPath", ""),
            },
            ensure_ascii=False,
        )
    )
    if args.require_pass and report["status"] != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
