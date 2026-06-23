#!/usr/bin/env python3
"""Build a self-contained onboarding bundle for a fresh Excel BI agent.

The bundle is an onboarding artifact, not proof of workbook behavior. It
collects capability discovery, release evidence, task recipes, validation
commands, and runtime boundaries into one directory that can be handed to a
fresh Codex, Claude, OpenCode, or generic agent.
"""

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
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


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


def rel(path: Path, base: Path) -> str:
    return path.relative_to(base).as_posix()


def markdown_table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    def clean(value: Any) -> str:
        return str(value if value is not None else "").replace("|", "\\|").replace("\n", " ")

    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(clean(value) for value in row) + " |")
    return lines


def load_plugin(project_root: Path) -> dict[str, Any]:
    manifest = read_json(project_root / ".codex-plugin" / "plugin.json")
    return {
        "name": manifest.get("name", ""),
        "version": manifest.get("version", ""),
        "description": manifest.get("description", ""),
        "skills": manifest.get("skills", ""),
    }


def write_validation_commands(out_path: Path, plugin: dict[str, Any]) -> list[str]:
    commands = [
        "python tools/run_release_gate.py --project-root . --out-json \"$env:TEMP/excel_bi_release_gate.json\" --out-md \"$env:TEMP/excel_bi_release_gate.md\"",
        "tools/run_release_gate.sh --profile structural",
        "python tools/validate_project_docs.py --project-root .",
        "python tools/validate_github_community_health.py --project-root .",
        "python tools/validate_task_recipes.py --project-root .",
        "python tools/build_goal_coverage_report.py --project-root . --require-pass",
        "python tools/build_capability_catalog.py --project-root . --require-pass",
        "python tools/build_release_evidence_bundle.py --project-root . --require-pass",
        "python <plugin-creator-skill-root>/scripts/validate_plugin.py .",
        "python tools/deploy-local-plugin.py --project-root . --replace --install --update-cachebuster",
        "codex plugin list",
    ]
    lines = [
        "# Validation Commands",
        "",
        f"- Plugin: `{plugin.get('name', '')}`",
        f"- Version: `{plugin.get('version', '')}`",
        "",
        "## Source Package",
        "",
        "```powershell",
        commands[0],
        commands[2],
        commands[3],
        commands[4],
        commands[5],
        commands[6],
        commands[7],
        commands[8],
        "```",
        "",
        "## Git Bash, Linux, macOS Structural Gate",
        "",
        "```bash",
        commands[1],
        "```",
        "",
        "## Deployment",
        "",
        "```powershell",
        commands[9],
        commands[10],
        "```",
        "",
        "## Boundary",
        "",
        "- Structural gates do not prove desktop Excel COM, VBA execution, Power Query refresh, or provider availability.",
        "- Full runtime evidence requires `tools/run_release_gate.py` on a Windows machine with desktop Excel and the required providers.",
        "- Write machine-specific reports to temp or deliverable folders, not into the plugin package.",
        "",
    ]
    out_path.write_text("\n".join(lines), encoding="utf-8")
    return commands


def write_bootstrap_markdown(
    out_path: Path,
    manifest: dict[str, Any],
    capability: dict[str, Any],
    release_evidence: dict[str, Any],
    validation_commands: list[str],
) -> None:
    plugin = manifest["plugin"]
    summary = capability.get("summary", {})
    workflows = capability.get("workflows", [])
    statuses = release_evidence.get("statuses", {})
    files = manifest.get("files", [])

    lines: list[str] = [
        "# Agent Bootstrap Bundle",
        "",
        f"- Status: `{manifest.get('status', '')}`",
        f"- Plugin: `{plugin.get('name', '')}`",
        f"- Version: `{plugin.get('version', '')}`",
        f"- Generated: `{manifest.get('generatedAt', '')}`",
        "",
        "## What This Bundle Is",
        "",
        "This bundle is onboarding infrastructure for a fresh agent. It explains the package capabilities, entry points, validation commands, and hard runtime boundaries without requiring customer files.",
        "",
        "This bundle is not proof of external-agent behavior or workbook-specific validation. A real workbook still needs task-specific inspection, Excel COM runtime checks where relevant, and recorded evidence.",
        "",
        "## Start Here",
        "",
        "1. Read `BOOTSTRAP.md` for the operating boundary.",
        "2. Read `capability-catalog.md` to choose the right skill, script, and workflow.",
        "3. Read `task-recipes.md` for sanitized workbook workflows.",
        "4. Run the relevant commands in `validation-commands.md` before claiming completion.",
        "5. Use `release-evidence.md` to understand the current package validation state.",
        "",
        "## Capability Snapshot",
        "",
        f"- Skills: `{summary.get('skillCount', 0)}`",
        f"- Scripts/tools: `{summary.get('toolCount', 0)}`",
        f"- Official docs entries: `{summary.get('officialDocEntryCount', 0)}`",
        f"- Release-gate checks: `{summary.get('releaseGateCheckCount', 0)}`",
        "",
        "## Core Workflows",
        "",
    ]
    lines.extend(
        markdown_table(
            ["Workflow", "Skills", "Boundary"],
            [[item.get("title", ""), ", ".join(item.get("skills", [])), item.get("boundary", "")] for item in workflows],
        )
    )
    lines.extend(["", "## Evidence Status", ""])
    lines.extend(
        markdown_table(
            ["Evidence", "Status"],
            [[key, value] for key, value in statuses.items()],
        )
    )
    lines.extend(["", "## Required Validation Commands", ""])
    for command in validation_commands[:7]:
        lines.append(f"- `{command}`")
    lines.extend(["", "## Included Files", ""])
    for file_info in files:
        lines.append(f"- `{file_info.get('path', '')}` - {file_info.get('description', '')}")
    lines.extend(
        [
            "",
            "## Boundaries",
            "",
            "- Do not put customer workbooks, screenshots, machine-specific reports, credentials, or external connection dumps inside the plugin package.",
            "- Use OpenXML/static checks for cross-platform structure; use Windows Excel COM checks only when live Excel behavior matters.",
            "- Generated prompts, stubs, and bootstrap materials are onboarding or collection infrastructure. They are not evidence until outputs are produced and scored.",
            "- Official documentation indexes route agents to Microsoft sources; they are not full offline copies of Microsoft documentation.",
            "",
        ]
    )
    out_path.write_text("\n".join(lines), encoding="utf-8")


def create_zip(out_dir: Path, zip_path: Path) -> None:
    zip_path = zip_path.expanduser().resolve()
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(out_dir.rglob("*")):
            if path.is_file() and path.resolve() != zip_path:
                archive.write(path, path.relative_to(out_dir).as_posix())


def build_bundle(project_root: Path, out_dir: Path, clean: bool = False, zip_bundle: bool = False) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    if clean:
        safe_clean_dir(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    plugin = load_plugin(project_root)
    capability_json = out_dir / "capability-catalog.json"
    capability_md = out_dir / "capability-catalog.md"
    release_json = out_dir / "release-evidence.json"
    release_md = out_dir / "release-evidence.md"
    task_recipes = out_dir / "task-recipes.md"
    validation_md = out_dir / "validation-commands.md"
    bootstrap_md = out_dir / "BOOTSTRAP.md"
    manifest_json = out_dir / "bootstrap-manifest.json"
    zip_path = out_dir / "agent-bootstrap-bundle.zip" if zip_bundle else None

    capability_result = run_python(
        [
            str(project_root / "tools" / "build_capability_catalog.py"),
            "--project-root",
            str(project_root),
            "--out-json",
            str(capability_json),
            "--out-md",
            str(capability_md),
            "--require-pass",
        ],
        project_root,
    )
    release_result = run_python(
        [
            str(project_root / "tools" / "build_release_evidence_bundle.py"),
            "--project-root",
            str(project_root),
            "--out-json",
            str(release_json),
            "--out-md",
            str(release_md),
            "--require-pass",
        ],
        project_root,
    )

    if (project_root / "docs" / "task-recipes.md").is_file():
        shutil.copy2(project_root / "docs" / "task-recipes.md", task_recipes)
    else:
        task_recipes.write_text("# Task Recipes\n\nNo task recipe document found.\n", encoding="utf-8")

    validation_commands = write_validation_commands(validation_md, plugin)
    capability = read_json(capability_json) if capability_json.is_file() else {"status": FAIL, "summary": {}, "workflows": []}
    release_evidence = read_json(release_json) if release_json.is_file() else {"status": FAIL, "statuses": {}}

    files = [
        {"path": "BOOTSTRAP.md", "description": "Fresh-agent orientation, start steps, and boundaries."},
        {"path": "bootstrap-manifest.json", "description": "Machine-readable bundle manifest."},
        {"path": "capability-catalog.json", "description": "Machine-readable skills, scripts, workflows, docs, and release-gate inventory."},
        {"path": "capability-catalog.md", "description": "Human-readable capability catalog."},
        {"path": "release-evidence.json", "description": "Machine-readable release evidence summary."},
        {"path": "release-evidence.md", "description": "Human-readable release evidence summary."},
        {"path": "task-recipes.md", "description": "Sanitized task recipes copied from project docs."},
        {"path": "validation-commands.md", "description": "Validation and deployment command reference."},
    ]
    if zip_path:
        files.append({"path": zip_path.name, "description": "Optional archive of the bootstrap bundle."})

    failures: list[str] = []
    if capability_result.returncode != 0:
        failures.append(f"capability catalog command returned {capability_result.returncode}")
    if release_result.returncode != 0:
        failures.append(f"release evidence command returned {release_result.returncode}")
    if capability.get("status") != PASS:
        failures.append(f"capability catalog status={capability.get('status')}")
    if release_evidence.get("status") != PASS:
        failures.append(f"release evidence status={release_evidence.get('status')}")

    manifest: dict[str, Any] = {
        "generatedAt": now_iso(),
        "status": PASS if not failures else FAIL,
        "plugin": plugin,
        "projectRoot": str(project_root),
        "outDir": str(out_dir),
        "files": files,
        "capabilityCatalog": {
            "status": capability.get("status", ""),
            "summary": capability.get("summary", {}),
            "path": rel(capability_json, out_dir),
            "markdownPath": rel(capability_md, out_dir),
        },
        "releaseEvidence": {
            "status": release_evidence.get("status", ""),
            "statuses": release_evidence.get("statuses", {}),
            "path": rel(release_json, out_dir),
            "markdownPath": rel(release_md, out_dir),
        },
        "validationCommands": validation_commands,
        "boundaries": [
            "This bundle is onboarding infrastructure, not proof of external-agent behavior.",
            "This bundle is not workbook-specific validation evidence.",
            "Excel COM, Power Query refresh, VBA execution, provider availability, and Power Pivot behavior require task-specific runtime validation.",
            "Do not store customer workbooks or machine-specific reports in the plugin package.",
        ],
        "commands": {
            "capabilityCatalog": command_text([str(project_root / "tools" / "build_capability_catalog.py"), "--project-root", str(project_root)]),
            "releaseEvidence": command_text([str(project_root / "tools" / "build_release_evidence_bundle.py"), "--project-root", str(project_root)]),
        },
        "failures": failures,
    }

    write_bootstrap_markdown(bootstrap_md, manifest, capability, release_evidence, validation_commands)
    write_json(manifest_json, manifest)

    # Re-read the Markdown after writing it so validation can inspect the exact artifact.
    bootstrap_text = read_text(bootstrap_md)
    for required in [
        "Agent Bootstrap Bundle",
        "onboarding infrastructure",
        "not proof of external-agent behavior",
        "validation-commands.md",
        "capability-catalog.md",
    ]:
        if required not in bootstrap_text:
            failures.append(f"BOOTSTRAP.md missing {required}")

    for required_path in [bootstrap_md, manifest_json, capability_json, capability_md, release_json, release_md, task_recipes, validation_md]:
        if not required_path.is_file():
            failures.append(f"missing expected file: {required_path}")

    if zip_path:
        create_zip(out_dir, zip_path)
        if not zip_path.is_file():
            failures.append(f"zip archive missing: {zip_path}")
        else:
            with zipfile.ZipFile(zip_path) as archive:
                names = set(archive.namelist())
            for required in [
                "BOOTSTRAP.md",
                "bootstrap-manifest.json",
                "capability-catalog.json",
                "capability-catalog.md",
                "release-evidence.json",
                "release-evidence.md",
                "task-recipes.md",
                "validation-commands.md",
            ]:
                if required not in names:
                    failures.append(f"zip archive missing {required}")

    if failures:
        manifest["status"] = FAIL
        manifest["failures"] = failures
        write_json(manifest_json, manifest)
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, default=Path("."), help="Plugin project root")
    parser.add_argument("--out-dir", type=Path, required=True, help="Output bootstrap bundle directory")
    parser.add_argument("--clean", action="store_true", help="Delete the output directory before generation")
    parser.add_argument("--zip", action="store_true", help="Create agent-bootstrap-bundle.zip inside the output directory")
    parser.add_argument("--print", action="store_true", help="Print the manifest JSON")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero unless the bundle validates")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = build_bundle(args.project_root, args.out_dir, clean=args.clean, zip_bundle=args.zip)
    if args.print:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    else:
        print(
            f"Agent bootstrap bundle {manifest.get('status')} "
            f"for {manifest.get('plugin', {}).get('name', '')}@{manifest.get('plugin', {}).get('version', '')}: "
            f"{manifest.get('outDir', '')}"
        )
    if args.require_pass and manifest.get("status") != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
