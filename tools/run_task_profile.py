#!/usr/bin/env python3
"""Generate common Excel BI task command plans.

This wrapper reduces the need for fresh agents to remember dozens of package
scripts. It does not replace specialist tools; it creates a reviewable command
plan for a common workflow and can optionally execute the plan.
"""

from __future__ import annotations

import argparse
import json
import shlex
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PROFILE_DESCRIPTIONS = {
    "audit": "Static workbook QA audit from an OpenXML inventory.",
    "publish": "Pure-deliverable cleanup planning and post-clean verification scaffolding.",
    "pq-refresh": "Windows Excel COM Power Query refresh and refresh-status reporting.",
    "dax-review": "Power Pivot/DAX compatibility and dependency review from exported model data.",
    "cube-trace": "CUBE/MDX report-layer dependency tracing from workbook inventory.",
    "env-diagnostics": "Capability-aware Office, Excel COM, provider, and target-environment diagnostics.",
    "report-build": "Report-surface validation checks before publish cleanup.",
    "fixture": "Customer-data-free sanitized fixture bundle generation and validation.",
    "case-regression": "Real/sanitized case regression library validation.",
    "release-structural": "Cross-platform structural release gate without Excel runtime checks.",
    "release-full": "Full Windows release gate including Excel COM/provider checks.",
}


def quote(value: str) -> str:
    return shlex.quote(value)


def command_display(command: list[str]) -> str:
    return " ".join(quote(part) for part in command)


def rel(path: Path) -> str:
    return str(path)


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def profile_commands(args: argparse.Namespace) -> list[dict[str, Any]]:
    project_root = Path(args.project_root).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    workbook = Path(args.workbook).expanduser().resolve() if args.workbook else None
    model_json = Path(args.model_json).expanduser().resolve() if args.model_json else None
    query_dir = Path(args.query_dir).expanduser().resolve() if args.query_dir else None
    captured_probe = (
        Path(getattr(args, "probe_json", "")).expanduser().resolve()
        if getattr(args, "probe_json", "")
        else None
    )
    required_capabilities = list(dict.fromkeys(getattr(args, "require_capability", []) or []))
    profile = args.profile

    if profile != "env-diagnostics" and (captured_probe is not None or required_capabilities):
        raise SystemExit("--probe-json and --require-capability are only valid for profile env-diagnostics")

    openxml_json = out_dir / "openxml.json"
    formula_json = out_dir / "formula-quality.json"
    controls_json = out_dir / "workbook-controls.json"
    external_json = out_dir / "external-dependencies.json"
    triage_json = out_dir / "workbook-triage.json"

    commands: list[dict[str, Any]] = []

    def add(name: str, command: list[str], boundary: str) -> None:
        commands.append({"name": name, "command": command, "boundary": boundary})

    if profile in {"audit", "publish", "cube-trace", "report-build"} and workbook is None:
        raise SystemExit(f"--workbook is required for profile {profile}")
    if profile == "pq-refresh" and workbook is None:
        raise SystemExit("--workbook is required for profile pq-refresh")
    if profile == "dax-review" and model_json is None:
        raise SystemExit("--model-json is required for profile dax-review")

    if profile in {"audit", "publish", "cube-trace", "report-build"}:
        assert workbook is not None
        add(
            "Inspect workbook OpenXML structure",
            [sys.executable, rel(project_root / "tools" / "inspect_excel_bi_workbook.py"), rel(workbook), "--markdown", "--out-json", rel(openxml_json)],
            "Static OpenXML inventory; does not prove live calculation, refresh, VBA, or Data Model behavior.",
        )

    if profile == "audit":
        add(
            "Build formula quality report",
            [sys.executable, rel(project_root / "tools" / "build_formula_quality_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(formula_json), "--out-md", rel(out_dir / "formula-quality.md")],
            "Static formula risk report; clean output is not numeric correctness proof.",
        )
        add(
            "Build workbook controls report",
            [sys.executable, rel(project_root / "tools" / "build_workbook_controls_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(controls_json), "--out-md", rel(out_dir / "workbook-controls.md")],
            "Static controls/visibility report; intentional hidden/protected elements need owner confirmation.",
        )
        add(
            "Build external dependency report",
            [sys.executable, rel(project_root / "tools" / "build_external_dependency_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(external_json), "--out-md", rel(out_dir / "external-dependencies.md")],
            "Static dependency report; live source reachability and credentials require runtime checks.",
        )
        add(
            "Build workbook triage report",
            [sys.executable, rel(project_root / "tools" / "build_workbook_triage_report.py"), "--openxml-json", rel(openxml_json), "--formula-json", rel(formula_json), "--controls-json", rel(controls_json), "--external-json", rel(external_json), "--out-json", rel(triage_json), "--out-md", rel(out_dir / "workbook-triage.md")],
            "Triage aggregates available evidence and coverage gaps; it does not modify the workbook.",
        )

    elif profile == "publish":
        add(
            "Build source external dependency report",
            [sys.executable, rel(project_root / "tools" / "build_external_dependency_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(external_json), "--out-md", rel(out_dir / "external-dependencies.md")],
            "Use before cleanup planning; it does not remove links or freeze formulas.",
        )
        add(
            "Build pure deliverable cleanup plan",
            [sys.executable, rel(project_root / "tools" / "build_pure_deliverable_cleanup_plan.py"), "--external-json", rel(external_json), "--out-json", rel(out_dir / "cleanup-plan.json"), "--out-md", rel(out_dir / "cleanup-plan.md")],
            "Plan only. Perform destructive cleanup only on a copied workbook after refresh/value-freeze.",
        )
        add(
            "Verify cleaned deliverable after cleanup",
            [sys.executable, rel(project_root / "tools" / "build_pure_deliverable_verification_report.py"), "--external-json", rel(out_dir / "post-clean-external-dependencies.json"), "--out-json", rel(out_dir / "post-clean-verification.json"), "--out-md", rel(out_dir / "post-clean-verification.md")],
            "Run after inspecting the cleaned copy. The placeholder post-clean dependency report must be generated from the cleaned workbook.",
        )

    elif profile == "pq-refresh":
        assert workbook is not None
        raw_refresh = out_dir / "pq-refresh-raw.json"
        add(
            "Refresh Power Query through Excel COM",
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", rel(project_root / ".agents" / "skills" / "power-query-m-engineering" / "scripts" / "refresh_power_queries_excel_com.ps1"), "-WorkbookPath", rel(workbook), "-OutJson", rel(raw_refresh)],
            "Requires Windows desktop Excel and workbook access. This is live runtime evidence.",
        )
        add(
            "Build Power Query refresh status report",
            [sys.executable, rel(project_root / ".agents" / "skills" / "power-query-m-engineering" / "scripts" / "build_power_query_refresh_report.py"), "--refresh-json", rel(raw_refresh), "--out-json", rel(out_dir / "pq-refresh-report.json"), "--out-md", rel(out_dir / "pq-refresh-report.md")],
            "Summarizes completed/slow/failed/still-refreshing status; does not refresh by itself.",
        )

    elif profile == "dax-review":
        assert model_json is not None
        add(
            "Build Excel BI model report",
            [sys.executable, rel(project_root / "tools" / "build_excel_bi_model_report.py"), "--model-json", rel(model_json), "--out-json", rel(out_dir / "model-report.json"), "--out-md", rel(out_dir / "model-report.md")],
            "Static model metadata report; DAX evaluation still needs a live compatible host.",
        )
        add(
            "Lint DAX compatibility",
            [sys.executable, rel(project_root / ".agents" / "skills" / "power-pivot-dax-modeling" / "scripts" / "lint_dax_compat.py"), "--model-json", rel(model_json), "--out-json", rel(out_dir / "dax-compat.json"), "--out-md", rel(out_dir / "dax-compat.md")],
            "Flags Excel Power Pivot compatibility risks; it does not prove semantic correctness.",
        )
        add(
            "Analyze DAX dependencies",
            [sys.executable, rel(project_root / ".agents" / "skills" / "power-pivot-dax-modeling" / "scripts" / "analyze_dax_dependencies.py"), "--model-json", rel(model_json), "--out-json", rel(out_dir / "dax-dependencies.json"), "--out-md", rel(out_dir / "dax-dependencies.md")],
            "Dependency graph only; validate changed measures in the target host.",
        )

    elif profile == "cube-trace":
        add(
            "Build CUBE/MDX dependency report",
            [sys.executable, rel(project_root / "tools" / "build_cube_dependency_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(out_dir / "cube-dependencies.json"), "--out-md", rel(out_dir / "cube-dependencies.md")],
            "Static CUBE formula reference report; Data Model measure existence requires model evidence.",
        )

    elif profile == "env-diagnostics":
        if captured_probe is None:
            provider_probe_json = out_dir / "provider-probe.json"
            add(
                "Probe Office and BI providers",
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", rel(project_root / "tools" / "probe_excel_bi_providers.ps1"), "-RunExcelComSmoke", "-RunAdoWorkbookSmoke", "-OutJson", rel(provider_probe_json)],
                "Machine-specific provider detail; keep output outside the plugin package and separate from workbook behavior evidence.",
            )
            add(
                "Build provider environment report",
                [sys.executable, rel(project_root / "tools" / "build_provider_environment_report.py"), "--project-root", rel(project_root), "--probe-json", rel(provider_probe_json), "--excel-com", "--ado-workbook-smoke", "--out-json", rel(out_dir / "provider-environment.json"), "--out-md", rel(out_dir / "provider-environment.md")],
                "Preserves detailed provider and drift evidence; it does not prove workbook business correctness.",
            )
            capability_probe_json = out_dir / "excel-capabilities.json"
            add(
                "Probe Excel compatibility capabilities",
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", rel(project_root / "tools" / "probe_excel_capabilities.ps1"), "-OutJson", rel(capability_probe_json), "-Profile", "runtime"],
                "Runtime capability evidence comes from the current execution environment and generated fixtures, not the target workbook.",
            )
        else:
            capability_probe_json = captured_probe

        compatibility_command = [
            sys.executable,
            rel(project_root / "tools" / "build_excel_compatibility_report.py"),
            "--probe-json",
            rel(capability_probe_json),
            "--out-json",
            rel(out_dir / "excel-compatibility.json"),
            "--out-md",
            rel(out_dir / "excel-compatibility.md"),
            "--require-pass",
        ]
        for capability_id in required_capabilities:
            compatibility_command.extend(["--require-capability", capability_id])
        add(
            "Build Excel compatibility report",
            compatibility_command,
            "A captured probe describes its source execution environment; match it to the target environment and obtain workbook behavior evidence separately.",
        )

    elif profile == "report-build":
        add(
            "Build report formula quality report",
            [sys.executable, rel(project_root / "tools" / "build_formula_quality_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(formula_json), "--out-md", rel(out_dir / "formula-quality.md")],
            "Checks report-surface formula risks before publishing.",
        )
        add(
            "Build workbook controls report",
            [sys.executable, rel(project_root / "tools" / "build_workbook_controls_report.py"), "--openxml-json", rel(openxml_json), "--out-json", rel(controls_json), "--out-md", rel(out_dir / "workbook-controls.md")],
            "Checks report controls, panes, protection, and validation rules.",
        )

    elif profile == "fixture":
        add(
            "Build sanitized fixture bundle",
            [sys.executable, rel(project_root / "tools" / "build_sanitized_fixture_bundle.py"), "--out-dir", rel(out_dir / "sanitized-fixtures"), "--clean", "--out-json", rel(out_dir / "sanitized-fixtures.json"), "--out-md", rel(out_dir / "sanitized-fixtures.md")],
            "Creates generic customer-data-free fixtures; not proof of every customer workbook shape.",
        )

    elif profile == "case-regression":
        add(
            "Validate real/sanitized case regression library",
            [sys.executable, rel(project_root / "tools" / "run_case_regression.py"), "--project-root", rel(project_root), "--out-json", rel(out_dir / "case-regression.json"), "--out-md", rel(out_dir / "case-regression.md"), "--require-pass"],
            "Validates case definitions, layer coverage, package-tool references, and safety boundaries; does not prove private workbook behavior.",
        )

    elif profile == "release-structural":
        add(
            "Run structural release gate",
            [sys.executable, rel(project_root / "tools" / "run_release_gate.py"), "--project-root", rel(project_root), "--profile", "structural", "--out-json", rel(out_dir / "release-gate-structural.json"), "--out-md", rel(out_dir / "release-gate-structural.md")],
            "Cross-platform package/OpenXML validation; skips Excel runtime and installed-plugin checks.",
        )

    elif profile == "release-full":
        add(
            "Run full release gate",
            [sys.executable, rel(project_root / "tools" / "run_release_gate.py"), "--project-root", rel(project_root), "--out-json", rel(out_dir / "release-gate-full.json"), "--out-md", rel(out_dir / "release-gate-full.md")],
            "Windows-oriented full gate including Excel COM/provider checks when available.",
        )

    else:
        raise SystemExit(f"Unknown profile: {profile}")

    if query_dir is not None and profile == "audit":
        add(
            "Build Power Query lineage report from exported M folder",
            [sys.executable, rel(project_root / "tools" / "build_power_query_lineage_report.py"), rel(query_dir), "--out-json", rel(out_dir / "power-query-lineage.json"), "--out-md", rel(out_dir / "power-query-lineage.md")],
            "Static exported-M lineage/source-risk report; live refresh needs Excel runtime evidence.",
        )

    return commands


def build_report(args: argparse.Namespace, commands: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "profile": args.profile,
        "description": PROFILE_DESCRIPTIONS[args.profile],
        "projectRoot": str(Path(args.project_root).expanduser().resolve()),
        "outDir": str(Path(args.out_dir).expanduser().resolve()),
        "workbook": str(Path(args.workbook).expanduser().resolve()) if args.workbook else "",
        "modelJson": str(Path(args.model_json).expanduser().resolve()) if args.model_json else "",
        "queryDir": str(Path(args.query_dir).expanduser().resolve()) if args.query_dir else "",
        "probeJson": str(Path(getattr(args, "probe_json", "")).expanduser().resolve()) if getattr(args, "probe_json", "") else "",
        "requiredCapabilities": list(dict.fromkeys(getattr(args, "require_capability", []) or [])),
        "executionMode": "execute" if args.execute else "plan-only",
        "commands": commands,
        "boundaries": [
            "Profile plans are convenience wrappers, not substitutes for specialist judgment.",
            "Run destructive workbook actions only on copied workbooks.",
            "Static OpenXML reports do not prove Excel runtime calculation, refresh, Data Model, or VBA behavior.",
            "Keep execution-environment capability evidence separate from target-environment workbook behavior evidence.",
        ],
    }


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Excel BI Task Profile Plan",
        "",
        f"- profile: `{report['profile']}`",
        f"- description: {report['description']}",
        f"- execution mode: `{report['executionMode']}`",
        f"- output directory: `{report['outDir']}`",
        "",
        "## Commands",
        "",
    ]
    for index, item in enumerate(report["commands"], start=1):
        lines.extend(
            [
                f"### {index}. {item['name']}",
                "",
                "```bash",
                command_display(item["command"]),
                "```",
                "",
                f"Boundary: {item['boundary']}",
                "",
            ]
        )
    lines.extend(["## Boundaries", ""])
    for boundary in report["boundaries"]:
        lines.append(f"- {boundary}")
    lines.append("")
    return "\n".join(lines)


def run_commands(commands: list[dict[str, Any]], cwd: Path) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for item in commands:
        completed = subprocess.run(item["command"], cwd=str(cwd), text=True, capture_output=True)
        results.append(
            {
                "name": item["name"],
                "returnCode": completed.returncode,
                "stdout": completed.stdout.strip(),
                "stderr": completed.stderr.strip(),
            }
        )
        if completed.returncode != 0:
            break
    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile", choices=sorted(PROFILE_DESCRIPTIONS), required=True)
    parser.add_argument("--project-root", default=".")
    parser.add_argument("--workbook", default="", help="Workbook path for workbook-specific profiles")
    parser.add_argument("--model-json", default="", help="Exported model JSON for DAX review profiles")
    parser.add_argument("--query-dir", default="", help="Exported Power Query .m folder for optional lineage checks")
    parser.add_argument("--probe-json", default="", help="Captured excel-capability-probe JSON for env-diagnostics; skips live probing")
    parser.add_argument("--require-capability", action="append", default=[], help="Capability ID that env-diagnostics must require; repeat as needed")
    parser.add_argument("--out-dir", default="tmp/excel-bi-task-profile")
    parser.add_argument("--out-json", default="")
    parser.add_argument("--out-md", default="")
    parser.add_argument("--execute", action="store_true", help="Execute the generated commands sequentially")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    ensure_dir(out_dir)

    commands = profile_commands(args)
    report = build_report(args, commands)
    if args.execute:
        report["results"] = run_commands(commands, project_root)
        report["status"] = "pass" if all(item["returnCode"] == 0 for item in report["results"]) else "fail"
    else:
        report["status"] = "planned"

    markdown = render_markdown(report)
    if args.out_json:
        out_json = Path(args.out_json).expanduser().resolve()
        ensure_dir(out_json.parent)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = Path(args.out_md).expanduser().resolve()
        ensure_dir(out_md.parent)
        out_md.write_text(markdown, encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(markdown)

    return 1 if report.get("status") == "fail" else 0


if __name__ == "__main__":
    raise SystemExit(main())
