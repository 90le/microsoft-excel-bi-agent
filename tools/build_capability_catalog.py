#!/usr/bin/env python3
"""Build a machine-readable capability catalog for the Excel BI agent pack.

The catalog is a static discovery artifact. It scans plugin files, skills,
scripts, official documentation indexes, and known workflow entry points without
opening Excel, refreshing Power Query, executing VBA, or touching customer
workbooks.
"""

from __future__ import annotations

import argparse
import ast
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REQUIRED_SKILLS = {
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

REQUIRED_TOOLS = {
    "inspect_excel_bi_workbook.py",
    "build_workbook_triage_report.py",
    "build_visual_qa_report.py",
    "export_visual_qa_render_evidence.ps1",
    "build_formula_quality_report.py",
    "build_workbook_controls_report.py",
    "build_external_dependency_report.py",
    "build_power_query_lineage_report.py",
    "build_power_query_refresh_report.py",
    "build_excel_bi_model_report.py",
    "build_cube_dependency_report.py",
    "build_vba_button_binding_report.py",
    "test_excel_ado_sql_access.ps1",
    "test_excel_adomd_query.ps1",
    "run_release_gate.py",
    "run_release_gate.sh",
    "run_task_profile.py",
    "run_case_regression.py",
    "create_visual_qa_fixture.py",
    "deploy-local-plugin.py",
}

WORKFLOWS = [
    {
        "id": "route-then-triage",
        "title": "Route and triage a mixed Excel BI workbook",
        "skills": ["excel-bi-router", "excel-vba-workbook-engineering"],
        "tools": [
            "route_excel_bi_task.py",
            "inspect_excel_bi_workbook.py",
            "build_formula_quality_report.py",
            "build_workbook_controls_report.py",
            "build_external_dependency_report.py",
            "build_workbook_triage_report.py",
        ],
        "boundary": "Static triage routes work and finds structural risks; it does not prove Excel runtime calculation or refresh.",
    },
    {
        "id": "power-query-lifecycle",
        "title": "Power Query M edit, refresh, lineage, and refresh-report workflow",
        "skills": ["power-query-m-engineering"],
        "tools": [
            "manage_power_queries_excel_com.ps1",
            "refresh_power_queries_excel_com.ps1",
            "build_power_query_lineage_report.py",
            "build_power_query_refresh_report.py",
            "classify_power_query_refresh_errors.py",
        ],
        "boundary": "Query editing and refresh wait need desktop Excel; static lineage works from exported M source.",
    },
    {
        "id": "power-pivot-cube",
        "title": "Power Pivot Data Model, DAX, CUBE, and MDX trace workflow",
        "skills": ["power-pivot-dax-modeling", "mdx-cubevalue-extraction"],
        "tools": [
            "inspect_excel_data_model_com.ps1",
            "build_excel_bi_model_report.py",
            "build_cube_dependency_report.py",
            "lint_dax_compat.py",
            "analyze_dax_dependencies.py",
            "analyze_measure_rename_impact.py",
        ],
        "boundary": "OpenXML can inspect CUBE formula text; Data Model metadata and DAX behavior require Excel COM or a live model export.",
    },
    {
        "id": "pure-deliverable",
        "title": "Pure-value deliverable cleanup and verification workflow",
        "skills": ["excel-deliverable-publisher", "excel-vba-workbook-engineering", "power-query-m-engineering"],
        "tools": [
            "build_external_dependency_report.py",
            "build_pure_deliverable_cleanup_plan.py",
            "build_pure_deliverable_verification_report.py",
        ],
        "boundary": "The cleanup plan is non-destructive; a copied workbook still needs refresh, value-freeze, and post-clean audit.",
    },
    {
        "id": "workbook-qa-audit",
        "title": "Pre-delivery workbook QA audit workflow",
        "skills": ["excel-workbook-qa-auditor", "excel-vba-workbook-engineering"],
        "tools": [
            "inspect_excel_bi_workbook.py",
            "build_workbook_triage_report.py",
            "build_visual_qa_report.py",
            "export_visual_qa_render_evidence.ps1",
            "build_formula_quality_report.py",
            "build_workbook_controls_report.py",
            "build_external_dependency_report.py",
        ],
        "boundary": "Static QA reports identify workbook risks; live calculation, refresh, VBA, and Data Model behavior still need Excel runtime evidence.",
    },
    {
        "id": "office-environment-diagnostics",
        "title": "Office, provider, and automation environment diagnostics workflow",
        "skills": ["office-environment-diagnostics", "excel-ado-sql-data-access"],
        "tools": [
            "probe_excel_bi_providers.ps1",
            "build_provider_environment_report.py",
            "create_provider_environment_fixture.py",
            "test_excel_adomd_query.ps1",
        ],
        "boundary": "Environment probes prove local capability and provider readiness, not workbook business correctness.",
    },
    {
        "id": "excel-report-build",
        "title": "Excel report workbook build and validation workflow",
        "skills": ["excel-report-builder", "excel-vba-workbook-engineering", "excel-deliverable-publisher"],
        "tools": [
            "inspect_excel_bi_workbook.py",
            "build_visual_qa_report.py",
            "export_visual_qa_render_evidence.ps1",
            "build_formula_quality_report.py",
            "build_workbook_controls_report.py",
            "build_workbook_triage_report.py",
        ],
        "boundary": "Report surface construction should be validated before publish cleanup; static generation does not prove Excel recalculation.",
    },
    {
        "id": "power-bi-semantic-model-review",
        "title": "Power BI semantic model and Excel Power Pivot portability workflow",
        "skills": ["power-bi-semantic-model", "power-pivot-dax-modeling"],
        "tools": [
            "search_official_docs.py",
            "lint_dax_compat.py",
            "analyze_dax_dependencies.py",
            "build_excel_bi_model_report.py",
        ],
        "boundary": "Power BI semantic models and Excel Power Pivot share concepts but are not identical runtime hosts.",
    },
    {
        "id": "sanitized-fixture-regression",
        "title": "Sanitized Excel BI fixture and regression workflow",
        "skills": ["excel-testing-fixtures"],
        "tools": [
            "build_sanitized_fixture_bundle.py",
            "run_case_regression.py",
            "create_visual_qa_fixture.py",
            "build_visual_qa_report.py",
            "export_visual_qa_render_evidence.ps1",
            "create_cube_formula_fixture.py",
            "create_workbook_surface_fixture.py",
            "build_cross_agent_forward_test_pack.py",
        ],
        "boundary": "Fixtures prove designed parser/report paths and scorer mechanics, not every customer workbook shape.",
    },
    {
        "id": "cross-agent-release",
        "title": "Cross-agent distribution, forward-test handoff, and release evidence workflow",
        "skills": ["excel-bi-router"],
        "tools": [
            "sync-skills.py",
            "build_cross_agent_forward_test_pack.py",
            "build_cross_agent_forward_test_runbook.py",
            "build_cross_agent_forward_test_handoff_bundle.py",
            "score_cross_agent_forward_test_results.py",
            "build_cross_agent_response_collection_report.py",
            "run_release_gate.py",
            "build_release_evidence_bundle.py",
            "build_completion_readiness_audit.py",
        ],
        "boundary": "Generated prompts and stubs are collection infrastructure; actual external-agent proof requires saved fresh-session responses and a passing scorer report.",
    },
    {
        "id": "task-profile-entrypoint",
        "title": "Common task profile command-plan workflow",
        "skills": [
            "excel-bi-router",
            "excel-deliverable-publisher",
            "excel-workbook-qa-auditor",
            "office-environment-diagnostics",
            "excel-report-builder",
        ],
        "tools": [
            "run_task_profile.py",
            "inspect_excel_bi_workbook.py",
            "build_workbook_triage_report.py",
            "build_pure_deliverable_cleanup_plan.py",
            "run_release_gate.py",
        ],
        "boundary": "Task profiles generate repeatable command plans for common workflows; specialist skills still own workbook-specific judgment and runtime boundaries.",
    },
    {
        "id": "real-sanitized-case-regression",
        "title": "Real/sanitized Excel BI case regression library workflow",
        "skills": ["excel-testing-fixtures", "excel-bi-router"],
        "tools": [
            "run_case_regression.py",
            "create_visual_qa_fixture.py",
            "build_visual_qa_report.py",
            "export_visual_qa_render_evidence.ps1",
            "run_task_profile.py",
            "build_capability_catalog.py",
        ],
        "boundary": "The case library tracks sanitized problem shapes and fixture-backed validation plans. It does not store or prove private customer workbooks.",
    },
    {
        "id": "visual-qa-fixture-regression",
        "title": "Workbook-backed sanitized visual QA fixture workflow",
        "skills": ["excel-testing-fixtures", "excel-workbook-qa-auditor", "excel-report-builder"],
        "tools": [
            "create_visual_qa_fixture.py",
            "build_visual_qa_report.py",
            "export_visual_qa_render_evidence.ps1",
            "inspect_excel_bi_workbook.py",
            "build_workbook_triage_report.py",
        ],
        "boundary": "Static OpenXML visual QA finds workbook-backed readability risks; rendered screenshots or PDF output still require Excel runtime or manual review.",
    },
    {
        "id": "rendered-visual-qa-evidence",
        "title": "Rendered Visual QA evidence chain workflow",
        "skills": ["excel-vba-workbook-engineering", "excel-workbook-qa-auditor", "excel-report-builder"],
        "tools": [
            "export_visual_qa_render_evidence.ps1",
            "create_visual_qa_fixture.py",
            "build_visual_qa_report.py",
            "run_release_gate.py",
        ],
        "boundary": "Windows Excel COM can export sanitized report sheets to PDF as runtime evidence; generated PDFs stay in temp/task-local paths and are not pixel-comparison proof.",
    },
]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_md(value: Any) -> str:
    return str(value if value is not None else "").replace("|", "\\|").replace("\n", " ")


def parse_frontmatter(text: str) -> dict[str, str]:
    if not text.startswith("---"):
        return {}
    parts = text.split("---", 2)
    if len(parts) < 3:
        return {}
    frontmatter = parts[1]
    result: dict[str, str] = {}
    current_key = ""
    for raw_line in frontmatter.splitlines():
        line = raw_line.rstrip()
        if not line.strip():
            continue
        if re.match(r"^[A-Za-z0-9_-]+:", line):
            key, value = line.split(":", 1)
            current_key = key.strip()
            result[current_key] = value.strip().strip("'\"")
        elif current_key:
            result[current_key] += " " + line.strip().strip("'\"")
    return result


def python_docstring(path: Path) -> str:
    try:
        module = ast.parse(path.read_text(encoding="utf-8-sig"))
        doc = ast.get_docstring(module) or ""
    except Exception:
        return ""
    return first_sentence(doc)


def first_sentence(text: str) -> str:
    text = " ".join(str(text or "").strip().split())
    if not text:
        return ""
    match = re.search(r"(?<=[.!?])\s+", text)
    if match:
        return text[: match.start()].strip()
    return text[:240]


def shell_description(path: Path) -> str:
    try:
        lines = path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    except OSError:
        return ""
    comments: list[str] = []
    for line in lines[:40]:
        stripped = line.strip()
        if not stripped or stripped.startswith("#!"):
            continue
        if stripped.startswith("#"):
            comments.append(stripped.lstrip("#").strip())
            continue
        if stripped.startswith(".SYNOPSIS") or stripped.startswith("<#"):
            continue
        if comments:
            break
    return first_sentence(" ".join(comments))


def script_domain(name: str) -> str:
    lower = name.lower()
    if "power_query" in lower or "power-query" in lower:
        return "Power Query M"
    if "dax" in lower or "model" in lower or "measure" in lower:
        return "Power Pivot DAX"
    if "cube" in lower or "mdx" in lower or "adomd" in lower:
        return "MDX/CUBE"
    if "ado" in lower or "provider" in lower or "sql" in lower:
        return "ADO/SQL"
    if "vba" in lower or "workbook" in lower or "formula" in lower or "external" in lower or "pure" in lower:
        return "Workbook/VBA"
    if "cross_agent" in lower or "sync" in lower or "release" in lower or "goal" in lower or "catalog" in lower:
        return "Cross-agent/validation"
    if "official" in lower or "docs" in lower:
        return "Documentation"
    return "General"


def discover_skills(project_root: Path) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    skill_root = project_root / ".agents" / "skills"
    for skill_dir in sorted(path for path in skill_root.iterdir() if path.is_dir()):
        skill_md = skill_dir / "SKILL.md"
        if not skill_md.is_file():
            continue
        text = skill_md.read_text(encoding="utf-8-sig")
        meta = parse_frontmatter(text)
        refs = sorted(str(path.relative_to(skill_dir)).replace("\\", "/") for path in (skill_dir / "references").rglob("*") if path.is_file()) if (skill_dir / "references").is_dir() else []
        scripts = sorted(str(path.relative_to(skill_dir)).replace("\\", "/") for path in (skill_dir / "scripts").rglob("*") if path.is_file()) if (skill_dir / "scripts").is_dir() else []
        official_indexes = [path for path in (skill_dir / "references").rglob("official-docs-index.json")] if (skill_dir / "references").is_dir() else []
        official_entries = 0
        for index in official_indexes:
            try:
                official_entries += len(load_json(index).get("entries", []))
            except Exception:
                pass
        result.append(
            {
                "name": meta.get("name", skill_dir.name),
                "path": str(skill_dir.relative_to(project_root)).replace("\\", "/"),
                "description": meta.get("description", ""),
                "referenceCount": len(refs),
                "scriptCount": len(scripts),
                "officialDocIndexCount": len(official_indexes),
                "officialDocEntryCount": official_entries,
                "references": refs,
                "scripts": scripts,
            }
        )
    return result


def discover_script_files(project_root: Path) -> list[Path]:
    roots = [project_root / "tools"]
    skills_root = project_root / ".agents" / "skills"
    if skills_root.is_dir():
        roots.extend(path / "scripts" for path in skills_root.iterdir() if (path / "scripts").is_dir())
    paths: list[Path] = []
    for root in roots:
        if not root.is_dir():
            continue
        paths.extend(path for path in root.iterdir() if path.is_file())
    return sorted(paths, key=lambda item: str(item.relative_to(project_root)).lower())


def discover_tools(project_root: Path) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in discover_script_files(project_root):
        suffix = path.suffix.lower()
        if suffix not in {".py", ".ps1", ".sh"}:
            continue
        if suffix == ".py":
            description = python_docstring(path)
            runtime = "python"
        elif suffix == ".ps1":
            description = shell_description(path)
            runtime = "powershell"
        else:
            description = shell_description(path)
            runtime = "bash"
        result.append(
            {
                "name": path.name,
                "path": str(path.relative_to(project_root)).replace("\\", "/"),
                "runtime": runtime,
                "domain": script_domain(path.name),
                "description": description,
                "required": path.name in REQUIRED_TOOLS,
            }
        )
    return result


def discover_official_docs(skills: list[dict[str, Any]]) -> dict[str, Any]:
    total_indexes = sum(int(skill["officialDocIndexCount"]) for skill in skills)
    total_entries = sum(int(skill["officialDocEntryCount"]) for skill in skills)
    by_skill = {
        str(skill["name"]): {
            "indexCount": skill["officialDocIndexCount"],
            "entryCount": skill["officialDocEntryCount"],
        }
        for skill in skills
        if int(skill["officialDocIndexCount"]) > 0
    }
    return {
        "indexCount": total_indexes,
        "entryCount": total_entries,
        "bySkill": by_skill,
    }


def parse_release_gate_checks(project_root: Path) -> list[dict[str, Any]]:
    source = (project_root / "tools" / "run_release_gate.py").read_text(encoding="utf-8")
    definitions = re.findall(r"^def\s+([A-Za-z0-9_]+)\(", source, flags=re.MULTILINE)
    checks = [
        name
        for name in definitions
        if name.endswith("_check") or name.endswith("_fixture_check") or name.endswith("_report_check")
    ]
    return [
        {
            "function": name,
            "domain": script_domain(name),
        }
        for name in sorted(checks)
    ]


def build_catalog(project_root: Path) -> dict[str, Any]:
    manifest = load_json(project_root / ".codex-plugin" / "plugin.json")
    skills = discover_skills(project_root)
    tools = discover_tools(project_root)
    official_docs = discover_official_docs(skills)
    gate_checks = parse_release_gate_checks(project_root)

    skill_names = {str(skill["name"]) for skill in skills}
    tool_names = {str(tool["name"]) for tool in tools}
    missing_skills = sorted(REQUIRED_SKILLS - skill_names)
    missing_tools = sorted(REQUIRED_TOOLS - tool_names)
    status = "pass"
    findings: list[dict[str, Any]] = []
    if missing_skills:
        status = "fail"
        findings.append({"severity": "high", "code": "missing-required-skills", "items": missing_skills})
    if missing_tools:
        status = "fail"
        findings.append({"severity": "high", "code": "missing-required-tools", "items": missing_tools})
    if official_docs["indexCount"] < 4 or official_docs["entryCount"] < 20:
        status = "fail"
        findings.append(
            {
                "severity": "high",
                "code": "insufficient-official-doc-indexes",
                "indexCount": official_docs["indexCount"],
                "entryCount": official_docs["entryCount"],
            }
        )
    if len(gate_checks) < 20:
        status = "fail"
        findings.append({"severity": "high", "code": "insufficient-release-gate-checks", "count": len(gate_checks)})

    by_domain: dict[str, int] = {}
    for tool in tools:
        domain = str(tool["domain"])
        by_domain[domain] = by_domain.get(domain, 0) + 1

    return {
        "generatedAt": now_iso(),
        "status": status,
        "plugin": {
            "name": manifest.get("name", ""),
            "version": manifest.get("version", ""),
            "description": manifest.get("description", ""),
        },
        "summary": {
            "skillCount": len(skills),
            "toolCount": len(tools),
            "officialDocIndexCount": official_docs["indexCount"],
            "officialDocEntryCount": official_docs["entryCount"],
            "releaseGateCheckCount": len(gate_checks),
            "toolCountByDomain": dict(sorted(by_domain.items())),
        },
        "skills": skills,
        "tools": tools,
        "officialDocs": official_docs,
        "workflows": WORKFLOWS,
        "releaseGateChecks": gate_checks,
        "validationCommands": [
            "python tools/run_release_gate.py --project-root .",
            "tools/run_release_gate.sh --profile structural",
            "python tools/build_release_evidence_bundle.py --project-root . --release-gate-json <release_gate.json> --require-pass",
            "python tools/build_completion_readiness_audit.py --project-root . --require-pass",
            "python tools/build_cross_agent_response_collection_report.py --manifest-json <forward-test-pack.json> --responses-dir <responses-dir> --out-json <collection.json>",
            "python tools/validate_project_docs.py --project-root .",
            "python tools/validate_github_community_health.py --project-root .",
            "python tools/validate_task_recipes.py --project-root .",
        ],
        "findings": findings,
        "boundaries": [
            "The catalog is static discovery metadata; it does not validate a workbook.",
            "Excel runtime behavior still requires release-gate or task-specific validation evidence.",
            "Official documentation indexes are local routing metadata unless online drift sampling is explicitly run.",
        ],
    }


def table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(clean_md(value) for value in row) + " |")
    return lines


def render_markdown(catalog: dict[str, Any]) -> str:
    summary = catalog["summary"]
    lines: list[str] = [
        "# Excel BI Capability Catalog",
        "",
        f"- status: `{catalog['status']}`",
        f"- plugin: `{catalog['plugin']['name']}`",
        f"- version: `{catalog['plugin']['version']}`",
        f"- skills: {summary['skillCount']}",
        f"- tools: {summary['toolCount']}",
        f"- official docs: {summary['officialDocIndexCount']} indexes / {summary['officialDocEntryCount']} entries",
        f"- release-gate checks: {summary['releaseGateCheckCount']}",
        "",
        "## Skills",
        "",
    ]
    lines.extend(
        table(
            ["Skill", "References", "Scripts", "Official docs", "Description"],
            [
                [
                    skill["name"],
                    skill["referenceCount"],
                    skill["scriptCount"],
                    skill["officialDocEntryCount"],
                    skill["description"],
                ]
                for skill in catalog["skills"]
            ],
        )
    )
    lines.extend(["", "## Tool Domains", ""])
    lines.extend(table(["Domain", "Tool count"], [[key, value] for key, value in catalog["summary"]["toolCountByDomain"].items()]))
    lines.extend(["", "## Core Workflows", ""])
    lines.extend(
        table(
            ["Workflow", "Skills", "Representative tools", "Boundary"],
            [
                [
                    workflow["title"],
                    ", ".join(workflow["skills"]),
                    ", ".join(workflow["tools"][:4]),
                    workflow["boundary"],
                ]
                for workflow in catalog["workflows"]
            ],
        )
    )
    lines.extend(["", "## Required Validation Commands", ""])
    for command in catalog["validationCommands"]:
        lines.append(f"- `{command}`")
    lines.extend(["", "## Boundary", ""])
    for boundary in catalog["boundaries"]:
        lines.append(f"- {boundary}")
    if catalog["findings"]:
        lines.extend(["", "## Findings", ""])
        lines.extend(table(["Severity", "Code", "Items"], [[item.get("severity"), item.get("code"), ", ".join(item.get("items", []))] for item in catalog["findings"]]))
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a static capability catalog for the Excel BI agent pack.")
    parser.add_argument("--project-root", type=Path, default=Path("."))
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-md", type=Path)
    parser.add_argument("--print", action="store_true", help="Print Markdown catalog")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero unless catalog status is pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = args.project_root.resolve()
    catalog = build_catalog(project_root)
    markdown = render_markdown(catalog)
    if args.out_json:
        write_json(args.out_json, catalog)
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(markdown, encoding="utf-8")
    if args.print:
        print(markdown)
    elif not args.out_json and not args.out_md:
        print(json.dumps(catalog, ensure_ascii=False, indent=2))
    if args.require_pass and catalog["status"] != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
