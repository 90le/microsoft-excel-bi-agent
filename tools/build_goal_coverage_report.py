#!/usr/bin/env python3
"""Build an auditable coverage report for the active Excel BI plugin goal.

This report checks that each core goal area has concrete package evidence:

- expected source files or generated mirrors exist
- validation documentation records the relevant gate evidence
- completion-evidence documentation maps the area to package artifacts

It is intentionally a coverage audit, not a substitute for running the release
gate. Runtime behavior still needs the release gate and Excel COM fixtures.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"


@dataclass(frozen=True)
class AreaRequirement:
    key: str
    title: str
    required_files: tuple[str, ...]
    validation_terms: tuple[str, ...]
    completion_terms: tuple[str, ...]
    project_doc_terms: tuple[str, ...] = ()


AREA_REQUIREMENTS: tuple[AreaRequirement, ...] = (
    AreaRequirement(
        key="plugin-packaging",
        title="Plugin packaging and install flow",
        required_files=(
            ".codex-plugin/plugin.json",
            "tools/deploy-local-plugin.py",
        ),
        validation_terms=(
            "Source plugin validation",
            "Local plugin copy validation",
            "Installed cache validation",
            "Installed plugin enabled",
        ),
        completion_terms=("Plugin packaging and install flow",),
        project_doc_terms=("cachebuster reinstall flow",),
    ),
    AreaRequirement(
        key="cross-agent-distribution",
        title="Cross-agent skill distribution",
        required_files=(
            ".agents/skills/excel-bi-router/SKILL.md",
            "skills/excel-bi-router/SKILL.md",
            ".claude/skills/excel-bi-router/SKILL.md",
            ".opencode/skills/excel-bi-router/SKILL.md",
            "tools/sync-skills.py",
        ),
        validation_terms=(
            "Cross-agent mirror drift",
            "Task recipe documentation validation",
            "Skill UI metadata validation",
        ),
        completion_terms=("Cross-agent skill distribution",),
        project_doc_terms=("generated Codex `skills/` mirror",),
    ),
    AreaRequirement(
        key="excel-vba-workbook-engineering",
        title="Excel/VBA workbook engineering",
        required_files=(
            ".agents/skills/excel-vba-workbook-engineering/SKILL.md",
            ".agents/skills/excel-vba-workbook-engineering/scripts/export_vba.ps1",
            ".agents/skills/excel-vba-workbook-engineering/scripts/import_vba.ps1",
            ".agents/skills/excel-vba-workbook-engineering/scripts/inspect_workbook.ps1",
            ".agents/skills/excel-vba-workbook-engineering/scripts/lint_vba_source.py",
        ),
        validation_terms=(
            "Excel workbook COM inventory fixture smoke",
            "VBA import/export/run fixture smoke",
            "VBA source lint fixture smoke",
            "External dependency OpenXML fixture smoke",
        ),
        completion_terms=("Reusable Excel/VBA workbook engineering",),
        project_doc_terms=("Excel workbook and VBA engineering",),
    ),
    AreaRequirement(
        key="power-query-m",
        title="Power Query M authoring, lifecycle control, refresh, diagnostics",
        required_files=(
            ".agents/skills/power-query-m-engineering/SKILL.md",
            ".agents/skills/power-query-m-engineering/scripts/manage_power_queries_excel_com.ps1",
            ".agents/skills/power-query-m-engineering/scripts/refresh_power_queries_excel_com.ps1",
            ".agents/skills/power-query-m-engineering/scripts/lint_power_query_m.py",
            ".agents/skills/power-query-m-engineering/scripts/classify_power_query_refresh_errors.py",
            ".agents/skills/power-query-m-engineering/references/official-docs-index.json",
        ),
        validation_terms=(
            "Power Query M lint fixture smoke",
            "Power Query live refresh fixture smoke",
            "Power Query refresh error classifier fixture smoke",
            "Official documentation index validation",
        ),
        completion_terms=("Power Query M authoring, lifecycle control, refresh, diagnostics",),
        project_doc_terms=("Power Query M authoring",),
    ),
    AreaRequirement(
        key="power-pivot-dax",
        title="Power Pivot Data Model and DAX modeling",
        required_files=(
            ".agents/skills/power-pivot-dax-modeling/SKILL.md",
            ".agents/skills/power-pivot-dax-modeling/scripts/lint_dax_compat.py",
            ".agents/skills/power-pivot-dax-modeling/scripts/analyze_dax_dependencies.py",
            ".agents/skills/power-pivot-dax-modeling/references/official-docs-index.json",
            "tools/inspect_excel_data_model_com.ps1",
            "tools/build_excel_bi_model_report.py",
            "tools/analyze_measure_rename_impact.py",
            "tools/build_measure_rename_rewrite_plan.py",
        ),
        validation_terms=(
            "Generic Power Pivot model-report fixture smoke",
            "DAX compatibility lint fixture smoke",
            "DAX dependency analysis fixture smoke",
            "Measure rename impact fixture smoke",
        ),
        completion_terms=("Power Pivot DAX modeling",),
        project_doc_terms=("Power Pivot Data Model and DAX modeling",),
    ),
    AreaRequirement(
        key="mdx-cube",
        title="MDX/CUBE formula extraction and report-layer tracing",
        required_files=(
            ".agents/skills/mdx-cubevalue-extraction/SKILL.md",
            ".agents/skills/mdx-cubevalue-extraction/references/official-docs-index.json",
            "tools/create_cube_formula_fixture.py",
            "tools/build_cube_dependency_report.py",
            "tools/mdx_references.py",
            "tools/build_measure_rename_rewrite_plan.py",
        ),
        validation_terms=(
            "Generic CUBE formula fixture smoke",
            "Measure rename rewrite plan fixture smoke",
            "Measure delete rewrite plan fixture smoke",
            "Escaped MDX measure reference fixture smoke",
        ),
        completion_terms=("MDX/CUBE formula extraction and report-layer tracing",),
        project_doc_terms=("MDX/CUBE formula extraction",),
    ),
    AreaRequirement(
        key="ado-sql-adomd",
        title="VBA ADO/OLEDB/ADOMD/SQL data access",
        required_files=(
            ".agents/skills/excel-ado-sql-data-access/SKILL.md",
            ".agents/skills/excel-ado-sql-data-access/references/official-docs-index.json",
            "tools/probe_excel_bi_providers.ps1",
            "tools/test_excel_ado_sql_access.ps1",
            "tools/test_excel_adomd_query.ps1",
        ),
        validation_terms=(
            "Provider probe fixture smoke",
            "ADO workbook SQL fixture smoke",
            "ADOMD COM probe fixture smoke",
        ),
        completion_terms=("VBA ADO/OLEDB/ADOMD/SQL data access",),
        project_doc_terms=("VBA ADO/OLEDB/ADOMD/SQL data access",),
    ),
    AreaRequirement(
        key="official-docs",
        title="Official documentation routing and local knowledge indexes",
        required_files=(
            "tools/search_official_docs.py",
            "tools/validate_official_docs_index.py",
            ".agents/skills/power-query-m-engineering/references/official-docs-index.json",
            ".agents/skills/power-pivot-dax-modeling/references/official-docs-index.json",
            ".agents/skills/mdx-cubevalue-extraction/references/official-docs-index.json",
            ".agents/skills/excel-ado-sql-data-access/references/official-docs-index.json",
        ),
        validation_terms=(
            "Official documentation index validation",
            "Optional online official URL sample",
        ),
        completion_terms=("Official documentation routing and local knowledge indexes",),
        project_doc_terms=("Official documentation routing",),
    ),
    AreaRequirement(
        key="cross-platform-boundaries",
        title="Windows PowerShell, Git Bash, Linux, and macOS compatibility boundaries",
        required_files=(
            "tools/run_release_gate.sh",
            "tools/invoke_excel_bi_com.sh",
            ".agents/skills/excel-vba-workbook-engineering/scripts/invoke_excel_com.sh",
            ".agents/skills/power-query-m-engineering/scripts/invoke_power_query_excel_com.sh",
            "docs/compatibility.md",
        ),
        validation_terms=(
            "Portable release gate wrapper",
            "Structural release gate profile",
            "Git Bash and portable wrapper syntax",
            "PowerShell script syntax",
            "checked 5 Bash scripts",
        ),
        completion_terms=(
            "Windows PowerShell compatibility",
            "Windows Git Bash compatibility",
            "Linux/macOS compatibility boundaries",
        ),
        project_doc_terms=("Linux, and macOS compatibility boundaries",),
    ),
    AreaRequirement(
        key="validation-workflow",
        title="One-command validation workflow",
        required_files=(
            "tools/run_release_gate.py",
            "tools/run_release_gate.sh",
            "tools/validate_project_docs.py",
            "tools/validate_task_recipes.py",
            "tools/validate-skills.py",
        ),
        validation_terms=(
            "Automated release gate runner",
            "Project documentation consistency validation",
            "Task recipe documentation validation",
            "Python script compile",
            "Python cache cleanup",
        ),
        completion_terms=("Validation workflow",),
        project_doc_terms=("One-command release gate",),
    ),
    AreaRequirement(
        key="goal-tracking-docs",
        title="Goal tracking and project documentation",
        required_files=(
            "docs/master-goal.md",
            "docs/goal-tracking.md",
            "docs/iteration-protocol.md",
            "docs/progress.md",
            "docs/validation.md",
            "docs/completion-evidence.md",
            "docs/task-recipes.md",
            "docs/project.md",
        ),
        validation_terms=(
            "Goal tracking document binding",
            "Sanitized recipe docs",
            "Project documentation consistency validation",
        ),
        completion_terms=("Project tracking docs",),
        project_doc_terms=("docs/goal-tracking.md",),
    ),
    AreaRequirement(
        key="upper-layer-scenario-skills",
        title="Upper-layer scenario skill expansion",
        required_files=(
            ".agents/skills/excel-deliverable-publisher/SKILL.md",
            ".agents/skills/excel-deliverable-publisher/agents/openai.yaml",
            ".agents/skills/excel-workbook-qa-auditor/SKILL.md",
            ".agents/skills/excel-workbook-qa-auditor/agents/openai.yaml",
            ".agents/skills/office-environment-diagnostics/SKILL.md",
            ".agents/skills/office-environment-diagnostics/agents/openai.yaml",
            ".agents/skills/excel-report-builder/SKILL.md",
            ".agents/skills/excel-report-builder/agents/openai.yaml",
            ".agents/skills/power-bi-semantic-model/SKILL.md",
            ".agents/skills/power-bi-semantic-model/agents/openai.yaml",
            ".agents/skills/excel-testing-fixtures/SKILL.md",
            ".agents/skills/excel-testing-fixtures/agents/openai.yaml",
            ".agents/skills/excel-bi-router/scripts/route_excel_bi_task.py",
            "tools/build_capability_catalog.py",
            "tools/build_cross_agent_forward_test_pack.py",
        ),
        validation_terms=(
            "Upper-layer scenario skills expansion",
            "12 canonical skills",
            "48 forward-test prompts",
        ),
        completion_terms=(
            "Excel deliverable publishing",
            "Excel workbook QA auditing",
            "Office environment diagnostics",
            "Excel report building",
            "Power BI semantic model review",
            "Excel testing fixtures",
        ),
        project_doc_terms=("six upper-layer scenario skills",),
    ),
    AreaRequirement(
        key="real-sanitized-case-regression",
        title="Real/sanitized case regression library V1",
        required_files=(
            "fixtures/real-sanitized-cases/manifest.json",
            "fixtures/real-sanitized-cases/cases/pq-folder-dynamic-expand-order.json",
            "fixtures/real-sanitized-cases/cases/dax-excel-powerpivot-compat.json",
            "fixtures/real-sanitized-cases/cases/cube-zero-result-debug.json",
            "fixtures/real-sanitized-cases/cases/vba-button-binding-runtime.json",
            "fixtures/real-sanitized-cases/cases/deliverable-clean-copy.json",
            "fixtures/real-sanitized-cases/cases/visual-report-readability.json",
            "tools/run_case_regression.py",
            "docs/real-case-regression.md",
        ),
        validation_terms=(
            "Real/sanitized case regression library V1",
            "run_case_regression.py --require-pass",
            "case-regression profile execution",
            "Real/sanitized case regression library smoke",
        ),
        completion_terms=("Real/sanitized case regression library V1",),
        project_doc_terms=("real/sanitized case regression library",),
    ),
    AreaRequirement(
        key="workbook-backed-visual-qa",
        title="Workbook-backed sanitized Visual QA case V1",
        required_files=(
            "tools/create_visual_qa_fixture.py",
            "tools/build_visual_qa_report.py",
            "fixtures/real-sanitized-cases/cases/visual-report-readability.json",
            "docs/real-case-regression.md",
            "docs/task-recipes.md",
        ),
        validation_terms=(
            "Workbook-backed sanitized Visual QA case V1",
            "Visual QA report fixture smoke",
            "blocked-for-delivery",
        ),
        completion_terms=("Workbook-backed sanitized Visual QA case V1",),
        project_doc_terms=("workbook-backed sanitized Visual QA case V1",),
    ),
    AreaRequirement(
        key="rendered-visual-qa-evidence",
        title="Rendered Visual QA evidence chain V1",
        required_files=(
            "tools/export_visual_qa_render_evidence.ps1",
            "tools/create_visual_qa_fixture.py",
            "tools/build_visual_qa_report.py",
            "tools/run_release_gate.py",
            "docs/task-recipes.md",
        ),
        validation_terms=(
            "Rendered Visual QA evidence chain V1",
            "Visual QA render evidence smoke",
            "Windows Excel COM PDF export",
        ),
        completion_terms=("Rendered Visual QA evidence chain V1",),
        project_doc_terms=("rendered Visual QA evidence chain V1",),
    ),
)


def read_text(path: Path) -> str:
    if not path.is_file():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def load_manifest_version(project_root: Path) -> str:
    manifest = project_root / ".codex-plugin" / "plugin.json"
    if not manifest.is_file():
        return ""
    try:
        data = json.loads(manifest.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return ""
    return str(data.get("version", "")).strip()


def check_terms(text: str, terms: tuple[str, ...]) -> list[str]:
    return [term for term in terms if term not in text]


def check_files(project_root: Path, files: tuple[str, ...]) -> list[str]:
    return [path for path in files if not (project_root / Path(path)).is_file()]


def evaluate_area(
    project_root: Path,
    requirement: AreaRequirement,
    validation_text: str,
    completion_text: str,
    project_text: str,
    master_goal_text: str,
) -> dict[str, Any]:
    doc_text = "\n".join([project_text, master_goal_text])
    missing_files = check_files(project_root, requirement.required_files)
    missing_validation_terms = check_terms(validation_text, requirement.validation_terms)
    missing_completion_terms = check_terms(completion_text, requirement.completion_terms)
    missing_project_terms = check_terms(doc_text, requirement.project_doc_terms)
    status = PASS if not (missing_files or missing_validation_terms or missing_completion_terms or missing_project_terms) else FAIL
    return {
        "key": requirement.key,
        "title": requirement.title,
        "status": status,
        "requiredFileCount": len(requirement.required_files),
        "validationTermCount": len(requirement.validation_terms),
        "completionTermCount": len(requirement.completion_terms),
        "projectDocTermCount": len(requirement.project_doc_terms),
        "missingFiles": missing_files,
        "missingValidationTerms": missing_validation_terms,
        "missingCompletionTerms": missing_completion_terms,
        "missingProjectDocTerms": missing_project_terms,
    }


def build_report(project_root: Path) -> dict[str, Any]:
    validation_text = read_text(project_root / "docs" / "validation.md")
    completion_text = read_text(project_root / "docs" / "completion-evidence.md")
    project_text = read_text(project_root / "docs" / "project.md")
    master_goal_text = read_text(project_root / "docs" / "master-goal.md")
    version = load_manifest_version(project_root)

    areas = [
        evaluate_area(project_root, requirement, validation_text, completion_text, project_text, master_goal_text)
        for requirement in AREA_REQUIREMENTS
    ]
    failed = [area for area in areas if area["status"] != PASS]
    report = {
        "projectRoot": str(project_root),
        "version": version,
        "status": PASS if not failed else FAIL,
        "areaCount": len(areas),
        "passedAreaCount": len(areas) - len(failed),
        "failedAreaCount": len(failed),
        "areas": areas,
        "limitations": [
            "This coverage report checks evidence presence and documentation alignment only.",
            "Runtime behavior still requires the release gate and Excel COM smoke fixtures.",
            "A pass here does not mean the long-running goal is complete; it means current evidence covers the named core areas.",
        ],
    }
    return report


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Goal Coverage Report",
        "",
        f"- project: `{report.get('projectRoot', '')}`",
        f"- version: `{report.get('version', '')}`",
        f"- status: **{report.get('status', '')}**",
        f"- areas: {report.get('areaCount', 0)}",
        f"- passed: {report.get('passedAreaCount', 0)}",
        f"- failed: {report.get('failedAreaCount', 0)}",
        "",
        "| Area | Status | Missing evidence |",
        "|---|---:|---|",
    ]
    for area in report.get("areas", []):
        missing: list[str] = []
        for key, label in [
            ("missingFiles", "files"),
            ("missingValidationTerms", "validation"),
            ("missingCompletionTerms", "completion"),
            ("missingProjectDocTerms", "project docs"),
        ]:
            values = area.get(key, [])
            if values:
                missing.append(f"{label}: {', '.join(str(item) for item in values)}")
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(area.get("title", "")),
                    clean_markdown(area.get("status", "")),
                    clean_markdown("; ".join(missing) if missing else "none"),
                ]
            )
            + " |"
        )
    lines.extend(["", "## Limitations", ""])
    for item in report.get("limitations", []):
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable coverage report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown coverage report")
    parser.add_argument("--require-pass", action="store_true", help="Exit with code 1 if coverage is incomplete")
    args = parser.parse_args()

    project_root = Path(args.project_root).expanduser().resolve()
    report = build_report(project_root)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.require_pass and report.get("status") != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
