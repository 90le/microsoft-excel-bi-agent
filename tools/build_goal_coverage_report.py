#!/usr/bin/env python3
"""Build an auditable public goal coverage report for this plugin.

The report verifies that the public maintenance goals are backed by concrete
repository files and public documentation. It intentionally avoids
maintainer-only release ledgers and machine-specific runtime evidence.
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
    public_doc_terms: tuple[str, ...]
    goal_terms: tuple[str, ...]


AREA_REQUIREMENTS: tuple[AreaRequirement, ...] = (
    AreaRequirement(
        key="install-truth",
        title="Install truth",
        required_files=("tools/install.mjs", "install.ps1", "install.cmd", "install.sh", "docs/install-and-sync.md"),
        public_doc_terms=("node tools/install.mjs", "codex plugin marketplace add 90le/microsoft-excel-bi-agent"),
        goal_terms=("Install truth", "安装真实性"),
    ),
    AreaRequirement(
        key="bilingual-docs",
        title="Bilingual independent docs",
        required_files=("README.md", "README.zh-CN.md", "docs/project.en-US.md", "docs/project.zh-CN.md", "docs/intro.html", "docs/intro.zh-CN.html"),
        public_doc_terms=("English project overview", "中文项目说明"),
        goal_terms=("Bilingual independence", "双语独立"),
    ),
    AreaRequirement(
        key="maintenance-goals",
        title="Public maintenance goals and risk backlog",
        required_files=("docs/maintenance-goals.en-US.md", "docs/maintenance-goals.zh-CN.md", "docs/maintenance-goals.md"),
        public_doc_terms=("maintenance-goals.en-US.md", "maintenance-goals.zh-CN.md"),
        goal_terms=("Risk Register", "风险清单", "Optimization Backlog", "优化 Backlog"),
    ),
    AreaRequirement(
        key="public-growth-goals",
        title="Public growth goals and maintainer trust",
        required_files=("docs/growth-goals.en-US.md", "docs/growth-goals.zh-CN.md", "docs/growth-goals.md", "README.md", "README.zh-CN.md"),
        public_doc_terms=("growth-goals.en-US.md", "growth-goals.zh-CN.md", "Qiu Binbin", "丘彬彬", "binstudy", "90le.cn"),
        goal_terms=("First-minute clarity", "一分钟清晰度", "Trust and attribution", "信任与署名"),
    ),
    AreaRequirement(
        key="marketing-readiness",
        title="Marketing copy and share readiness",
        required_files=("docs/marketing-copy.en-US.md", "docs/marketing-copy.zh-CN.md", "docs/marketing-copy.md", "assets/social-preview.png", "assets/social-preview.zh-CN.png"),
        public_doc_terms=("marketing-copy.en-US.md", "marketing-copy.zh-CN.md", "summary_large_image", "https://90le.github.io/microsoft-excel-bi-agent/assets/social-preview"),
        goal_terms=("Advertising readiness", "广告准备度", "Do Not Claim", "不要这样宣传"),
    ),
    AreaRequirement(
        key="repository-governance",
        title="Repository governance and safe public intake",
        required_files=("CONTRIBUTING.md", "SECURITY.md", ".github/pull_request_template.md", ".github/ISSUE_TEMPLATE/config.yml", ".github/ISSUE_TEMPLATE/bug_report.yml", ".github/ISSUE_TEMPLATE/feature_request.yml", ".github/ISSUE_TEMPLATE/docs_install.yml", "tools/validate_github_community_health.py", "docs/repository-governance-goals.en-US.md", "docs/repository-governance-goals.zh-CN.md", "docs/repository-governance-goals.md"),
        public_doc_terms=("CONTRIBUTING.md", "SECURITY.md", "repository-governance-goals.en-US.md", "repository-governance-goals.zh-CN.md"),
        goal_terms=("Safe issue intake", "安全 issue 入口", "PR review discipline", "PR 审阅纪律"),
    ),
    AreaRequirement(
        key="release-versioning",
        title="Release notes and version visibility",
        required_files=("docs/release-notes.en-US.md", "docs/release-notes.zh-CN.md", "docs/release-notes.md", ".codex-plugin/plugin.json"),
        public_doc_terms=("release-notes.en-US.md", "release-notes.zh-CN.md", "v0.1.5"),
        goal_terms=("release decisions", "发布判断"),
    ),
    AreaRequirement(
        key="skill-source-discipline",
        title="Skill source and mirror discipline",
        required_files=(".agents/skills/excel-bi-router/SKILL.md", "skills/excel-bi-router/SKILL.md", ".claude/skills/excel-bi-router/SKILL.md", ".opencode/skills/excel-bi-router/SKILL.md", "tools/sync-skills.py"),
        public_doc_terms=(".agents/skills/", "generated mirrors"),
        goal_terms=("Skill source discipline", "技能源纪律"),
    ),
    AreaRequirement(
        key="excel-vba-workbook-engineering",
        title="Excel/VBA workbook engineering",
        required_files=(".agents/skills/excel-vba-workbook-engineering/SKILL.md", ".agents/skills/excel-vba-workbook-engineering/scripts/export_vba.ps1", ".agents/skills/excel-vba-workbook-engineering/scripts/import_vba.ps1", ".agents/skills/excel-vba-workbook-engineering/scripts/inspect_workbook.ps1"),
        public_doc_terms=("Excel/VBA workbook engineering", "`excel-vba-workbook-engineering`"),
        goal_terms=("Windows desktop Excel", "Windows 桌面版 Excel"),
    ),
    AreaRequirement(
        key="power-query-m",
        title="Power Query M workflows",
        required_files=(".agents/skills/power-query-m-engineering/SKILL.md", ".agents/skills/power-query-m-engineering/scripts/lint_power_query_m.py", ".agents/skills/power-query-m-engineering/scripts/refresh_power_queries_excel_com.ps1"),
        public_doc_terms=("Power Query M", "`power-query-m-engineering`"),
        goal_terms=("Power Query refresh", "Power Query 刷新"),
    ),
    AreaRequirement(
        key="power-pivot-dax",
        title="Power Pivot DAX workflows",
        required_files=(".agents/skills/power-pivot-dax-modeling/SKILL.md", ".agents/skills/power-pivot-dax-modeling/scripts/lint_dax_compat.py", ".agents/skills/power-pivot-dax-modeling/scripts/analyze_dax_dependencies.py"),
        public_doc_terms=("Power Pivot DAX", "`power-pivot-dax-modeling`"),
        goal_terms=("Power Pivot runtime", "Power Pivot 运行时"),
    ),
    AreaRequirement(
        key="mdx-cube",
        title="MDX/CUBE formula workflows",
        required_files=(".agents/skills/mdx-cubevalue-extraction/SKILL.md", "tools/create_cube_formula_fixture.py", "tools/build_cube_dependency_report.py", "tools/mdx_references.py"),
        public_doc_terms=("MDX/CUBE", "`mdx-cubevalue-extraction`"),
        goal_terms=("structural validation", "结构校验"),
    ),
    AreaRequirement(
        key="ado-sql",
        title="ADO/OLEDB/ADOMD/SQL workflows",
        required_files=(".agents/skills/excel-ado-sql-data-access/SKILL.md", "tools/probe_excel_bi_providers.ps1", "tools/test_excel_ado_sql_access.ps1", "tools/test_excel_adomd_query.ps1"),
        public_doc_terms=("ADO / SQL", "`excel-ado-sql-data-access`"),
        goal_terms=("runtime behavior", "运行时行为"),
    ),
    AreaRequirement(
        key="deliverable-qa-reporting",
        title="Deliverable, QA, diagnostics, report, semantic model, and fixture skills",
        required_files=(".agents/skills/excel-deliverable-publisher/SKILL.md", ".agents/skills/excel-workbook-qa-auditor/SKILL.md", ".agents/skills/office-environment-diagnostics/SKILL.md", ".agents/skills/excel-report-builder/SKILL.md", ".agents/skills/power-bi-semantic-model/SKILL.md", ".agents/skills/excel-testing-fixtures/SKILL.md"),
        public_doc_terms=("Client deliverables", "Workbook QA", "`excel-report-builder`", "`excel-testing-fixtures`"),
        goal_terms=("Artifact hygiene", "Artifact hygiene"),
    ),
    AreaRequirement(
        key="sanitized-regression",
        title="Sanitized regression cases",
        required_files=("fixtures/real-sanitized-cases/manifest.json", "fixtures/real-sanitized-cases/cases/pq-folder-dynamic-expand-order.json", "fixtures/real-sanitized-cases/cases/dax-excel-powerpivot-compat.json", "fixtures/real-sanitized-cases/cases/cube-zero-result-debug.json", "tools/run_case_regression.py", "docs/real-case-regression.md"),
        public_doc_terms=("sanitized", "real/sanitized case regression"),
        goal_terms=("sanitized regression cases", "脱敏回归案例"),
    ),
    AreaRequirement(
        key="public-validation",
        title="Public validation workflow",
        required_files=("tools/validate-skills.py", "tools/validate_project_docs.py", "tools/validate_task_recipes.py", "tools/build_artifact_hygiene_report.py", "tools/install.mjs"),
        public_doc_terms=("python tools/validate-skills.py .", "node tools/install.mjs --check"),
        goal_terms=("Public validation", "公开校验"),
    ),
    AreaRequirement(
        key="ci-validation",
        title="GitHub Actions public validation",
        required_files=(".github/workflows/validate.yml",),
        public_doc_terms=("GitHub Actions", "public structural validation"),
        goal_terms=("CI", "GitHub Actions"),
    ),
    AreaRequirement(
        key="artifact-hygiene",
        title="Artifact hygiene and privacy boundary",
        required_files=("tools/build_artifact_hygiene_report.py", ".gitignore", "docs/distribution-checklist.md"),
        public_doc_terms=("Do Not Include", "customer workbooks"),
        goal_terms=("Artifact hygiene", "客户文件"),
    ),
    AreaRequirement(
        key="runtime-boundaries",
        title="Runtime and platform boundaries",
        required_files=("docs/compatibility.md", "tools/run_release_gate.py", "tools/run_release_gate.sh"),
        public_doc_terms=("Windows desktop Excel", "macOS", "Linux"),
        goal_terms=("Runtime boundary clarity", "运行时边界清晰"),
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


def public_doc_text(project_root: Path) -> str:
    paths = [
        "README.md",
        "README.zh-CN.md",
        "docs/project.md",
        "docs/project.en-US.md",
        "docs/project.zh-CN.md",
        "docs/current-status.md",
        "docs/install-and-sync.md",
        "docs/distribution-checklist.md",
        "docs/compatibility.md",
        "docs/task-recipes.md",
        "docs/intro.html",
        "docs/intro.zh-CN.html",
        "CONTRIBUTING.md",
        "SECURITY.md",
    ]
    return "\n".join(read_text(project_root / path) for path in paths)


def goal_doc_text(project_root: Path) -> str:
    paths = [
        "docs/maintenance-goals.en-US.md",
        "docs/maintenance-goals.zh-CN.md",
        "docs/maintenance-goals.md",
        "docs/growth-goals.en-US.md",
        "docs/growth-goals.zh-CN.md",
        "docs/growth-goals.md",
        "docs/repository-governance-goals.en-US.md",
        "docs/repository-governance-goals.zh-CN.md",
        "docs/repository-governance-goals.md",
        "docs/marketing-copy.en-US.md",
        "docs/marketing-copy.zh-CN.md",
        "docs/marketing-copy.md",
        "docs/release-notes.en-US.md",
        "docs/release-notes.zh-CN.md",
        "docs/release-notes.md",
    ]
    return "\n".join(read_text(project_root / path) for path in paths)


def evaluate_area(project_root: Path, requirement: AreaRequirement, docs_text: str, goals_text: str) -> dict[str, Any]:
    missing_files = check_files(project_root, requirement.required_files)
    missing_public_terms = check_terms(docs_text, requirement.public_doc_terms)
    missing_goal_terms = check_terms(goals_text, requirement.goal_terms)
    status = PASS if not (missing_files or missing_public_terms or missing_goal_terms) else FAIL
    return {
        "key": requirement.key,
        "title": requirement.title,
        "status": status,
        "requiredFileCount": len(requirement.required_files),
        "publicDocTermCount": len(requirement.public_doc_terms),
        "goalTermCount": len(requirement.goal_terms),
        "missingFiles": missing_files,
        "missingPublicDocTerms": missing_public_terms,
        "missingGoalTerms": missing_goal_terms,
    }


def build_report(project_root: Path) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    docs_text = public_doc_text(project_root)
    goals_text = goal_doc_text(project_root)
    version = load_manifest_version(project_root)

    areas = [evaluate_area(project_root, requirement, docs_text, goals_text) for requirement in AREA_REQUIREMENTS]
    failed = [area for area in areas if area["status"] != PASS]
    return {
        "projectRoot": str(project_root),
        "version": version,
        "status": PASS if not failed else FAIL,
        "areaCount": len(areas),
        "passedAreaCount": len(areas) - len(failed),
        "failedAreaCount": len(failed),
        "areas": areas,
        "limitations": [
            "This coverage report checks public repository files and public documentation alignment.",
            "It does not validate private workbooks or Windows Excel COM runtime behavior.",
            "A pass means the public maintenance goals are covered by shipped files and docs, not that every future optimization is complete.",
        ],
    }


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
            ("missingPublicDocTerms", "public docs"),
            ("missingGoalTerms", "maintenance goals"),
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
    parser.add_argument("--out-json", default="", help="Write machine-readable coverage report")
    parser.add_argument("--out-md", default="", help="Write Markdown coverage report")
    parser.add_argument("--require-pass", action="store_true", help="Exit with code 1 if coverage is incomplete")
    args = parser.parse_args()

    report = build_report(Path(args.project_root))
    if args.out_json:
        out_json = Path(args.out_json).expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = Path(args.out_md).expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")

    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    if args.require_pass and report.get("status") != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
