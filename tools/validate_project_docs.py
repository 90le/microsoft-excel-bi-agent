#!/usr/bin/env python3
"""Validate public project documentation for Microsoft Excel BI Agent.

The public repository intentionally excludes maintainer-only release ledgers,
machine-specific runtime evidence, customer workbooks, and generated QA
reports. This validator checks the documentation contract that is actually
shipped to users and contributors.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path
from typing import Any


PUBLIC_DOCS = [
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
    "docs/maintenance-goals.md",
    "docs/maintenance-goals.en-US.md",
    "docs/maintenance-goals.zh-CN.md",
    "docs/growth-goals.md",
    "docs/growth-goals.en-US.md",
    "docs/growth-goals.zh-CN.md",
    "docs/repository-governance-goals.md",
    "docs/repository-governance-goals.en-US.md",
    "docs/repository-governance-goals.zh-CN.md",
    "docs/marketing-copy.md",
    "docs/marketing-copy.en-US.md",
    "docs/marketing-copy.zh-CN.md",
    "docs/release-notes.md",
    "docs/release-notes.en-US.md",
    "docs/release-notes.zh-CN.md",
    "docs/intro.html",
    "docs/intro.zh-CN.html",
    "docs/index.html",
    "CONTRIBUTING.md",
    "SECURITY.md",
]

REAL_INSTALL_COMMANDS = [
    "codex plugin marketplace add 90le/microsoft-excel-bi-agent",
    "codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent",
    "node tools/install.mjs",
]

PUBLIC_CHECK_COMMANDS = [
    "python tools/validate-skills.py .",
    "python tools/validate_project_docs.py --project-root .",
    "python tools/validate_github_community_health.py --project-root .",
    "python tools/validate_task_recipes.py --project-root .",
    "python tools/validate_official_docs_index.py --project-root .",
    "python tools/build_artifact_hygiene_report.py --project-root . --require-pass",
    "python tools/build_goal_coverage_report.py --project-root . --require-pass",
    "node tools/install.mjs --check",
]

EN_GOAL_HEADINGS = [
    "## Objective",
    "## Constraints",
    "## Boundaries",
    "## Can Do",
    "## Cannot Do",
    "## Detailed Goals",
    "## Risk Register",
    "## Optimization Backlog",
    "## Required Public Checks",
    "## Must-Worthy Optimization Rule",
]

ZH_GOAL_HEADINGS = [
    "## 目标",
    "## 约束",
    "## 边界",
    "## 可以做",
    "## 不能做",
    "## 详细 Goal",
    "## 风险清单",
    "## 优化 Backlog",
    "## 必跑公开校验",
    "## 必须值得的优化规则",
]

EN_GROWTH_HEADINGS = [
    "## Objective",
    "## Constraints",
    "## Boundaries",
    "## Can Do",
    "## Cannot Do",
    "## Detailed Goals",
    "## High-Value Optimization Backlog",
    "## Required Public Checks",
]

ZH_GROWTH_HEADINGS = [
    "## 目标",
    "## 约束",
    "## 边界",
    "## 可以做",
    "## 不能做",
    "## 详细 Goal",
    "## 高价值优化 Backlog",
    "## 必跑公开校验",
]

SIGNATURE_NAME_TERMS = ("Qiu Binbin", "丘彬彬")
SIGNATURE_CONTACT_TERMS = ("binstudy", "90le.cn")

MOJIBAKE_MARKERS = (
    "涓",
    "绔",
    "椤圭",
    "瀹夎",
    "鏍￠",
    "闁",
    "鐎",
    "閻",
)

PRIVATE_LEDGER_LINKS = (
    "docs/master-goal.md",
    "docs/goal-tracking.md",
    "docs/iteration-protocol.md",
    "docs/progress.md",
    "docs/validation.md",
    "docs/completion-evidence.md",
)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def rel(path: str) -> Path:
    return Path(path.replace("/", os.sep))


def contains_cjk(text: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def has_mojibake(text: str) -> list[str]:
    return [marker for marker in MOJIBAKE_MARKERS if marker in text]


def check_required_links(label: str, text: str, links: list[str], errors: list[str]) -> None:
    for link in links:
        if link not in text:
            errors.append(f"{label} does not link {link}")


def validate(project_root: Path) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    errors: list[str] = []
    warnings: list[str] = []

    missing_docs = [path for path in PUBLIC_DOCS if not (project_root / rel(path)).is_file()]
    for path in missing_docs:
        errors.append(f"missing public document: {path}")

    texts: dict[str, str] = {}
    for path in PUBLIC_DOCS:
        full_path = project_root / rel(path)
        if full_path.is_file():
            texts[path] = read_text(full_path)

    for path, text in texts.items():
        markers = has_mojibake(text)
        if markers:
            errors.append(f"{path} contains possible Chinese mojibake markers: {', '.join(markers)}")

    readme_en = texts.get("README.md", "")
    readme_zh = texts.get("README.zh-CN.md", "")
    install_doc = texts.get("docs/install-and-sync.md", "")
    distribution_doc = texts.get("docs/distribution-checklist.md", "")
    project_index = texts.get("docs/project.md", "")
    project_en = texts.get("docs/project.en-US.md", "")
    project_zh = texts.get("docs/project.zh-CN.md", "")
    current_status = texts.get("docs/current-status.md", "")
    goals_en = texts.get("docs/maintenance-goals.en-US.md", "")
    goals_zh = texts.get("docs/maintenance-goals.zh-CN.md", "")
    goals_index = texts.get("docs/maintenance-goals.md", "")
    growth_en = texts.get("docs/growth-goals.en-US.md", "")
    growth_zh = texts.get("docs/growth-goals.zh-CN.md", "")
    growth_index = texts.get("docs/growth-goals.md", "")
    governance_en = texts.get("docs/repository-governance-goals.en-US.md", "")
    governance_zh = texts.get("docs/repository-governance-goals.zh-CN.md", "")
    governance_index = texts.get("docs/repository-governance-goals.md", "")
    marketing_en = texts.get("docs/marketing-copy.en-US.md", "")
    marketing_zh = texts.get("docs/marketing-copy.zh-CN.md", "")
    marketing_index = texts.get("docs/marketing-copy.md", "")
    release_notes_en = texts.get("docs/release-notes.en-US.md", "")
    release_notes_zh = texts.get("docs/release-notes.zh-CN.md", "")
    release_notes_index = texts.get("docs/release-notes.md", "")
    intro_en = texts.get("docs/intro.html", "")
    intro_zh = texts.get("docs/intro.zh-CN.html", "")
    site_index = texts.get("docs/index.html", "")
    contributing = texts.get("CONTRIBUTING.md", "")
    security = texts.get("SECURITY.md", "")

    english_readme_probe = readme_en.replace("[中文]", "").replace("丘彬彬", "")
    if contains_cjk(english_readme_probe):
        warnings.append("README.md contains CJK characters outside the language switcher")
    if not contains_cjk(readme_zh):
        errors.append("README.zh-CN.md does not appear to contain Chinese text")

    check_required_links(
        "README.md",
        readme_en,
        [
            "README.zh-CN.md",
            "docs/project.en-US.md",
            "docs/maintenance-goals.en-US.md",
            "docs/growth-goals.en-US.md",
            "docs/repository-governance-goals.en-US.md",
            "docs/marketing-copy.en-US.md",
            "docs/release-notes.en-US.md",
            "docs/install-and-sync.md",
            "docs/intro.html",
            "CONTRIBUTING.md",
            "SECURITY.md",
        ],
        errors,
    )
    check_required_links(
        "README.zh-CN.md",
        readme_zh,
        [
            "README.md",
            "docs/project.zh-CN.md",
            "docs/maintenance-goals.zh-CN.md",
            "docs/growth-goals.zh-CN.md",
            "docs/repository-governance-goals.zh-CN.md",
            "docs/marketing-copy.zh-CN.md",
            "docs/release-notes.zh-CN.md",
            "docs/install-and-sync.md",
            "docs/intro.zh-CN.html",
            "CONTRIBUTING.md",
            "SECURITY.md",
        ],
        errors,
    )

    for command in REAL_INSTALL_COMMANDS:
        if command not in readme_en or command not in readme_zh or command not in install_doc:
            errors.append(f"real install command is not present in README EN, README ZH, and install guide: {command}")

    fake_current_commands = [
        "npx microsoft-excel-bi-agent",
        "npm install microsoft-excel-bi-agent",
    ]
    for path in ["README.md", "README.zh-CN.md", "docs/intro.html", "docs/intro.zh-CN.html"]:
        text = texts.get(path, "")
        for command in fake_current_commands:
            if command in text:
                errors.append(f"{path} advertises unsupported current command: {command}")

    check_required_links(
        "docs/project.md",
        project_index,
        [
            "project.en-US.md",
            "project.zh-CN.md",
            "maintenance-goals.en-US.md",
            "maintenance-goals.zh-CN.md",
            "growth-goals.en-US.md",
            "growth-goals.zh-CN.md",
            "repository-governance-goals.en-US.md",
            "repository-governance-goals.zh-CN.md",
            "marketing-copy.en-US.md",
            "marketing-copy.zh-CN.md",
            "release-notes.en-US.md",
            "release-notes.zh-CN.md",
            "intro.html",
            "intro.zh-CN.html",
        ],
        errors,
    )
    check_required_links("docs/project.en-US.md", project_en, ["maintenance-goals.en-US.md", "install-and-sync.md"], errors)
    check_required_links("docs/project.zh-CN.md", project_zh, ["maintenance-goals.zh-CN.md", "install-and-sync.md"], errors)
    check_required_links("docs/maintenance-goals.md", goals_index, ["maintenance-goals.en-US.md", "maintenance-goals.zh-CN.md"], errors)
    check_required_links("docs/growth-goals.md", growth_index, ["growth-goals.en-US.md", "growth-goals.zh-CN.md"], errors)
    check_required_links("docs/repository-governance-goals.md", governance_index, ["repository-governance-goals.en-US.md", "repository-governance-goals.zh-CN.md", "CONTRIBUTING.md", "SECURITY.md"], errors)
    check_required_links("docs/marketing-copy.md", marketing_index, ["marketing-copy.en-US.md", "marketing-copy.zh-CN.md"], errors)
    check_required_links("docs/release-notes.md", release_notes_index, ["release-notes.en-US.md", "release-notes.zh-CN.md"], errors)
    check_required_links(
        "docs/current-status.md",
        current_status,
        ["docs/maintenance-goals.en-US.md", "docs/maintenance-goals.zh-CN.md", "docs/growth-goals.en-US.md", "docs/growth-goals.zh-CN.md", "docs/repository-governance-goals.en-US.md", "docs/repository-governance-goals.zh-CN.md", "docs/marketing-copy.en-US.md", "docs/marketing-copy.zh-CN.md", "CONTRIBUTING.md", "SECURITY.md"],
        errors,
    )

    for heading in EN_GOAL_HEADINGS:
        if heading not in goals_en:
            errors.append(f"docs/maintenance-goals.en-US.md missing heading: {heading}")
    for heading in ZH_GOAL_HEADINGS:
        if heading not in goals_zh:
            errors.append(f"docs/maintenance-goals.zh-CN.md missing heading: {heading}")
    for heading in EN_GROWTH_HEADINGS:
        if heading not in growth_en:
            errors.append(f"docs/growth-goals.en-US.md missing heading: {heading}")
    for heading in ZH_GROWTH_HEADINGS:
        if heading not in growth_zh:
            errors.append(f"docs/growth-goals.zh-CN.md missing heading: {heading}")

    for path, text in [
        ("README.md", readme_en),
        ("README.zh-CN.md", readme_zh),
        ("docs/project.en-US.md", project_en),
        ("docs/project.zh-CN.md", project_zh),
        ("docs/growth-goals.en-US.md", growth_en),
        ("docs/growth-goals.zh-CN.md", growth_zh),
        ("docs/repository-governance-goals.en-US.md", governance_en),
        ("docs/repository-governance-goals.zh-CN.md", governance_zh),
        ("docs/marketing-copy.en-US.md", marketing_en),
        ("docs/marketing-copy.zh-CN.md", marketing_zh),
        ("docs/intro.html", intro_en),
        ("docs/intro.zh-CN.html", intro_zh),
        ("CONTRIBUTING.md", contributing),
        ("SECURITY.md", security),
    ]:
        if not any(term in text for term in SIGNATURE_NAME_TERMS):
            errors.append(f"{path} does not contain a maintainer name")
        for term in SIGNATURE_CONTACT_TERMS:
            if term not in text:
                errors.append(f"{path} does not contain maintainer signature term: {term}")

    for command in PUBLIC_CHECK_COMMANDS:
        if command not in goals_en or command not in goals_zh:
            errors.append(f"maintenance goals do not document public check: {command}")
        if command not in release_notes_en or command not in release_notes_zh:
            errors.append(f"release notes do not document public check: {command}")

    for path, text in [("docs/release-notes.en-US.md", release_notes_en), ("docs/release-notes.zh-CN.md", release_notes_zh)]:
        if "v0.1.5" not in text or "0.1.5+codex.20260623175347" not in text:
            errors.append(f"{path} does not document the v0.1.5 release and plugin version")

    for command in PUBLIC_CHECK_COMMANDS:
        if command not in distribution_doc.replace("\\", "/"):
            errors.append(f"distribution checklist does not document public check: {command}")
        if command not in install_doc.replace("\\", "/"):
            errors.append(f"install guide does not document public check: {command}")

    for private_link in PRIVATE_LEDGER_LINKS:
        for path in [
            "README.md",
            "README.zh-CN.md",
            "docs/project.md",
            "docs/project.en-US.md",
            "docs/project.zh-CN.md",
            "docs/current-status.md",
            "docs/distribution-checklist.md",
            "docs/maintenance-goals.en-US.md",
            "docs/maintenance-goals.zh-CN.md",
        ]:
            if private_link in texts.get(path, ""):
                errors.append(f"{path} links maintainer-only ignored document: {private_link}")

    for path, text in [("docs/intro.html", intro_en), ("docs/intro.zh-CN.html", intro_zh)]:
        if 'name="viewport"' not in text:
            errors.append(f"{path} is missing responsive viewport meta")
        if "<html" not in text or "</html>" not in text:
            errors.append(f"{path} does not look like a complete HTML page")
        if "v0.1.5" not in text:
            errors.append(f"{path} does not expose the latest release")
        if "release-notes" not in text:
            errors.append(f"{path} does not link release notes")
        if "growth-goals" not in text:
            errors.append(f"{path} does not link growth goals")
        if "repository-governance-goals" not in text:
            errors.append(f"{path} does not link repository governance goals")
        if "marketing-copy" not in text:
            errors.append(f"{path} does not link marketing copy")
        if "CONTRIBUTING.md" not in text or "SECURITY.md" not in text:
            errors.append(f"{path} does not link contributing and security files")
        if "https://90le.github.io/microsoft-excel-bi-agent/assets/social-preview" not in text:
            errors.append(f"{path} does not use an absolute social preview image URL")
        if 'rel="canonical"' not in text or "og:url" not in text:
            errors.append(f"{path} does not expose canonical and og:url metadata")
    if 'lang="en"' not in intro_en:
        errors.append("docs/intro.html does not declare lang=\"en\"")
    if 'lang="zh-CN"' not in intro_zh:
        errors.append("docs/intro.zh-CN.html does not declare lang=\"zh-CN\"")
    if "navigator.language" not in site_index or "intro.zh-CN.html" not in site_index or "intro.html" not in site_index:
        errors.append("docs/index.html does not contain the browser-language redirect contract")

    result: dict[str, Any] = {
        "status": "pass" if not errors else "fail",
        "projectRoot": str(project_root),
        "publicDocs": PUBLIC_DOCS,
        "publicDocCount": len(PUBLIC_DOCS),
        "realInstallCommands": REAL_INSTALL_COMMANDS,
        "publicCheckCommands": PUBLIC_CHECK_COMMANDS,
        "errors": errors,
        "warnings": warnings,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--out-json", default="", help="Write validation report JSON")
    args = parser.parse_args()

    result = validate(Path(args.project_root))

    if args.out_json:
        out_path = Path(args.out_json).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    if result["status"] == "pass":
        print(
            "Project docs validation OK: "
            f"publicDocs={result['publicDocCount']}, "
            f"installCommands={len(result['realInstallCommands'])}, "
            f"publicChecks={len(result['publicCheckCommands'])}"
        )
        for warning in result["warnings"]:
            print(f"warning: {warning}", file=sys.stderr)
        return 0

    print("Project docs validation failed:", file=sys.stderr)
    for error in result["errors"]:
        print(f"- {error}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
