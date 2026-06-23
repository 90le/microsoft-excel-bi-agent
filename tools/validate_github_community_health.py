#!/usr/bin/env python3
"""Validate GitHub community-health files and safe public intake templates."""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any


REQUIRED_FILES = [
    "CONTRIBUTING.md",
    "SECURITY.md",
    ".github/pull_request_template.md",
    ".github/ISSUE_TEMPLATE/config.yml",
    ".github/ISSUE_TEMPLATE/bug_report.yml",
    ".github/ISSUE_TEMPLATE/feature_request.yml",
    ".github/ISSUE_TEMPLATE/docs_install.yml",
    "docs/repository-governance-goals.md",
    "docs/repository-governance-goals.en-US.md",
    "docs/repository-governance-goals.zh-CN.md",
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

SAFETY_TERMS = [
    "customer workbooks",
    "screenshots",
    "PDFs",
    "credentials",
    "local private paths",
    "generated QA reports",
    "unsanitized runtime evidence",
    "sanitized",
]

ZH_SAFETY_TERMS = [
    "客户工作簿",
    "截图",
    "PDF",
    "凭证",
    "本机路径",
    "生成 QA 报告",
    "未脱敏运行证据",
    "脱敏",
]


def rel(path: str) -> Path:
    return Path(path.replace("/", os.sep))


def read_text(project_root: Path, path: str) -> str:
    full_path = project_root / rel(path)
    if not full_path.is_file():
        return ""
    return full_path.read_text(encoding="utf-8")


def missing_terms(text: str, terms: list[str]) -> list[str]:
    return [term for term in terms if term not in text]


def validate(project_root: Path) -> dict[str, Any]:
    project_root = project_root.expanduser().resolve()
    errors: list[str] = []

    for path in REQUIRED_FILES:
        if not (project_root / rel(path)).is_file():
            errors.append(f"missing GitHub community-health file: {path}")

    texts = {path: read_text(project_root, path) for path in REQUIRED_FILES}
    issue_text = "\n".join(
        texts.get(path, "")
        for path in [
            ".github/ISSUE_TEMPLATE/bug_report.yml",
            ".github/ISSUE_TEMPLATE/feature_request.yml",
            ".github/ISSUE_TEMPLATE/docs_install.yml",
        ]
    )
    for term in SAFETY_TERMS:
        if term not in issue_text:
            errors.append(f"issue templates do not contain safety term: {term}")

    config = texts.get(".github/ISSUE_TEMPLATE/config.yml", "")
    if "blank_issues_enabled: false" not in config:
        errors.append("issue template config must set blank_issues_enabled: false")
    if "security/policy" not in config:
        errors.append("issue template config must link the security policy")

    pr_template = texts.get(".github/pull_request_template.md", "")
    for command in PUBLIC_CHECK_COMMANDS:
        if command not in pr_template:
            errors.append(f"PR template does not document required check: {command}")
    for term in [".agents/skills/", "generated mirrors", "npm/npx", "Windows desktop Excel"]:
        if term not in pr_template:
            errors.append(f"PR template missing governance term: {term}")

    contributing = texts.get("CONTRIBUTING.md", "")
    for command in PUBLIC_CHECK_COMMANDS:
        if command not in contributing:
            errors.append(f"CONTRIBUTING.md does not document required check: {command}")
    for term in ["Qiu Binbin", "丘彬彬", "binstudy", "90le.cn", ".agents/skills/", "generated skill mirrors"]:
        if term not in contributing:
            errors.append(f"CONTRIBUTING.md missing term: {term}")

    security = texts.get("SECURITY.md", "")
    for term in ["Do Not Put These In Public Issues", "binstudy", "credentials", "customer workbooks", "Windows desktop Excel"]:
        if term not in security:
            errors.append(f"SECURITY.md missing term: {term}")

    governance_en = texts.get("docs/repository-governance-goals.en-US.md", "")
    governance_zh = texts.get("docs/repository-governance-goals.zh-CN.md", "")
    for heading in ["## Objective", "## Constraints", "## Boundaries", "## Can Do", "## Cannot Do", "## Detailed Goals", "## High-Value Backlog", "## Required Public Checks"]:
        if heading not in governance_en:
            errors.append(f"English governance goals missing heading: {heading}")
    for heading in ["## 目标", "## 约束", "## 边界", "## 可以做", "## 不能做", "## 详细 Goal", "## 高价值 Backlog", "## 必跑公开校验"]:
        if heading not in governance_zh:
            errors.append(f"Chinese governance goals missing heading: {heading}")
    for term in SAFETY_TERMS:
        if term not in governance_en:
            errors.append(f"English governance goals missing safety term: {term}")
    for term in ZH_SAFETY_TERMS:
        if term not in governance_zh:
            errors.append(f"Chinese governance goals missing safety term: {term}")
    for command in PUBLIC_CHECK_COMMANDS:
        if command not in governance_en or command not in governance_zh:
            errors.append(f"governance goals do not document required check: {command}")

    result: dict[str, Any] = {
        "status": "pass" if not errors else "fail",
        "projectRoot": str(project_root),
        "requiredFiles": REQUIRED_FILES,
        "requiredFileCount": len(REQUIRED_FILES),
        "publicCheckCommands": PUBLIC_CHECK_COMMANDS,
        "errors": errors,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Project root")
    parser.add_argument("--out-json", default="", help="Write validation report JSON")
    args = parser.parse_args()

    result = validate(Path(args.project_root))
    if args.out_json:
        out_path = Path(args.out_json).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    if result["status"] == "pass":
        print(
            "GitHub community health validation OK: "
            f"files={result['requiredFileCount']}, "
            f"publicChecks={len(result['publicCheckCommands'])}"
        )
        return 0

    print("GitHub community health validation failed:", file=sys.stderr)
    for error in result["errors"]:
        print(f"- {error}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
