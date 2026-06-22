#!/usr/bin/env python3
"""Validate task recipe documentation against the current package shape."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


REQUIRED_SECTIONS = [
    "## Routing First",
    "## Fast Profiles",
    "## Recipe 1: Triage A Mixed Excel BI Workbook",
    "## Recipe 2: Edit And Validate VBA In An `.xlsm`",
    "## Recipe 3: Add, Update, Delete, Or Refresh Power Query",
    "## Recipe 4: Inspect Power Pivot And DAX",
    "## Recipe 5: Trace CUBE Formulas And MDX References",
    "## Recipe 6: Validate ADO/OLEDB Workbook SQL",
    "## Recipe 7: Git Bash On Windows",
    "## Recipe 8: Linux Or macOS Structural Review",
    "## Recipe 9: Publish A Clean Excel Deliverable",
    "## Recipe 10: Audit Workbook QA Before Delivery",
    "## Recipe 11: Diagnose Office Environment Readiness",
    "## Recipe 12: Build A Polished Excel Report Workbook",
    "## Recipe 13: Review Power BI Semantic Model Portability",
    "## Recipe 14: Create Sanitized Testing Fixtures",
    "## Recipe 15: Run Real/Sanitized Case Regression",
    "## Delivery Rule",
]

PATH_RE = re.compile(r"(?P<path>(?:tools|\.agents[\\/]+skills)[\\/]+[A-Za-z0-9_.\\/\-]+)")


def canonical_skills(project_root: Path) -> list[str]:
    skills_dir = project_root / ".agents" / "skills"
    return sorted(path.name for path in skills_dir.iterdir() if (path / "SKILL.md").is_file())


def normalize_doc_path(raw: str) -> str:
    value = raw.strip().strip("`'\"")
    value = value.rstrip(".,;:)`")
    return value.replace("\\", "/")


def referenced_package_paths(text: str) -> list[str]:
    paths: list[str] = []
    for match in PATH_RE.finditer(text):
        value = normalize_doc_path(match.group("path"))
        if value and value not in paths:
            paths.append(value)
    return paths


def validate(project_root: Path) -> dict[str, object]:
    recipe_path = project_root / "docs" / "task-recipes.md"
    errors: list[str] = []
    warnings: list[str] = []

    if not recipe_path.is_file():
        return {
            "status": "fail",
            "recipePath": str(recipe_path),
            "errors": [f"missing {recipe_path}"],
            "warnings": [],
        }

    text = recipe_path.read_text(encoding="utf-8")
    skills = canonical_skills(project_root)

    missing_sections = [section for section in REQUIRED_SECTIONS if section not in text]
    for section in missing_sections:
        errors.append(f"missing required section: {section}")

    missing_skills = [skill for skill in skills if skill not in text]
    for skill in missing_skills:
        errors.append(f"canonical skill not referenced in task recipes: {skill}")

    paths = referenced_package_paths(text)
    missing_paths: list[str] = []
    for path in paths:
        candidate = project_root / Path(path)
        if not candidate.exists():
            missing_paths.append(path)
    for path in missing_paths:
        errors.append(f"referenced package path does not exist: {path}")

    code_blocks = re.findall(r"```[^\n]*\n(.*?)```", text, flags=re.S)
    if len(code_blocks) < 8:
        warnings.append(f"expected at least 8 command/code blocks, found {len(code_blocks)}")

    result = {
        "status": "pass" if not errors else "fail",
        "recipePath": str(recipe_path),
        "requiredSectionCount": len(REQUIRED_SECTIONS),
        "skillCount": len(skills),
        "referencedPathCount": len(paths),
        "codeBlockCount": len(code_blocks),
        "skills": skills,
        "referencedPaths": paths,
        "errors": errors,
        "warnings": warnings,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--out-json", default="", help="Write validation report JSON")
    args = parser.parse_args()

    project_root = Path(args.project_root).expanduser().resolve()
    result = validate(project_root)

    if args.out_json:
        out_path = Path(args.out_json).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    if result["status"] == "pass":
        print(
            "Task recipe validation OK: "
            f"{result['requiredSectionCount']} sections, "
            f"{result['skillCount']} skills, "
            f"{result['referencedPathCount']} package paths"
        )
        return 0

    print("Task recipe validation failed:", file=sys.stderr)
    for error in result["errors"]:
        print(f"- {error}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
