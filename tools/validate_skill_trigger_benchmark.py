#!/usr/bin/env python3
"""Validate Excel BI skill trigger metadata and the checked-in trigger corpus."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


REPORT_KIND = "excel-skill-trigger-benchmark-report"
REPORT_SCHEMA_VERSION = "1.0"
CORPUS_KIND = "excel-skill-trigger-benchmark"
CORPUS_SCHEMA_VERSION = "1.0"

CANONICAL_SKILL_IDS = (
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
)
CANONICAL_SKILL_SET = frozenset(CANONICAL_SKILL_IDS)

MAX_DESCRIPTION_CHARS = 300
MAX_AGGREGATE_DESCRIPTION_CHARS = 2300
DEFAULT_PROMPT_COUNT = 3
MAX_DEFAULT_PROMPT_CHARS = 110

ROOT_FIELDS = frozenset({"kind", "schemaVersion", "cases"})
CASE_FIELDS = frozenset(
    {"id", "kind", "targetSkill", "expectedSkill", "text", "successChecklist"}
)
CASE_KINDS = frozenset({"positive", "confusable-negative"})

WINDOWS_ABSOLUTE_PATH = re.compile(r"(?i)(?<![a-z0-9])(?:[a-z]:[\\/]|\\\\)")
POSIX_ABSOLUTE_PATH = re.compile(r"(?<![a-z0-9:/])/(?!\s)", re.IGNORECASE)
FILE_URI = re.compile(r"file://", re.IGNORECASE)
CREDENTIAL_TEXT = re.compile(
    r"(?i)\b(?:password|passwd|pwd|secret|api[-_ ]?key|access[-_ ]?token|client[-_ ]?secret)\s*[:=]"
)
CUSTOMER_ARTIFACT = re.compile(
    r"(?i)\b[\w-]+\.(?:xlsx|xlsm|xlsb|xls|csv|pbix|pbit|accdb|mdb)\b"
)


def _empty_summary() -> dict[str, int]:
    return {
        "skillCount": 0,
        "caseCount": 0,
        "positiveCount": 0,
        "confusableNegativeCount": 0,
    }


def _new_report() -> dict[str, Any]:
    return {
        "kind": REPORT_KIND,
        "schemaVersion": REPORT_SCHEMA_VERSION,
        "status": "pass",
        "summary": _empty_summary(),
        "skills": {},
        "errors": [],
    }


def _finish(report: dict[str, Any]) -> dict[str, Any]:
    report["status"] = "fail" if report["errors"] else "pass"
    return report


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def _parse_scalar(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] == '"':
        try:
            parsed = json.loads(value)
            return parsed if isinstance(parsed, str) else value
        except json.JSONDecodeError:
            return value[1:-1]
    if len(value) >= 2 and value[0] == value[-1] == "'":
        return value[1:-1].replace("''", "'")
    return value


def _read_frontmatter(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8-sig")
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        raise ValueError("missing opening YAML frontmatter delimiter")
    try:
        closing = next(index for index, line in enumerate(lines[1:], 1) if line.strip() == "---")
    except StopIteration as exc:
        raise ValueError("missing closing YAML frontmatter delimiter") from exc

    result: dict[str, str] = {}
    index = 1
    while index < closing:
        line = lines[index]
        if not line.strip() or line.lstrip().startswith("#"):
            index += 1
            continue
        if ":" not in line:
            raise ValueError(f"unsupported YAML frontmatter line {index + 1}")
        key, raw_value = line.split(":", 1)
        key = key.strip()
        raw_value = raw_value.strip()
        if raw_value in {">", "|-", "|", ">-"}:
            folded: list[str] = []
            index += 1
            while index < closing and (not lines[index].strip() or lines[index][0].isspace()):
                folded.append(lines[index].strip())
                index += 1
            separator = "\n" if raw_value.startswith("|") else " "
            result[key] = separator.join(part for part in folded if part)
            continue
        result[key] = _parse_scalar(raw_value)
        index += 1
    return result


def _discover_skill_ids(project_root: Path, errors: list[str]) -> tuple[str, ...]:
    skills_root = project_root / ".agents" / "skills"
    if not skills_root.is_dir():
        errors.append("missing canonical skill directory: .agents/skills")
        return ()
    discovered = tuple(sorted(path.name for path in skills_root.iterdir() if path.is_dir()))
    missing = sorted(CANONICAL_SKILL_SET.difference(discovered))
    unknown = sorted(set(discovered).difference(CANONICAL_SKILL_SET))
    if len(discovered) != len(CANONICAL_SKILL_IDS):
        errors.append(f"expected exactly 12 canonical skill directories, found {len(discovered)}")
    if missing:
        errors.append(f"missing canonical skill IDs: {', '.join(missing)}")
    if unknown:
        errors.append(f"unknown canonical skill IDs: {', '.join(unknown)}")
    return discovered


def _unsafe_reason(value: str) -> str | None:
    if WINDOWS_ABSOLUTE_PATH.search(value) or POSIX_ABSOLUTE_PATH.search(value) or FILE_URI.search(value):
        return "absolute path"
    if CREDENTIAL_TEXT.search(value):
        return "credential-like text"
    if CUSTOMER_ARTIFACT.search(value):
        return "customer artifact filename"
    return None


def validate_trigger_cases(project_root: Path | str, cases_json: Path | str) -> dict[str, Any]:
    """Validate the trigger corpus independently of not-yet-migrated metadata."""

    project_root = Path(project_root).resolve()
    cases_json = Path(cases_json).resolve()
    report = _new_report()
    errors: list[str] = report["errors"]
    discovered = _discover_skill_ids(project_root, errors)
    report["summary"]["skillCount"] = len(discovered)

    try:
        document = _read_json(cases_json)
    except (OSError, json.JSONDecodeError) as exc:
        errors.append(f"cannot read trigger corpus: {exc}")
        return _finish(report)

    if not isinstance(document, dict):
        errors.append("trigger corpus root must be an object")
        return _finish(report)

    root_fields = set(document)
    missing_root = sorted(ROOT_FIELDS.difference(root_fields))
    unknown_root = sorted(root_fields.difference(ROOT_FIELDS))
    if missing_root:
        errors.append(f"missing corpus fields: {', '.join(missing_root)}")
    if unknown_root:
        errors.append(f"unknown corpus fields: {', '.join(unknown_root)}")
    if document.get("kind") != CORPUS_KIND:
        errors.append(f"corpus kind must be {CORPUS_KIND!r}")
    if document.get("schemaVersion") != CORPUS_SCHEMA_VERSION:
        errors.append(f"corpus schemaVersion must be {CORPUS_SCHEMA_VERSION!r}")

    cases = document.get("cases")
    if not isinstance(cases, list):
        errors.append("corpus cases must be an array")
        return _finish(report)

    report["summary"]["caseCount"] = len(cases)
    ids: set[str] = set()
    skill_counts = {
        skill_id: {"positiveCount": 0, "confusableNegativeCount": 0}
        for skill_id in CANONICAL_SKILL_IDS
    }

    for index, case in enumerate(cases):
        label = f"case[{index}]"
        if not isinstance(case, dict):
            errors.append(f"{label} must be an object")
            continue
        fields = set(case)
        missing = sorted(CASE_FIELDS.difference(fields))
        unknown = sorted(fields.difference(CASE_FIELDS))
        if missing:
            errors.append(f"{label} missing fields: {', '.join(missing)}")
        if unknown:
            errors.append(f"{label} unknown fields: {', '.join(unknown)}")
        if missing or unknown:
            continue

        case_id = case["id"]
        if not isinstance(case_id, str) or not case_id.strip():
            errors.append(f"{label} id must be a non-empty string")
        elif case_id in ids:
            errors.append(f"duplicate case id: {case_id}")
        else:
            ids.add(case_id)

        case_kind = case["kind"]
        target = case["targetSkill"]
        expected = case["expectedSkill"]
        text = case["text"]
        checklist = case["successChecklist"]

        kind_is_string = isinstance(case_kind, str)
        target_is_string = isinstance(target, str)
        expected_is_string = isinstance(expected, str)
        if not kind_is_string:
            errors.append(f"{label} kind must be a string")
        elif case_kind not in CASE_KINDS:
            errors.append(f"{label} kind must be positive or confusable-negative")
        if not target_is_string:
            errors.append(f"{label} targetSkill must be a string")
        elif target not in CANONICAL_SKILL_SET:
            errors.append(f"{label} targetSkill is not canonical: {target!r}")
        if not expected_is_string:
            errors.append(f"{label} expectedSkill must be a string")
        elif expected not in CANONICAL_SKILL_SET:
            errors.append(f"{label} expectedSkill is not canonical: {expected!r}")
        if not isinstance(text, str) or not text.strip():
            errors.append(f"{label} text must be a non-empty string")
        if not isinstance(checklist, list) or not checklist or not all(
            isinstance(item, str) and item.strip() for item in checklist
        ):
            errors.append(f"{label} successChecklist must be a non-empty string array")

        if isinstance(text, str):
            reason = _unsafe_reason(text)
            if reason:
                errors.append(f"{label} text contains {reason}")
        if isinstance(checklist, list):
            for checklist_index, item in enumerate(checklist):
                if isinstance(item, str):
                    reason = _unsafe_reason(item)
                    if reason:
                        errors.append(
                            f"{label} successChecklist[{checklist_index}] contains {reason}"
                        )

        if (
            target_is_string
            and target in CANONICAL_SKILL_SET
            and kind_is_string
            and case_kind in CASE_KINDS
        ):
            if case_kind == "positive":
                report["summary"]["positiveCount"] += 1
                skill_counts[target]["positiveCount"] += 1
                if expected != target:
                    errors.append(f"{label} positive expectedSkill must equal targetSkill")
            else:
                report["summary"]["confusableNegativeCount"] += 1
                skill_counts[target]["confusableNegativeCount"] += 1
                if expected == target:
                    errors.append(
                        f"{label} confusable-negative expectedSkill must differ from targetSkill"
                    )

    if len(cases) != 36:
        errors.append(f"expected exactly 36 trigger cases, found {len(cases)}")
    for skill_id, counts in skill_counts.items():
        if counts["positiveCount"] != 2:
            errors.append(
                f"{skill_id} requires exactly 2 positive cases; found {counts['positiveCount']}"
            )
        if counts["confusableNegativeCount"] != 1:
            errors.append(
                f"{skill_id} requires exactly 1 confusable-negative case; "
                f"found {counts['confusableNegativeCount']}"
            )
    report["skills"] = skill_counts
    return _finish(report)


def _validate_metadata(project_root: Path, report: dict[str, Any]) -> None:
    errors: list[str] = report["errors"]
    descriptions: dict[str, str] = {}

    for skill_id in CANONICAL_SKILL_IDS:
        skill_path = project_root / ".agents" / "skills" / skill_id / "SKILL.md"
        try:
            frontmatter = _read_frontmatter(skill_path)
        except (OSError, ValueError) as exc:
            errors.append(f"{skill_id}: cannot read SKILL.md frontmatter: {exc}")
            continue
        name = frontmatter.get("name")
        description = frontmatter.get("description")
        if name != skill_id:
            errors.append(f"{skill_id}: frontmatter name must exactly match the skill ID")
        if not isinstance(description, str) or not description:
            errors.append(f"{skill_id}: description must be a non-empty string")
            continue
        descriptions[skill_id] = description
        if not description.startswith("Use when "):
            errors.append(f"{skill_id}: description must start with 'Use when '")
        if len(description) > MAX_DESCRIPTION_CHARS:
            errors.append(
                f"{skill_id}: description has {len(description)} characters; "
                f"maximum is {MAX_DESCRIPTION_CHARS}"
            )
        report["skills"].setdefault(skill_id, {})["descriptionChars"] = len(description)

    aggregate = sum(len(description) for description in descriptions.values())
    if aggregate > MAX_AGGREGATE_DESCRIPTION_CHARS:
        errors.append(
            f"aggregate description characters are {aggregate}; "
            f"maximum is {MAX_AGGREGATE_DESCRIPTION_CHARS}"
        )

    manifest_path = project_root / ".codex-plugin" / "plugin.json"
    try:
        manifest = _read_json(manifest_path)
    except (OSError, json.JSONDecodeError) as exc:
        errors.append(f"cannot read plugin manifest: {exc}")
        return
    if not isinstance(manifest, dict):
        errors.append("plugin manifest must be an object")
        return
    interface = manifest.get("interface")
    if not isinstance(interface, dict):
        errors.append("plugin interface must be an object")
        return
    prompts = interface.get("defaultPrompt")
    if not isinstance(prompts, list):
        errors.append("plugin interface.defaultPrompt must be an array")
        return
    if len(prompts) != DEFAULT_PROMPT_COUNT:
        errors.append(f"expected exactly 3 default prompts, found {len(prompts)}")
    for index, prompt in enumerate(prompts):
        if not isinstance(prompt, str) or not prompt.strip():
            errors.append(f"defaultPrompt[{index}] must be a non-empty string")
        elif len(prompt) > MAX_DEFAULT_PROMPT_CHARS:
            errors.append(
                f"defaultPrompt[{index}] has {len(prompt)} characters; "
                f"maximum is {MAX_DEFAULT_PROMPT_CHARS}"
            )


def validate_benchmark(project_root: Path | str, cases_json: Path | str) -> dict[str, Any]:
    project_root = Path(project_root).resolve()
    report = validate_trigger_cases(project_root, cases_json)
    _validate_metadata(project_root, report)
    return _finish(report)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, required=True)
    parser.add_argument("--cases-json", type=Path, required=True)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--require-pass", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    report = validate_benchmark(args.project_root, args.cases_json)
    payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    sys.stdout.write(payload)
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(payload, encoding="utf-8")
    return 1 if args.require_pass and report["status"] == "fail" else 0


if __name__ == "__main__":
    raise SystemExit(main())
