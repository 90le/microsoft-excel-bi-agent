#!/usr/bin/env python3
"""Static DAX compatibility lint for Excel Power Pivot-oriented work.

This script catches version-sensitive or host-specific DAX patterns before a
measure is handed to Excel Power Pivot. It does not parse a live model, validate
relationships, evaluate DAX, or replace Excel/Power Pivot runtime validation.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SOURCE_EXTENSIONS = {".dax", ".txt", ".md"}
JSON_FORMULA_KEYS = {"formula", "expression", "daxFormula"}
FUNCTION_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\(", re.IGNORECASE)

EXCEL_ERROR_FUNCTIONS = {
    "REMOVEFILTERS": "Use ALL(table_or_column) or ALLEXCEPT(...) for Excel Power Pivot compatibility unless the target host is confirmed to support REMOVEFILTERS.",
}

EXCEL_WARNING_FUNCTIONS = {
    "SELECTEDVALUE": "Verify Excel Power Pivot version support; use IF(HASONEVALUE(col), VALUES(col), alternate) when support is uncertain.",
}


@dataclass
class DaxExpression:
    label: str
    text: str


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def strip_comments_and_strings(text: str) -> str:
    result: list[str] = []
    index = 0
    in_string = False
    in_line_comment = False
    in_block_comment = False

    while index < len(text):
        char = text[index]
        next_char = text[index + 1] if index + 1 < len(text) else ""

        if in_line_comment:
            if char in "\r\n":
                in_line_comment = False
                result.append(char)
            else:
                result.append(" ")
            index += 1
            continue

        if in_block_comment:
            if char == "*" and next_char == "/":
                result.extend("  ")
                in_block_comment = False
                index += 2
            else:
                result.append("\n" if char in "\r\n" else " ")
                index += 1
            continue

        if in_string:
            if char == '"':
                if next_char == '"':
                    result.extend("  ")
                    index += 2
                else:
                    in_string = False
                    result.append(" ")
                    index += 1
            else:
                result.append("\n" if char in "\r\n" else " ")
                index += 1
            continue

        if char == "/" and next_char == "/":
            in_line_comment = True
            result.extend("  ")
            index += 2
            continue
        if char == "-" and next_char == "-":
            in_line_comment = True
            result.extend("  ")
            index += 2
            continue
        if char == "/" and next_char == "*":
            in_block_comment = True
            result.extend("  ")
            index += 2
            continue
        if char == '"':
            in_string = True
            result.append(" ")
            index += 1
            continue

        result.append(char)
        index += 1

    return "".join(result)


def find_unbalanced_parentheses(cleaned: str) -> str:
    depth = 0
    for index, char in enumerate(cleaned, start=1):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth < 0:
                return f"extra closing parenthesis near character {index}"
    if depth > 0:
        return f"{depth} unclosed opening parenthesis"
    return ""


def iter_source_files(source: Path) -> list[Path]:
    if source.is_file():
        return [source]
    return sorted(
        path
        for path in source.rglob("*")
        if path.is_file() and (path.suffix.lower() in SOURCE_EXTENSIONS or path.suffix.lower() == ".json")
    )


def extract_json_expressions(value: Any, label: str) -> list[DaxExpression]:
    expressions: list[DaxExpression] = []
    if isinstance(value, dict):
        name = str(value.get("name") or value.get("measure") or value.get("id") or label)
        for key, item in value.items():
            if key in JSON_FORMULA_KEYS and isinstance(item, str) and item.strip():
                expressions.append(DaxExpression(label=f"{label}:{name}.{key}", text=item))
            else:
                expressions.extend(extract_json_expressions(item, f"{label}:{key}"))
    elif isinstance(value, list):
        for index, item in enumerate(value):
            expressions.extend(extract_json_expressions(item, f"{label}[{index}]"))
    return expressions


def load_expressions(source: Path) -> list[DaxExpression]:
    expressions: list[DaxExpression] = []
    for path in iter_source_files(source):
        suffix = path.suffix.lower()
        if suffix == ".json":
            try:
                data = json.loads(read_text(path))
            except json.JSONDecodeError as exc:
                expressions.append(DaxExpression(label=str(path), text=f"/* JSON parse error: {exc} */"))
                continue
            expressions.extend(extract_json_expressions(data, str(path)))
        elif suffix in SOURCE_EXTENSIONS:
            text = read_text(path)
            if text.strip():
                expressions.append(DaxExpression(label=str(path), text=text))
    return expressions


def lint_expression(expr: DaxExpression, profile: str, warn_division: bool) -> tuple[list[dict[str, Any]], dict[str, int]]:
    issues: list[dict[str, Any]] = []
    cleaned = strip_comments_and_strings(expr.text)
    functions = [match.group(1).upper() for match in FUNCTION_RE.finditer(cleaned)]
    counts: dict[str, int] = {}
    for function in functions:
        counts[function] = counts.get(function, 0) + 1

    balance_error = find_unbalanced_parentheses(cleaned)
    if balance_error:
        issues.append(
            {
                "severity": "error",
                "code": "unbalanced-parentheses",
                "location": expr.label,
                "message": balance_error,
            }
        )

    if profile == "excel":
        for function, suggestion in EXCEL_ERROR_FUNCTIONS.items():
            if counts.get(function, 0):
                issues.append(
                    {
                        "severity": "error",
                        "code": "excel-incompatible-function",
                        "location": expr.label,
                        "function": function,
                        "count": counts[function],
                        "message": suggestion,
                    }
                )
        for function, suggestion in EXCEL_WARNING_FUNCTIONS.items():
            if counts.get(function, 0):
                issues.append(
                    {
                        "severity": "warning",
                        "code": "excel-version-sensitive-function",
                        "location": expr.label,
                        "function": function,
                        "count": counts[function],
                        "message": suggestion,
                    }
                )

    if warn_division and "/" in cleaned:
        issues.append(
            {
                "severity": "warning",
                "code": "operator-division",
                "location": expr.label,
                "message": "The / operator appears in the expression. Prefer DIVIDE(numerator, denominator) for ratios with possible zero or blank denominators.",
            }
        )

    return issues, counts


def lint_source(source: Path, profile: str, warn_division: bool, warnings_as_errors: bool) -> dict[str, Any]:
    expressions = load_expressions(source)
    all_issues: list[dict[str, Any]] = []
    function_counts: dict[str, int] = {}

    for expr in expressions:
        issues, counts = lint_expression(expr, profile, warn_division)
        all_issues.extend(issues)
        for function, count in counts.items():
            function_counts[function] = function_counts.get(function, 0) + count

    if warnings_as_errors:
        for issue in all_issues:
            if issue["severity"] == "warning":
                issue["severity"] = "error"

    errors = [issue for issue in all_issues if issue["severity"] == "error"]
    warnings = [issue for issue in all_issues if issue["severity"] == "warning"]

    return {
        "source": str(source),
        "profile": profile,
        "expressionCount": len(expressions),
        "functionCounts": dict(sorted(function_counts.items())),
        "issues": all_issues,
        "errorCount": len(errors),
        "warningCount": len(warnings),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="DAX source file, JSON model report, or directory")
    parser.add_argument("--profile", choices=["excel", "generic"], default="excel", help="Compatibility profile")
    parser.add_argument("--warn-division", action="store_true", help="Warn when the / operator appears")
    parser.add_argument("--warnings-as-errors", action="store_true", help="Treat warnings as errors")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.expanduser().resolve()
    report = lint_source(source, args.profile, args.warn_division, args.warnings_as_errors)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if report["errorCount"]:
        print("DAX compatibility lint failed:")
        for issue in report["issues"]:
            if issue["severity"] == "error":
                function = f" [{issue.get('function')}]" if issue.get("function") else ""
                print(f"- {issue['location']}{function}: {issue['message']}")
        return 1

    print(
        "DAX compatibility lint OK: "
        f"{report['expressionCount']} expressions, "
        f"{len(report['functionCounts'])} distinct functions, "
        f"{report['warningCount']} warnings"
    )
    for issue in report["issues"]:
        if issue["severity"] == "warning":
            function = f" [{issue.get('function')}]" if issue.get("function") else ""
            print(f"- warning: {issue['location']}{function}: {issue['message']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
