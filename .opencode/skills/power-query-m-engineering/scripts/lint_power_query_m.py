#!/usr/bin/env python3
"""Static Power Query M lint for workbook query edits.

This is a source-level guard before M code is imported into Excel or Power BI.
It does not compile M, refresh sources, evaluate privacy levels, validate
credentials, or replace host refresh testing.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SOURCE_EXTENSIONS = {".m", ".pq", ".txt", ".md"}
JSON_FORMULA_KEYS = {
    "formula",
    "Formula",
    "expression",
    "Expression",
    "mFormula",
    "MFormula",
    "query",
    "Query",
}
IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
QUOTED_IDENTIFIER_RE = re.compile(r'^#"(?:[^"]|"")*"$')


@dataclass
class QuerySource:
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


def iter_source_files(source: Path) -> list[Path]:
    if source.is_file():
        return [source]
    return sorted(
        path
        for path in source.rglob("*")
        if path.is_file() and (path.suffix.lower() in SOURCE_EXTENSIONS or path.suffix.lower() == ".json")
    )


def extract_json_queries(value: Any, label: str) -> list[QuerySource]:
    queries: list[QuerySource] = []
    if isinstance(value, dict):
        name = str(value.get("name") or value.get("Name") or value.get("id") or label)
        for key, item in value.items():
            if key in JSON_FORMULA_KEYS and isinstance(item, str) and item.strip():
                queries.append(QuerySource(label=f"{label}:{name}.{key}", text=item))
            else:
                queries.extend(extract_json_queries(item, f"{label}:{key}"))
    elif isinstance(value, list):
        for index, item in enumerate(value):
            queries.extend(extract_json_queries(item, f"{label}[{index}]"))
    return queries


def load_queries(source: Path) -> list[QuerySource]:
    queries: list[QuerySource] = []
    for path in iter_source_files(source):
        if path.suffix.lower() == ".json":
            try:
                data = json.loads(read_text(path))
            except json.JSONDecodeError as exc:
                queries.append(QuerySource(label=str(path), text=f"let Source = \"JSON parse error: {exc}\" in MissingStep"))
                continue
            queries.extend(extract_json_queries(data, str(path)))
        else:
            text = read_text(path)
            if text.strip():
                queries.append(QuerySource(label=str(path), text=text))
    return queries


def strip_comments(text: str, mask_strings: bool = False) -> str:
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
                    result.extend("  " if mask_strings else '""')
                    index += 2
                else:
                    in_string = False
                    result.append(" " if mask_strings else char)
                    index += 1
            else:
                result.append("\n" if char in "\r\n" else (" " if mask_strings else char))
                index += 1
            continue

        if char == "/" and next_char == "/":
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
            result.append(" " if mask_strings else char)
            index += 1
            continue

        result.append(char)
        index += 1

    return "".join(result)


def bracket_balance_errors(text: str) -> list[str]:
    pairs = {")": "(", "]": "[", "}": "{"}
    opening = set(pairs.values())
    stack: list[tuple[str, int]] = []
    for index, char in enumerate(text, start=1):
        if char in opening:
            stack.append((char, index))
        elif char in pairs:
            if not stack or stack[-1][0] != pairs[char]:
                return [f"extra closing {char} near character {index}"]
            stack.pop()
    if stack:
        char, index = stack[-1]
        return [f"unclosed opening {char} near character {index}"]
    return []


def is_word_at(text: str, index: int, word: str) -> bool:
    end = index + len(word)
    if text[index:end].lower() != word:
        return False
    before = text[index - 1] if index > 0 else ""
    after = text[end] if end < len(text) else ""
    return not (before.isalnum() or before == "_") and not (after.isalnum() or after == "_")


def find_top_level_in(cleaned: str) -> tuple[int, int]:
    let_match = re.search(r"\blet\b", cleaned, re.IGNORECASE)
    if not let_match:
        return -1, -1
    depth_round = depth_square = depth_curly = 0
    last_in = -1
    index = let_match.end()
    while index < len(cleaned):
        char = cleaned[index]
        if char == "(":
            depth_round += 1
        elif char == ")":
            depth_round = max(0, depth_round - 1)
        elif char == "[":
            depth_square += 1
        elif char == "]":
            depth_square = max(0, depth_square - 1)
        elif char == "{":
            depth_curly += 1
        elif char == "}":
            depth_curly = max(0, depth_curly - 1)
        elif depth_round == depth_square == depth_curly == 0 and is_word_at(cleaned, index, "in"):
            last_in = index
            index += 2
            continue
        index += 1
    return let_match.start(), last_in


def split_top_level_assignments(text: str) -> list[str]:
    assignments: list[str] = []
    start = 0
    depth_round = depth_square = depth_curly = 0
    nested_let = 0
    index = 0
    while index < len(text):
        char = text[index]
        if char == "(":
            depth_round += 1
        elif char == ")":
            depth_round = max(0, depth_round - 1)
        elif char == "[":
            depth_square += 1
        elif char == "]":
            depth_square = max(0, depth_square - 1)
        elif char == "{":
            depth_curly += 1
        elif char == "}":
            depth_curly = max(0, depth_curly - 1)
        elif depth_round == depth_square == depth_curly == 0 and is_word_at(text, index, "let"):
            nested_let += 1
            index += 3
            continue
        elif depth_round == depth_square == depth_curly == 0 and is_word_at(text, index, "in") and nested_let > 0:
            nested_let -= 1
            index += 2
            continue
        elif char == "," and depth_round == depth_square == depth_curly == 0 and nested_let == 0:
            piece = text[start:index].strip()
            if piece:
                assignments.append(piece)
            start = index + 1
        index += 1
    tail = text[start:].strip()
    if tail:
        assignments.append(tail)
    return assignments


def normalize_step_name(name: str) -> str:
    name = name.strip()
    if QUOTED_IDENTIFIER_RE.match(name):
        return name[2:-1].replace('""', '"').lower()
    return name.lower()


def parse_steps(cleaned: str) -> tuple[list[str], str, list[str]]:
    errors: list[str] = []
    let_start, in_pos = find_top_level_in(cleaned)
    if let_start < 0:
        return [], "", ["missing top-level let expression"]
    if in_pos < 0:
        return [], "", ["missing top-level in expression"]

    body = cleaned[let_start + 3 : in_pos]
    final_expr = cleaned[in_pos + 2 :].strip()
    step_names: list[str] = []
    for assignment in split_top_level_assignments(body):
        left = assignment.split("=", 1)[0].strip()
        if not left:
            continue
        if IDENTIFIER_RE.match(left) or QUOTED_IDENTIFIER_RE.match(left):
            step_names.append(left)
        else:
            errors.append(f"could not parse step name from assignment: {assignment[:80]}")

    normalized: dict[str, int] = {}
    for step in step_names:
        key = normalize_step_name(step)
        normalized[key] = normalized.get(key, 0) + 1
    duplicates = sorted(name for name, count in normalized.items() if count > 1)
    for name in duplicates:
        errors.append(f"duplicate step name: {name}")

    simple_final = final_expr.rstrip(",")
    if (IDENTIFIER_RE.match(simple_final) or QUOTED_IDENTIFIER_RE.match(simple_final)) and normalize_step_name(simple_final) not in normalized:
        errors.append(f"final in expression references undefined step: {simple_final}")

    return step_names, final_expr, errors


def has_any(text: str, values: list[str]) -> bool:
    lower = text.lower()
    return any(value.lower() in lower for value in values)


def lint_query(query: QuerySource) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    comments_removed = strip_comments(query.text, mask_strings=False)
    strings_masked = strip_comments(query.text, mask_strings=True)

    for message in bracket_balance_errors(strings_masked):
        issues.append({"severity": "error", "code": "unbalanced-delimiters", "location": query.label, "message": message})

    steps, final_expr, parse_errors = parse_steps(strings_masked)
    for message in parse_errors:
        issues.append({"severity": "error", "code": "query-structure", "location": query.label, "message": message})

    if "Folder.Files" in comments_removed:
        if "~$" not in comments_removed and "Text.StartsWith([Name]" not in comments_removed:
            issues.append(
                {
                    "severity": "warning",
                    "code": "folder-files-temp-filter",
                    "location": query.label,
                    "message": "Folder.Files is used without an obvious early filter for Excel temporary files such as names starting with ~$.",
                }
            )
        if "Hidden" not in comments_removed:
            issues.append(
                {
                    "severity": "warning",
                    "code": "folder-files-hidden-filter",
                    "location": query.label,
                    "message": "Folder.Files is used without an obvious hidden-file filter such as [Attributes]?[Hidden]? <> true.",
                }
            )

    uses_join_or_group = has_any(comments_removed, ["Table.NestedJoin", "Table.Group"])
    if uses_join_or_group and "Table.Sort" not in comments_removed:
        issues.append(
            {
                "severity": "warning",
                "code": "order-restoration",
                "location": query.label,
                "message": "Table.NestedJoin or Table.Group appears without a final Table.Sort; report-facing output order may drift.",
            }
        )

    if "Table.NestedJoin" in comments_removed and not has_any(comments_removed, ["Table.Group", "Table.Distinct", "Table.First"]):
        issues.append(
            {
                "severity": "warning",
                "code": "join-cardinality",
                "location": query.label,
                "message": "Table.NestedJoin appears without an obvious duplicate-control step on the lookup side.",
            }
        )

    if "Table.ExpandTableColumn" in comments_removed and "Table.ColumnNames" not in comments_removed:
        issues.append(
            {
                "severity": "warning",
                "code": "hard-coded-expand-columns",
                "location": query.label,
                "message": "Table.ExpandTableColumn appears with no Table.ColumnNames guard; verify schema drift will not break the query.",
            }
        )

    if "List.Max" in comments_removed and not has_any(comments_removed, ["List.IsEmpty", "try "]):
        issues.append(
            {
                "severity": "warning",
                "code": "unguarded-list-max",
                "location": query.label,
                "message": "List.Max appears without an obvious empty-list guard such as List.IsEmpty or try.",
            }
        )

    return issues, {"label": query.label, "stepCount": len(steps), "finalExpression": final_expr[:120]}


def lint_source(source: Path, warnings_as_errors: bool) -> dict[str, Any]:
    queries = load_queries(source)
    issues: list[dict[str, Any]] = []
    query_reports: list[dict[str, Any]] = []

    if not queries:
        issues.append({"severity": "error", "code": "no-query-source", "location": str(source), "message": "No Power Query M source was found."})

    for query in queries:
        query_issues, query_report = lint_query(query)
        issues.extend(query_issues)
        query_reports.append(query_report)

    if warnings_as_errors:
        for issue in issues:
            if issue["severity"] == "warning":
                issue["severity"] = "error"

    errors = [issue for issue in issues if issue["severity"] == "error"]
    warnings = [issue for issue in issues if issue["severity"] == "warning"]
    return {
        "source": str(source),
        "queryCount": len(queries),
        "queries": query_reports,
        "issues": issues,
        "errorCount": len(errors),
        "warningCount": len(warnings),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="M source file, JSON export, or directory containing .m/.pq/.txt/.md/.json files")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--warnings-as-errors", action="store_true", help="Treat warnings as errors")
    parser.add_argument("--strict", action="store_true", help="Alias for --warnings-as-errors")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.expanduser().resolve()
    report = lint_source(source, args.warnings_as_errors or args.strict)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(
        f"Power Query M lint: {report['queryCount']} queries, "
        f"{report['errorCount']} errors, {report['warningCount']} warnings"
    )
    return 1 if report["errorCount"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
