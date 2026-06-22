#!/usr/bin/env python3
"""Analyze DAX measure dependencies from model JSON or DAX source.

This static analyzer identifies measure-to-measure references, missing measure
references, direct self references, and dependency cycles. It does not evaluate
DAX, validate table/column references, inspect live relationships, or replace
Excel Power Pivot calculation checks.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SOURCE_EXTENSIONS = {".dax", ".txt", ".md"}
FORMULA_KEYS = {"formula", "expression", "daxFormula"}
MEASURE_DEF_RE = re.compile(r"^\s*(?:MEASURE\s+)?(?:'[^']+'\s*\[([^\]]+)\]|([A-Za-z_][^\r\n:=]*?))\s*(?::=|=)\s*(.+)$", re.IGNORECASE)
BRACKET_RE = re.compile(r"\[([^\]]+)\]")
TABLE_COLUMN_RE = re.compile(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)\s*\[([^\]]+)\]")


@dataclass
class Measure:
    name: str
    formula: str
    source: str
    table: str = ""


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


def table_name(value: Any) -> str:
    if isinstance(value, dict):
        return str(value.get("name", ""))
    return str(value or "")


def iter_source_files(source: Path) -> list[Path]:
    if source.is_file():
        return [source]
    return sorted(
        path
        for path in source.rglob("*")
        if path.is_file() and (path.suffix.lower() in SOURCE_EXTENSIONS or path.suffix.lower() == ".json")
    )


def extract_json_measures(value: Any, label: str) -> list[Measure]:
    measures: list[Measure] = []
    if isinstance(value, dict):
        if isinstance(value.get("measures"), list):
            for item in value["measures"]:
                if isinstance(item, dict) and str(item.get("name", "")).strip():
                    measures.append(
                        Measure(
                            name=str(item.get("name", "")).strip(),
                            formula=str(item.get("formula", "") or item.get("expression", "") or item.get("daxFormula", "")),
                            source=f"{label}:measures[{item.get('name', '')}]",
                            table=table_name(item.get("associatedTable")),
                        )
                    )
        name = str(value.get("name") or value.get("measure") or value.get("id") or "")
        formula = ""
        for key in FORMULA_KEYS:
            if isinstance(value.get(key), str) and value.get(key, "").strip():
                formula = str(value[key])
                break
        if name and formula:
            measures.append(Measure(name=name.strip(), formula=formula, source=f"{label}:{name}", table=table_name(value.get("associatedTable"))))
        for key, item in value.items():
            if key != "measures":
                measures.extend(extract_json_measures(item, f"{label}:{key}"))
    elif isinstance(value, list):
        for index, item in enumerate(value):
            measures.extend(extract_json_measures(item, f"{label}[{index}]"))
    return measures


def parse_dax_text(text: str, label: str) -> list[Measure]:
    measures: list[Measure] = []
    current_name = ""
    current_formula: list[str] = []
    current_line = 0

    def flush() -> None:
        nonlocal current_name, current_formula, current_line
        if current_name and current_formula:
            measures.append(Measure(name=current_name.strip(), formula="\n".join(current_formula).strip(), source=f"{label}:{current_line}"))
        current_name = ""
        current_formula = []
        current_line = 0

    for line_no, line in enumerate(text.splitlines(), start=1):
        match = MEASURE_DEF_RE.match(line)
        if match:
            flush()
            current_name = (match.group(1) or match.group(2) or "").strip()
            current_formula = [match.group(3).strip()]
            current_line = line_no
        elif current_name:
            current_formula.append(line)
    flush()
    if not measures and text.strip():
        measures.append(Measure(name=Path(label).stem, formula=text, source=label))
    return measures


def load_measures(source: Path) -> list[Measure]:
    measures: list[Measure] = []
    for path in iter_source_files(source):
        suffix = path.suffix.lower()
        if suffix == ".json":
            try:
                data = json.loads(read_text(path))
            except json.JSONDecodeError:
                continue
            measures.extend(extract_json_measures(data, str(path)))
        elif suffix in SOURCE_EXTENSIONS:
            measures.extend(parse_dax_text(read_text(path), str(path)))
    unique: dict[str, Measure] = {}
    for measure in measures:
        key = measure.name.lower()
        if key not in unique:
            unique[key] = measure
    return list(unique.values())


def extract_measure_references(formula: str) -> list[str]:
    cleaned = strip_comments_and_strings(formula)
    table_column_spans = [match.span() for match in TABLE_COLUMN_RE.finditer(cleaned)]
    refs: list[str] = []
    for match in BRACKET_RE.finditer(cleaned):
        start, end = match.span()
        if any(span_start <= start and end <= span_end for span_start, span_end in table_column_spans):
            continue
        name = match.group(1).strip()
        if name and not name.lower().startswith("measures]."):
            refs.append(name)
    return sorted(set(refs), key=str.lower)


def find_cycles(graph: dict[str, list[str]]) -> list[list[str]]:
    cycles: list[list[str]] = []
    visiting: list[str] = []
    visited: set[str] = set()

    def dfs(node: str) -> None:
        if node in visiting:
            start = visiting.index(node)
            cycle = visiting[start:] + [node]
            normalized = [item.lower() for item in cycle]
            if not any([item.lower() for item in existing] == normalized for existing in cycles):
                cycles.append(cycle)
            return
        if node in visited:
            return
        visiting.append(node)
        for neighbor in graph.get(node, []):
            if neighbor in graph:
                dfs(neighbor)
        visiting.pop()
        visited.add(node)

    for node in graph:
        dfs(node)
    return cycles


def analyze(source: Path, warnings_as_errors: bool) -> dict[str, Any]:
    measures = load_measures(source)
    names_by_lower = {measure.name.lower(): measure.name for measure in measures}
    dependencies: dict[str, list[str]] = {}
    missing: list[dict[str, Any]] = []
    self_refs: list[dict[str, Any]] = []
    duplicate_sources: dict[str, list[str]] = {}
    seen_sources: dict[str, list[str]] = {}

    for measure in measures:
        seen_sources.setdefault(measure.name.lower(), []).append(measure.source)
        refs = extract_measure_references(measure.formula)
        resolved: list[str] = []
        missing_refs: list[str] = []
        for ref in refs:
            resolved_name = names_by_lower.get(ref.lower())
            if resolved_name:
                resolved.append(resolved_name)
            else:
                missing_refs.append(ref)
        dependencies[measure.name] = sorted(set(resolved), key=str.lower)
        if any(ref.lower() == measure.name.lower() for ref in resolved):
            self_refs.append({"measure": measure.name, "source": measure.source})
        if missing_refs:
            missing.append({"measure": measure.name, "source": measure.source, "missingReferences": sorted(set(missing_refs), key=str.lower)})

    for key, sources in seen_sources.items():
        if len(sources) > 1:
            duplicate_sources[names_by_lower.get(key, key)] = sources

    cycles = find_cycles(dependencies)
    issues: list[dict[str, Any]] = []
    for item in missing:
        issues.append({"severity": "error", "code": "missing-measure-reference", **item})
    for item in self_refs:
        issues.append({"severity": "error", "code": "self-reference", **item})
    for cycle in cycles:
        issues.append({"severity": "error", "code": "dependency-cycle", "cycle": cycle})
    for name, sources in duplicate_sources.items():
        issues.append({"severity": "warning", "code": "duplicate-measure-name", "measure": name, "sources": sources})

    if warnings_as_errors:
        for issue in issues:
            if issue["severity"] == "warning":
                issue["severity"] = "error"

    errors = [issue for issue in issues if issue["severity"] == "error"]
    warnings = [issue for issue in issues if issue["severity"] == "warning"]
    return {
        "source": str(source),
        "measureCount": len(measures),
        "measures": [{"name": measure.name, "table": measure.table, "source": measure.source} for measure in measures],
        "dependencies": dependencies,
        "issues": issues,
        "errorCount": len(errors),
        "warningCount": len(warnings),
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "# DAX Measure Dependency Report",
        "",
        f"- Source: `{report['source']}`",
        f"- Measures: `{report['measureCount']}`",
        f"- Errors: `{report['errorCount']}`",
        f"- Warnings: `{report['warningCount']}`",
        "",
        "## Dependencies",
        "",
        "| Measure | References |",
        "| --- | --- |",
    ]
    for name, refs in report["dependencies"].items():
        lines.append(f"| {name} | {', '.join(refs) if refs else '-'} |")
    if report["issues"]:
        lines.extend(["", "## Issues", ""])
        for issue in report["issues"]:
            lines.append(f"- `{issue['severity']}` `{issue['code']}`: {json.dumps(issue, ensure_ascii=False)}")
    return "\n".join(lines) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="DAX source file, JSON model/report file, or directory")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--out-md", type=Path, help="Optional Markdown report path")
    parser.add_argument("--warnings-as-errors", action="store_true", help="Treat duplicate-name warnings as errors")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.expanduser().resolve()
    report = analyze(source, args.warnings_as_errors)
    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_report(report), encoding="utf-8")
    print(
        f"DAX dependency analysis: {report['measureCount']} measures, "
        f"{report['errorCount']} errors, {report['warningCount']} warnings"
    )
    return 1 if report["errorCount"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
