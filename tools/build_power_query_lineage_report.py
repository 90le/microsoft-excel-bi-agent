#!/usr/bin/env python3
"""Build a static lineage and source-risk report for exported Power Query M files.

Input is a directory of exported ``.m`` or ``.pq`` files, optionally with the
``power_queries.json`` manifest produced by the Excel COM export script. The
report does not evaluate M, refresh sources, inspect credentials, or prove data
privacy behavior. It highlights query graph and source risks visible from M text.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SOURCE_EXTENSIONS = {".m", ".pq", ".txt"}
SEVERITY_ORDER = {"info": 0, "low": 1, "medium": 2, "high": 3}
LOCAL_PATH_RE = re.compile(r"(?i)(?:[A-Z]:\\|file:///|\\\\|/Users/|/home/)")
WEB_URL_RE = re.compile(r"(?i)\bhttps?://")
NATIVE_QUERY_RE = re.compile(r"(?i)\bValue\.NativeQuery\s*\(")
SENSITIVE_STRING_RE = re.compile(
    r"(?i)\b("
    r"password|pwd|account[_ -]?key|sharedaccesssignature|sas[_ -]?token|"
    r"api[_ -]?key|client[_ -]?secret|authorization|bearer|"
    r"access[_ -]?token|refresh[_ -]?token|secret"
    r")\b"
)
SENSITIVE_KEY_RE = re.compile(
    r"(?i)\b("
    r"Password|Pwd|AccountKey|SharedAccessSignature|ApiKey|Api_Key|"
    r"ClientSecret|Client_Secret|Authorization|AccessToken|RefreshToken|Secret"
    r")\b\s*="
)
SOURCE_FUNCTION_RE = re.compile(
    r"(?i)\b("
    r"Folder\.Files|Folder\.Contents|File\.Contents|Excel\.Workbook|Excel\.CurrentWorkbook|Csv\.Document|Json\.Document|"
    r"Web\.Contents|SharePoint\.Files|SharePoint\.Contents|OData\.Feed|"
    r"Sql\.Database|Odbc\.DataSource|OleDb\.DataSource|AnalysisServices\.Database|Access\.Database|"
    r"Oracle\.Database|MySQL\.Database|PostgreSQL\.Database|Snowflake\.Databases|GoogleBigQuery\.Database|"
    r"AzureStorage\.Blobs|AzureStorage\.DataLake|AzureStorage\.Tables|"
    r"PowerPlatform\.Dataflows|CommonDataService\.Database|Dynamics365BusinessCentral\.Contents"
    r")\s*\("
)
ESCAPED_IDENTIFIER_RE = re.compile(r'#"(?:[^"]|"")*"')
STRING_RE = re.compile(r'"(?:[^"]|"")*"')
IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

DATABASE_FUNCTIONS = {
    "Sql.Database",
    "Odbc.DataSource",
    "OleDb.DataSource",
    "AnalysisServices.Database",
    "Access.Database",
    "Oracle.Database",
    "MySQL.Database",
    "PostgreSQL.Database",
    "Snowflake.Databases",
    "GoogleBigQuery.Database",
}
WEB_FUNCTIONS = {"Web.Contents", "SharePoint.Files", "SharePoint.Contents", "OData.Feed"}
LOCAL_FUNCTIONS = {"Folder.Files", "Folder.Contents", "File.Contents"}
CONFIG_FUNCTIONS = {"Excel.CurrentWorkbook"}
CLOUD_SERVICE_FUNCTIONS = {
    "AzureStorage.Blobs",
    "AzureStorage.DataLake",
    "AzureStorage.Tables",
    "PowerPlatform.Dataflows",
    "CommonDataService.Database",
    "Dynamics365BusinessCentral.Contents",
}
NON_PRIVACY_SOURCE_KINDS = {"transform-or-parser", "workbook-config"}


@dataclass
class QuerySource:
    name: str
    path: Path
    text: str


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def strip_export_prefix(stem: str) -> str:
    return re.sub(r"^\d+[_ -]+", "", stem)


def decode_m_string(value: str) -> str:
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        return value[1:-1].replace('""', '"')
    return value


def decode_escaped_identifier(value: str) -> str:
    if value.startswith('#"') and value.endswith('"'):
        return value[2:-1].replace('""', '"')
    return value


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


def load_manifest_names(query_dir: Path) -> dict[str, str]:
    manifest_path = query_dir / "power_queries.json"
    if not manifest_path.is_file():
        return {}
    try:
        manifest = json.loads(read_text(manifest_path))
    except json.JSONDecodeError:
        return {}
    result: dict[str, str] = {}
    for item in manifest.get("queries", []):
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        formula_file = str(item.get("formulaFile", "")).strip()
        if not name or not formula_file:
            continue
        result[Path(formula_file).name] = name
        try:
            result[str(Path(formula_file).resolve())] = name
        except OSError:
            pass
    return result


def iter_query_files(query_dir: Path) -> list[Path]:
    if query_dir.is_file():
        return [query_dir]
    return sorted(
        path
        for path in query_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in SOURCE_EXTENSIONS
    )


def load_queries(query_dir: Path) -> list[QuerySource]:
    manifest_names = load_manifest_names(query_dir if query_dir.is_dir() else query_dir.parent)
    queries: list[QuerySource] = []
    for path in iter_query_files(query_dir):
        text = read_text(path)
        if not text.strip():
            continue
        try:
            resolved_key = str(path.resolve())
        except OSError:
            resolved_key = str(path)
        name = manifest_names.get(resolved_key) or manifest_names.get(path.name) or strip_export_prefix(path.stem)
        queries.append(QuerySource(name=name, path=path, text=text))
    return queries


def source_kind(function: str, strings: list[str]) -> str:
    canonical = canonical_source_function(function)
    joined = " ".join(strings)
    if canonical in CONFIG_FUNCTIONS:
        return "workbook-config"
    if canonical in CLOUD_SERVICE_FUNCTIONS:
        return "cloud-service"
    if canonical in DATABASE_FUNCTIONS:
        return "database"
    if canonical in WEB_FUNCTIONS or WEB_URL_RE.search(joined):
        return "web"
    if canonical in LOCAL_FUNCTIONS or LOCAL_PATH_RE.search(joined):
        return "local-file"
    return "transform-or-parser"


def canonical_source_function(function: str) -> str:
    parts = function.split(".")
    if len(parts) != 2:
        return function
    return f"{parts[0][0].upper()}{parts[0][1:]}.{parts[1][0].upper()}{parts[1][1:]}"


def call_argument_text(text: str, start_index: int) -> str:
    """Return text inside the function call that started before ``start_index``.

    ``start_index`` is the character after the opening parenthesis matched by
    ``SOURCE_FUNCTION_RE``.
    """
    result: list[str] = []
    depth = 1
    index = start_index
    in_string = False
    while index < len(text):
        char = text[index]
        next_char = text[index + 1] if index + 1 < len(text) else ""
        if in_string:
            result.append(char)
            if char == '"':
                if next_char == '"':
                    result.append(next_char)
                    index += 2
                    continue
                in_string = False
            index += 1
            continue
        if char == '"':
            in_string = True
            result.append(char)
            index += 1
            continue
        if char == "(":
            depth += 1
            result.append(char)
            index += 1
            continue
        if char == ")":
            depth -= 1
            if depth == 0:
                break
            result.append(char)
            index += 1
            continue
        result.append(char)
        index += 1
    return "".join(result)


def source_functions(text: str) -> list[dict[str, Any]]:
    cleaned = strip_comments(text, mask_strings=False)
    result: list[dict[str, Any]] = []
    for match in SOURCE_FUNCTION_RE.finditer(cleaned):
        function = canonical_source_function(match.group(1))
        arguments = call_argument_text(cleaned, match.end())
        strings = [decode_m_string(item.group(0)) for item in STRING_RE.finditer(arguments)]
        result.append(
            {
                "function": function,
                "kind": source_kind(function, strings),
                "stringArguments": strings[:3],
            }
        )
    return result


def local_paths(text: str) -> list[str]:
    values: list[str] = []
    for match in STRING_RE.finditer(strip_comments(text, mask_strings=False)):
        value = decode_m_string(match.group(0))
        if LOCAL_PATH_RE.search(value):
            values.append(value)
    return sorted(set(values))


def web_urls(text: str) -> list[str]:
    values: list[str] = []
    for match in STRING_RE.finditer(strip_comments(text, mask_strings=False)):
        value = decode_m_string(match.group(0))
        if WEB_URL_RE.search(value):
            values.append(value)
    return sorted(set(values))


def native_query_count(text: str) -> int:
    return len(NATIVE_QUERY_RE.findall(strip_comments(text, mask_strings=True)))


def credential_like_indicators(text: str) -> dict[str, Any]:
    cleaned = strip_comments(text, mask_strings=False)
    indicators: set[str] = set()
    literal_count = 0
    for match in STRING_RE.finditer(cleaned):
        value = decode_m_string(match.group(0))
        matches = {item.group(1).lower() for item in SENSITIVE_STRING_RE.finditer(value)}
        if matches:
            literal_count += 1
            indicators.update(matches)

    for match in SENSITIVE_KEY_RE.finditer(strip_comments(text, mask_strings=True)):
        indicators.add(match.group(1).lower())

    return {
        "literalCount": literal_count,
        "indicators": sorted(indicators),
    }


def extract_dependencies(query: QuerySource, known_names: set[str]) -> list[str]:
    cleaned = strip_comments(query.text, mask_strings=False)
    masked = strip_comments(query.text, mask_strings=True)
    dependencies: set[str] = set()
    for match in ESCAPED_IDENTIFIER_RE.finditer(cleaned):
        candidate = decode_escaped_identifier(match.group(0))
        if candidate in known_names and candidate != query.name:
            dependencies.add(candidate)

    for name in known_names:
        if name == query.name or not IDENTIFIER_RE.match(name):
            continue
        if re.search(rf"(?<![A-Za-z0-9_]){re.escape(name)}(?![A-Za-z0-9_])", masked):
            dependencies.add(name)

    return sorted(dependencies)


def detect_cycles(graph: dict[str, list[str]]) -> list[list[str]]:
    cycles: list[list[str]] = []
    seen_cycles: set[tuple[str, ...]] = set()
    visiting: set[str] = set()
    visited: set[str] = set()
    stack: list[str] = []

    def canonical_cycle(path: list[str]) -> tuple[str, ...]:
        cycle = path[:-1]
        if not cycle:
            return tuple(path)
        rotations = [tuple(cycle[index:] + cycle[:index]) for index in range(len(cycle))]
        return min(rotations)

    def visit(node: str) -> None:
        if node in visiting:
            if node in stack:
                index = stack.index(node)
                cycle_path = stack[index:] + [node]
                key = canonical_cycle(cycle_path)
                if key not in seen_cycles:
                    seen_cycles.add(key)
                    cycles.append(cycle_path)
            return
        if node in visited:
            return
        visiting.add(node)
        stack.append(node)
        for dep in graph.get(node, []):
            visit(dep)
        stack.pop()
        visiting.remove(node)
        visited.add(node)

    for node in graph:
        visit(node)
    return cycles


def source_kind_closure(name: str, direct: dict[str, set[str]], graph: dict[str, list[str]]) -> set[str]:
    result: set[str] = set()
    visited: set[str] = set()

    def collect(node: str) -> None:
        if node in visited:
            return
        visited.add(node)
        result.update(kind for kind in direct.get(node, set()) if kind != "transform-or-parser")
        for dep in graph.get(node, []):
            collect(dep)

    collect(name)
    return result


def add_finding(
    findings: list[dict[str, Any]],
    query: str,
    code: str,
    severity: str,
    title: str,
    evidence: dict[str, Any],
    action: str,
) -> None:
    findings.append(
        {
            "query": query,
            "code": code,
            "severity": severity,
            "title": title,
            "evidence": evidence,
            "recommendedAction": action,
        }
    )


def build_report(query_dir: Path) -> dict[str, Any]:
    queries = load_queries(query_dir)
    known_names = {query.name for query in queries}
    query_reports: list[dict[str, Any]] = []
    graph: dict[str, list[str]] = {}
    direct_source_kinds: dict[str, set[str]] = {}
    findings: list[dict[str, Any]] = []
    source_function_counts: dict[str, int] = {}

    for query in queries:
        sources = source_functions(query.text)
        dependencies = extract_dependencies(query, known_names)
        graph[query.name] = dependencies
        direct_source_kinds[query.name] = {item["kind"] for item in sources}
        for source in sources:
            function = str(source.get("function", ""))
            source_function_counts[function] = source_function_counts.get(function, 0) + 1

        paths = local_paths(query.text)
        urls = web_urls(query.text)
        if paths:
            add_finding(
                findings,
                query.name,
                "hard-coded-local-path",
                "high",
                "Power Query source uses a hard-coded local filesystem path",
                {"paths": paths, "file": str(query.path)},
                "Move the source path into a parameter/config cell or confirm the delivery workbook does not depend on this local path.",
            )
        web_source_functions = [source.get("function") for source in sources if source.get("kind") == "web"]
        has_non_web_connector = any(source.get("kind") in {"cloud-service", "database"} for source in sources)
        if web_source_functions or (urls and not has_non_web_connector):
            add_finding(
                findings,
                query.name,
                "web-source",
                "medium",
                "Power Query source uses a web or SharePoint endpoint",
                {"urls": urls, "functions": web_source_functions},
                "Confirm credentials, privacy levels, and refresh availability in the target environment.",
            )
        if any(source.get("kind") == "database" for source in sources):
            add_finding(
                findings,
                query.name,
                "database-source",
                "medium",
                "Power Query source uses a database/provider endpoint",
                {"functions": [source.get("function") for source in sources if source.get("kind") == "database"]},
                "Confirm driver/provider availability, credentials, gateway, and refresh ownership before handoff.",
            )
        if any(source.get("kind") == "cloud-service" for source in sources):
            add_finding(
                findings,
                query.name,
                "cloud-service-source",
                "medium",
                "Power Query source uses a cloud service connector",
                {"functions": [source.get("function") for source in sources if source.get("kind") == "cloud-service"]},
                "Confirm tenant access, connector availability, credentials, gateway needs, and refresh ownership in the target environment.",
            )

        native_queries = native_query_count(query.text)
        if native_queries:
            add_finding(
                findings,
                query.name,
                "native-query-review",
                "medium",
                "Power Query uses Value.NativeQuery or a native SQL pass-through",
                {"count": native_queries, "file": str(query.path)},
                "Review SQL parameterization, query folding expectations, permissions, and source-owner approval before delivery.",
            )

        credential_indicators = credential_like_indicators(query.text)
        if credential_indicators["indicators"]:
            add_finding(
                findings,
                query.name,
                "credential-like-literal",
                "high",
                "Power Query text contains credential-like literals or authorization keys",
                {
                    "indicators": credential_indicators["indicators"],
                    "literalCount": credential_indicators["literalCount"],
                    "file": str(query.path),
                },
                "Remove secrets from M code. Use workbook parameters, environment-specific credential stores, or Excel/Power Query Data Source Settings instead.",
            )

        query_reports.append(
            {
                "name": query.name,
                "file": str(query.path),
                "dependencies": dependencies,
                "sourceFunctions": sources,
                "directSourceKinds": sorted(direct_source_kinds[query.name]),
            }
        )

    for cycle in detect_cycles(graph):
        add_finding(
            findings,
            " -> ".join(cycle),
            "query-dependency-cycle",
            "high",
            "Power Query dependency graph contains a cycle",
            {"cycle": cycle},
            "Break the cycle by materializing an upstream query, removing recursive references, or splitting shared logic into a separate parameter/function.",
        )

    for query_report in query_reports:
        name = str(query_report["name"])
        closure = source_kind_closure(name, direct_source_kinds, graph)
        query_report["sourceKindClosure"] = sorted(closure)
        privacy_relevant_closure = sorted(kind for kind in closure if kind not in NON_PRIVACY_SOURCE_KINDS)
        if len(privacy_relevant_closure) > 1:
            add_finding(
                findings,
                name,
                "mixed-source-lineage",
                "medium",
                "Power Query lineage combines multiple source kinds",
                {
                    "sourceKinds": privacy_relevant_closure,
                    "fullSourceKindClosure": sorted(closure),
                    "dependencies": query_report.get("dependencies", []),
                },
                "Review privacy-level behavior and refresh credentials because combining source kinds can trigger Formula.Firewall or credential prompts.",
            )

    severity_counts = {severity: 0 for severity in SEVERITY_ORDER}
    max_severity = "info"
    for finding in findings:
        severity = str(finding.get("severity", "info"))
        severity_counts[severity] = severity_counts.get(severity, 0) + 1
        if SEVERITY_ORDER.get(severity, 0) > SEVERITY_ORDER.get(max_severity, 0):
            max_severity = severity

    if not findings:
        readiness = "clean"
    elif max_severity == "high":
        readiness = "blocked-for-delivery"
    elif max_severity == "medium":
        readiness = "review-required"
    else:
        readiness = "low-risk"

    dependency_count = sum(len(item.get("dependencies", [])) for item in query_reports)
    source_kind_counts: dict[str, int] = {}
    for item in query_reports:
        for kind in item.get("sourceKindClosure", []):
            source_kind_counts[kind] = source_kind_counts.get(kind, 0) + 1

    return {
        "queryDirectory": str(query_dir),
        "summary": {
            "readiness": readiness,
            "maxSeverity": max_severity,
            "queryCount": len(query_reports),
            "dependencyCount": dependency_count,
            "findingCount": len(findings),
            "highFindingCount": severity_counts.get("high", 0),
            "mediumFindingCount": severity_counts.get("medium", 0),
            "lowFindingCount": severity_counts.get("low", 0),
            "sourceFunctionCounts": dict(sorted(source_function_counts.items())),
            "sourceKindCounts": dict(sorted(source_kind_counts.items())),
        },
        "queries": query_reports,
        "findings": findings,
        "limitations": [
            "Static M text inspection only; sources were not refreshed or authenticated.",
            "Escaped identifiers are treated as query dependencies only when they match known exported query names.",
            "A clean report means no selected static lineage/source risks were found, not that refresh will succeed.",
        ],
    }


def render_markdown(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    lines = [
        "# Power Query Lineage And Source-Risk Report",
        "",
        f"- query directory: `{report.get('queryDirectory', '')}`",
        f"- readiness: **{summary.get('readiness', '')}**",
        f"- max severity: `{summary.get('maxSeverity', '')}`",
        f"- queries: `{summary.get('queryCount', 0)}`",
        f"- dependencies: `{summary.get('dependencyCount', 0)}`",
        f"- findings: `{summary.get('findingCount', 0)}`",
        "",
        "## Findings",
        "",
        "| Code | Severity | Query | Evidence | Recommended action |",
        "|---|---:|---|---|---|",
    ]
    for finding in report.get("findings", []):
        evidence = json.dumps(finding.get("evidence", {}), ensure_ascii=False, sort_keys=True)
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(finding.get("code", "")),
                    clean_markdown(finding.get("severity", "")),
                    clean_markdown(finding.get("query", "")),
                    clean_markdown(evidence),
                    clean_markdown(finding.get("recommendedAction", "")),
                ]
            )
            + " |"
        )
    if not report.get("findings"):
        lines.append("| none | info | n/a | no reviewed source-risk findings | No action required by this static check. |")

    lines.extend(["", "## Query Graph", "", "| Query | Depends On | Direct Source Kinds | Source Kind Closure | Source Functions |", "|---|---|---|---|---|"])
    for query in report.get("queries", []):
        functions = [source.get("function", "") for source in query.get("sourceFunctions", [])]
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(query.get("name", "")),
                    clean_markdown(", ".join(query.get("dependencies", []))),
                    clean_markdown(", ".join(query.get("directSourceKinds", []))),
                    clean_markdown(", ".join(query.get("sourceKindClosure", []))),
                    clean_markdown(", ".join(functions)),
                ]
            )
            + " |"
        )

    lines.extend(["", "## Limitations", ""])
    for item in report.get("limitations", []):
        lines.append(f"- {item}")
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("query_dir", type=Path, help="Directory or single file containing exported Power Query M source")
    parser.add_argument("--out-json", type=Path, help="Write JSON report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--fail-on-high-risk", action="store_true", help="Exit non-zero when high-risk findings exist")
    parser.add_argument("--fail-on-review-required", action="store_true", help="Exit non-zero when medium or high findings exist")
    args = parser.parse_args()

    query_dir = args.query_dir.expanduser().resolve()
    report = build_report(query_dir)
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), report)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    summary = report.get("summary", {})
    if args.fail_on_high_risk and int(summary.get("highFindingCount") or 0) > 0:
        return 1
    if args.fail_on_review_required and int(summary.get("mediumFindingCount") or 0) + int(summary.get("highFindingCount") or 0) > 0:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
