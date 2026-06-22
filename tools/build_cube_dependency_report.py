#!/usr/bin/env python3
"""Build a dependency report for Excel CUBE formulas.

The input is the JSON produced by inspect_excel_bi_workbook.py. Optionally pass
Data Model metadata from inspect_excel_data_model_com.ps1 or a normalized model
summary from build_excel_bi_model_report.py to validate measure references.

This script is cross-platform and never opens Excel.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from mdx_references import cube_measure_refs, member_refs as extract_member_refs

CUBE_FUNCTION_RE = re.compile(
    r"\b(CUBEVALUE|CUBEMEMBER|CUBESET|CUBERANKEDMEMBER|CUBEKPIMEMBER|CUBEMEMBERPROPERTY)\s*\(",
    re.IGNORECASE,
)
CELL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])(?:'[^']+'!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![A-Za-z0-9_])"
)
FIRST_QUOTED_RE = re.compile(r'^\s*[A-Z.]+\s*\(\s*"([^"]+)"', re.IGNORECASE)
HARD_CODED_PERIOD_RE = re.compile(r"\[All\]\.\[(new|-1|pre|previous|latest)\]", re.IGNORECASE)
DYNAMIC_MDX_RE = re.compile(r'"\s*&|&\s*"|\["\s*&|&\s*"\]', re.IGNORECASE)


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def markdown_table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    def clean(value: Any) -> str:
        text = "" if value is None else str(value)
        return text.replace("\n", "<br>").replace("|", "\\|")

    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(clean(item) for item in row) + " |")
    return lines


def mermaid_label(value: Any) -> str:
    return str(value or "").replace("\\", "\\\\").replace('"', '\\"')


def formula_function(formula: str) -> str:
    match = CUBE_FUNCTION_RE.search(formula)
    return match.group(1).upper() if match else ""


def connection_name(formula: str) -> str:
    match = FIRST_QUOTED_RE.search(formula)
    return match.group(1) if match else ""


def measure_refs(formula: str) -> list[str]:
    return cube_measure_refs(formula)


def member_refs(formula: str) -> list[str]:
    return extract_member_refs(formula)


def helper_cell_refs(formula: str) -> list[str]:
    refs = CELL_REF_RE.findall(formula)
    # Remove obvious references embedded inside MDX bracket strings only when
    # they are not part of a concatenation/helper reference. This keeps common
    # parameterized fragments such as "&$B5&".
    return sorted(set(refs))


def model_measure_names(model: dict[str, Any] | None) -> set[str]:
    if not model:
        return set()
    names = set()
    for item in model.get("measures", []):
        name = str(item.get("name", "")).strip()
        if name:
            names.add(name)
    for item in model.get("measuresReferencedByCubeFormulas", []):
        name = str(item).strip()
        if name:
            names.add(name)
    return names


def formula_flags(record: dict[str, Any], known_measures: set[str]) -> list[str]:
    formula = str(record.get("formula", ""))
    fn = formula_function(formula)
    measures = measure_refs(formula)
    helpers = helper_cell_refs(formula)
    flags: list[str] = []

    if fn == "CUBEVALUE" and not measures:
        flags.append("cubevalue_without_measure")
    if known_measures:
        missing = [name for name in measures if name not in known_measures]
        if missing:
            flags.append("measure_not_found_in_model")
    if HARD_CODED_PERIOD_RE.search(formula):
        flags.append("hard_coded_period_marker")
    if DYNAMIC_MDX_RE.search(formula):
        flags.append("dynamic_mdx_string")
    if len(formula) > 500 and not helpers:
        flags.append("long_formula_without_helper_cells")
    if fn in {"CUBEVALUE", "CUBEMEMBER", "CUBESET"} and not connection_name(formula):
        flags.append("dynamic_or_missing_connection_name")
    if record.get("cachedValue") in ("#N/A", "#GETTING_DATA", "#VALUE!", "#NAME?"):
        flags.append("error_cached_value")
    return flags


def build_dependency_summary(openxml: dict[str, Any], model: dict[str, Any] | None) -> dict[str, Any]:
    known_measures = model_measure_names(model)
    formulas: list[dict[str, Any]] = []
    by_sheet: Counter[str] = Counter()
    by_function: Counter[str] = Counter()
    by_measure: Counter[str] = Counter()
    by_member: Counter[str] = Counter()
    by_helper: Counter[str] = Counter()
    by_flag: Counter[str] = Counter()
    sheet_to_measures: dict[str, Counter[str]] = defaultdict(Counter)
    measure_to_cells: dict[str, list[str]] = defaultdict(list)
    helper_to_cells: dict[str, list[str]] = defaultdict(list)

    for item in openxml.get("cubeFormulas", []):
        formula = str(item.get("formula", ""))
        sheet = str(item.get("sheet", ""))
        cell = str(item.get("cell", ""))
        address = f"{sheet}!{cell}" if sheet and cell else cell
        fn = formula_function(formula)
        measures = measure_refs(formula)
        members = member_refs(formula)
        helpers = helper_cell_refs(formula)
        flags = formula_flags(item, known_measures)

        by_sheet[sheet] += 1
        by_function[fn or "UNKNOWN"] += 1
        for measure in measures:
            by_measure[measure] += 1
            sheet_to_measures[sheet][measure] += 1
            measure_to_cells[measure].append(address)
        for member in members:
            by_member[member] += 1
        for helper in helpers:
            by_helper[helper] += 1
            helper_to_cells[helper].append(address)
        for flag in flags:
            by_flag[flag] += 1

        formulas.append(
            {
                "sheet": sheet,
                "cell": cell,
                "address": address,
                "function": fn,
                "connection": connection_name(formula),
                "measures": measures,
                "memberReferences": members,
                "helperCellReferences": helpers,
                "flags": flags,
                "cachedValue": item.get("cachedValue"),
                "formula": formula,
            }
        )

    referenced_measures = set(by_measure)
    missing_measures = sorted(referenced_measures - known_measures) if known_measures else []
    known_not_referenced = sorted(known_measures - referenced_measures) if known_measures else []

    return {
        "workbookPath": openxml.get("workbookPath"),
        "cubeFormulaCount": len(formulas),
        "knownModelMeasureCount": len(known_measures),
        "bySheet": dict(by_sheet.most_common()),
        "byFunction": dict(by_function.most_common()),
        "byMeasure": dict(by_measure.most_common()),
        "byMemberReference": dict(by_member.most_common()),
        "byHelperCellReference": dict(by_helper.most_common()),
        "byDiagnosticFlag": dict(by_flag.most_common()),
        "sheetToMeasures": {sheet: dict(counter.most_common()) for sheet, counter in sheet_to_measures.items()},
        "measureToCells": {measure: cells for measure, cells in sorted(measure_to_cells.items())},
        "helperCellToCells": {helper: cells for helper, cells in sorted(helper_to_cells.items())},
        "missingModelMeasures": missing_measures,
        "modelMeasuresNotReferencedByCubeFormulas": known_not_referenced,
        "formulas": formulas,
    }


def diagnostic_recommendation(flag: str) -> str:
    return {
        "cubevalue_without_measure": "Confirm whether the formula should be CUBEMEMBER/CUBESET or add an explicit [Measures].[...] argument.",
        "measure_not_found_in_model": "Fix the measure name or create/restore the missing model measure.",
        "hard_coded_period_marker": "Consider moving period selection to a helper cell or named range so reports do not silently go stale.",
        "dynamic_mdx_string": "Prefer CUBEMEMBER helper cells for parameterized members, then reference those helpers in CUBEVALUE.",
        "long_formula_without_helper_cells": "Split long MDX member strings into auditable helper CUBEMEMBER cells.",
        "dynamic_or_missing_connection_name": "Use a stable connection name or inspect the workbook connection before rewriting.",
        "error_cached_value": "Open in Excel, refresh the model, and verify the member path or measure reference.",
    }.get(flag, "Inspect the formula and validate the returned value in Excel.")


def render_mermaid(summary: dict[str, Any], max_edges: int = 80) -> str:
    lines = ["graph LR"]
    edge_count = 0
    for sheet, measures in summary.get("sheetToMeasures", {}).items():
        sheet_id = "s_" + re.sub(r"\W+", "_", sheet)
        lines.append(f'  {sheet_id}["Sheet: {mermaid_label(sheet)}"]')
        for measure, count in measures.items():
            if edge_count >= max_edges:
                lines.append(f'  more["... {sum(len(v) for v in summary.get("sheetToMeasures", {}).values()) - edge_count} more edges"]')
                return "\n".join(lines)
            measure_id = "m_" + re.sub(r"\W+", "_", measure)
            lines.append(f'  {measure_id}["Measure: {mermaid_label(measure)}"]')
            lines.append(f'  {sheet_id} -->|"{count} formulas"| {measure_id}')
            edge_count += 1
    return "\n".join(lines)


def render_markdown(summary: dict[str, Any], detail_limit: int) -> str:
    lines: list[str] = [
        "# Excel CUBE Dependency Report",
        "",
        f"- workbook: `{summary.get('workbookPath', '')}`",
        f"- CUBE formulas: {summary.get('cubeFormulaCount', 0)}",
        f"- known model measures: {summary.get('knownModelMeasureCount', 0)}",
        f"- diagnostic flags: {sum(summary.get('byDiagnosticFlag', {}).values())}",
        "",
    ]

    if summary.get("cubeFormulaCount"):
        lines.extend(["## Dependency Graph", "", "```mermaid", render_mermaid(summary), "```", ""])

    if summary.get("bySheet"):
        lines.extend(["## Formulas By Sheet", ""])
        lines.extend(markdown_table(["Sheet", "CUBE Formulas"], [[k, v] for k, v in summary["bySheet"].items()]))
        lines.append("")

    if summary.get("byMeasure"):
        lines.extend(["## Formulas By Measure", ""])
        lines.extend(markdown_table(["Measure", "CUBE Formulas"], [[k, v] for k, v in summary["byMeasure"].items()]))
        lines.append("")

    if summary.get("missingModelMeasures"):
        lines.extend(["## Missing Model Measures", ""])
        for measure in summary["missingModelMeasures"]:
            lines.append(f"- `{measure}`")
        lines.append("")

    if summary.get("modelMeasuresNotReferencedByCubeFormulas"):
        lines.extend(["## Model Measures Not Referenced By CUBE Formulas", ""])
        for measure in summary["modelMeasuresNotReferencedByCubeFormulas"]:
            lines.append(f"- `{measure}`")
        lines.append("")

    if summary.get("byDiagnosticFlag"):
        lines.extend(["## Diagnostics", ""])
        rows = [
            [flag, count, diagnostic_recommendation(flag)]
            for flag, count in summary["byDiagnosticFlag"].items()
        ]
        lines.extend(markdown_table(["Flag", "Count", "Recommendation"], rows))
        lines.append("")

    if summary.get("byMemberReference"):
        lines.extend(["## Top Member References", ""])
        rows = list(summary["byMemberReference"].items())[:50]
        lines.extend(markdown_table(["Member Reference", "Count"], rows))
        lines.append("")

    if summary.get("byHelperCellReference"):
        lines.extend(["## Helper / Dynamic Cell References", ""])
        rows = list(summary["byHelperCellReference"].items())[:50]
        lines.extend(markdown_table(["Cell Reference", "Used By CUBE Formulas"], rows))
        lines.append("")

    formulas = summary.get("formulas", [])
    if formulas:
        lines.extend(["## Formula Inventory", ""])
        rows = []
        for item in formulas[:detail_limit]:
            rows.append(
                [
                    item.get("address"),
                    item.get("function"),
                    ", ".join(item.get("measures", [])),
                    ", ".join(item.get("helperCellReferences", [])),
                    ", ".join(item.get("flags", [])),
                    item.get("cachedValue"),
                ]
            )
        lines.extend(markdown_table(["Cell", "Function", "Measures", "Helper Cells", "Flags", "Cached Value"], rows))
        if len(formulas) > detail_limit:
            lines.append(f"\n_{len(formulas) - detail_limit} more formulas omitted from Markdown. See JSON output for full detail._")
        lines.append("")

    lines.extend(
        [
            "## Rewrite Workflow",
            "",
            "1. Fix missing measure references first; formula rewrites cannot compensate for absent model measures.",
            "2. Replace long or dynamic member strings with helper `CUBEMEMBER` cells where the same member is reused.",
            "3. Move period selectors such as latest/previous markers into visible helper cells or named ranges when business users must audit them.",
            "4. Validate rewritten formulas in Excel after model refresh; OpenXML inspection is structural and does not evaluate values.",
        ]
    )
    return "\n".join(lines).rstrip() + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--openxml-json", required=True, type=Path, help="JSON from inspect_excel_bi_workbook.py")
    parser.add_argument("--model-json", type=Path, help="Optional JSON from inspect_excel_data_model_com.ps1")
    parser.add_argument("--model-summary-json", type=Path, help="Optional JSON from build_excel_bi_model_report.py")
    parser.add_argument("--out-md", type=Path, help="Write Markdown report")
    parser.add_argument("--out-json", type=Path, help="Write dependency JSON")
    parser.add_argument("--out-mermaid", type=Path, help="Write Mermaid dependency graph only")
    parser.add_argument("--detail-limit", type=int, default=200, help="Formula inventory rows to include in Markdown")
    parser.add_argument("--print", action="store_true", help="Print Markdown report")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    openxml = load_json(args.openxml_json.expanduser().resolve())
    model = None
    if args.model_json:
        model = load_json(args.model_json.expanduser().resolve())
    elif args.model_summary_json:
        model = load_json(args.model_summary_json.expanduser().resolve())

    summary = build_dependency_summary(openxml, model)
    markdown = render_markdown(summary, max(args.detail_limit, 1))
    mermaid = render_mermaid(summary) + "\n"

    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Wrote CUBE dependency JSON: {args.out_json}")
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(markdown, encoding="utf-8")
        print(f"Wrote CUBE dependency Markdown: {args.out_md}")
    if args.out_mermaid:
        args.out_mermaid.parent.mkdir(parents=True, exist_ok=True)
        args.out_mermaid.write_text(mermaid, encoding="utf-8")
        print(f"Wrote CUBE dependency Mermaid: {args.out_mermaid}")
    if args.print or (not args.out_json and not args.out_md and not args.out_mermaid):
        print(markdown)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
