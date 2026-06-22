#!/usr/bin/env python3
"""Analyze impact before renaming or deleting Power Pivot measures.

The analyzer combines model JSON and OpenXML workbook inspection JSON to find
references in DAX measure formulas and report-layer CUBE formulas. It is a
static guard: it does not open Excel, rewrite formulas, refresh the model, or
prove numeric equivalence after a rename.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

from mdx_references import cube_measure_refs as extract_cube_measure_refs

BRACKET_RE = re.compile(r"\[([^\]]+)\]")
TABLE_COLUMN_RE = re.compile(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)\s*\[([^\]]+)\]")


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def table_name(value: Any) -> str:
    if isinstance(value, dict):
        return str(value.get("name", ""))
    return str(value or "")


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


def extract_dax_measure_refs(formula: str) -> set[str]:
    cleaned = strip_comments_and_strings(formula)
    table_column_spans = [match.span() for match in TABLE_COLUMN_RE.finditer(cleaned)]
    refs: set[str] = set()
    for match in BRACKET_RE.finditer(cleaned):
        start, end = match.span()
        if any(span_start <= start and end <= span_end for span_start, span_end in table_column_spans):
            continue
        name = match.group(1).strip()
        if name:
            refs.add(name)
    return refs


def model_measures(model: dict[str, Any]) -> list[dict[str, Any]]:
    measures: list[dict[str, Any]] = []
    for item in model.get("measures", []):
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        if not name:
            continue
        measures.append(
            {
                "name": name,
                "associatedTable": table_name(item.get("associatedTable")),
                "formula": str(item.get("formula", "") or item.get("expression", "") or item.get("daxFormula", "")),
            }
        )
    return measures


def cube_formulas(openxml: dict[str, Any]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for item in openxml.get("cubeFormulas", []):
        if not isinstance(item, dict):
            continue
        formula = str(item.get("formula", ""))
        records.append(
            {
                "sheet": str(item.get("sheet", "")),
                "cell": str(item.get("cell", "")),
                "address": f"{item.get('sheet', '')}!{item.get('cell', '')}",
                "formula": formula,
                "cachedValue": item.get("cachedValue"),
                "measureReferences": extract_cube_measure_refs(formula),
            }
        )
    return records


def parse_rename(value: str) -> dict[str, str]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("rename must use Old=New")
    old, new = value.split("=", 1)
    old = old.strip()
    new = new.strip()
    if not old or not new:
        raise argparse.ArgumentTypeError("rename must include both old and new measure names")
    return {"old": old, "new": new, "operation": "rename"}


def parse_delete(value: str) -> dict[str, str]:
    value = value.strip()
    if not value:
        raise argparse.ArgumentTypeError("measure name cannot be empty")
    return {"old": value, "new": "", "operation": "delete"}


def risk_level(source_exists: bool, target_exists: bool, dax_hits: int, cube_hits: int, operation: str) -> str:
    if not source_exists:
        return "warning"
    if operation == "rename" and target_exists:
        return "error"
    if dax_hits or cube_hits:
        return "high"
    return "low"


def analyze(model: dict[str, Any], openxml: dict[str, Any], changes: list[dict[str, str]]) -> dict[str, Any]:
    measures = model_measures(model)
    formulas = cube_formulas(openxml)
    measure_names = {item["name"].lower(): item["name"] for item in measures}
    results: list[dict[str, Any]] = []

    for change in changes:
        old = change["old"]
        new = change["new"]
        old_key = old.lower()
        new_key = new.lower()

        dax_hits = []
        for measure in measures:
            refs = {ref.lower() for ref in extract_dax_measure_refs(measure["formula"])}
            if old_key in refs:
                dax_hits.append(
                    {
                        "measure": measure["name"],
                        "associatedTable": measure["associatedTable"],
                        "formula": measure["formula"],
                    }
                )

        cube_hits = []
        for record in formulas:
            refs = {ref.lower() for ref in record["measureReferences"]}
            if old_key in refs:
                cube_hits.append(record)

        source_exists = old_key in measure_names
        target_exists = bool(new) and new_key in measure_names
        result = {
            "operation": change["operation"],
            "oldMeasure": old,
            "newMeasure": new,
            "sourceMeasureExists": source_exists,
            "targetMeasureExists": target_exists,
            "riskLevel": risk_level(source_exists, target_exists, len(dax_hits), len(cube_hits), change["operation"]),
            "daxFormulaHitCount": len(dax_hits),
            "cubeFormulaHitCount": len(cube_hits),
            "affectedSheets": sorted({hit["sheet"] for hit in cube_hits if hit.get("sheet")}),
            "daxFormulaHits": dax_hits,
            "cubeFormulaHits": cube_hits,
        }
        recommendations: list[str] = []
        if not source_exists:
            recommendations.append("Confirm the old measure name; it was not found in the model JSON.")
        if target_exists:
            recommendations.append("Choose a different new measure name or explicitly merge definitions; the target name already exists.")
        if dax_hits:
            recommendations.append("Update dependent DAX measures before or together with the measure rename.")
        if cube_hits:
            recommendations.append("Update report-layer CUBE formulas or helper cells that reference the old measure path.")
        if not dax_hits and not cube_hits and source_exists:
            recommendations.append("No static DAX or CUBE formula references were found; still validate PivotTables and live Excel calculation.")
        result["recommendations"] = recommendations
        results.append(result)

    return {
        "workbookPath": openxml.get("workbookPath") or model.get("workbookPath"),
        "measureCount": len(measures),
        "cubeFormulaCount": len(formulas),
        "changeCount": len(results),
        "changes": results,
    }


def markdown_table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    def clean(value: Any) -> str:
        return str(value if value is not None else "").replace("\n", "<br>").replace("|", "\\|")

    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(clean(item) for item in row) + " |")
    return lines


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Measure Rename Impact Report",
        "",
        f"- Workbook: `{report.get('workbookPath', '')}`",
        f"- Measures: `{report.get('measureCount')}`",
        f"- CUBE formulas: `{report.get('cubeFormulaCount')}`",
        f"- Changes: `{report.get('changeCount')}`",
        "",
        "## Summary",
        "",
    ]
    rows = []
    for item in report["changes"]:
        label = f"{item['oldMeasure']} -> {item['newMeasure']}" if item["operation"] == "rename" else f"delete {item['oldMeasure']}"
        rows.append(
            [
                item["operation"],
                label,
                item["riskLevel"],
                item["daxFormulaHitCount"],
                item["cubeFormulaHitCount"],
                ", ".join(item["affectedSheets"]),
            ]
        )
    lines.extend(markdown_table(["Operation", "Measure", "Risk", "DAX hits", "CUBE hits", "Affected sheets"], rows))
    lines.append("")
    for item in report["changes"]:
        heading = f"{item['oldMeasure']} -> {item['newMeasure']}" if item["operation"] == "rename" else f"delete {item['oldMeasure']}"
        lines.extend([f"## {heading}", ""])
        lines.extend(["Recommendations:", ""])
        for recommendation in item["recommendations"]:
            lines.append(f"- {recommendation}")
        lines.append("")
        if item["daxFormulaHits"]:
            lines.extend(["### DAX Formula Hits", ""])
            lines.extend(markdown_table(["Measure", "Table", "Formula"], [[hit["measure"], hit["associatedTable"], hit["formula"]] for hit in item["daxFormulaHits"]]))
            lines.append("")
        if item["cubeFormulaHits"]:
            lines.extend(["### CUBE Formula Hits", ""])
            lines.extend(markdown_table(["Sheet", "Cell", "Cached", "Formula"], [[hit["sheet"], hit["cell"], hit.get("cachedValue"), hit["formula"]] for hit in item["cubeFormulaHits"]]))
            lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--model-json", required=True, type=Path, help="Data Model JSON or model summary JSON")
    parser.add_argument("--openxml-json", required=True, type=Path, help="OpenXML workbook inspection JSON")
    parser.add_argument("--rename", action="append", type=parse_rename, default=[], help="Measure rename mapping, e.g. Revenue=Net Revenue")
    parser.add_argument("--delete", action="append", type=parse_delete, default=[], help="Measure name planned for deletion")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--out-md", type=Path, help="Optional Markdown report path")
    parser.add_argument("--fail-on-high-risk", action="store_true", help="Exit with 1 when a high-risk or error change is detected")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    changes = list(args.rename) + list(args.delete)
    if not changes:
        raise SystemExit("At least one --rename Old=New or --delete Measure is required.")
    model = load_json(args.model_json.expanduser().resolve())
    openxml = load_json(args.openxml_json.expanduser().resolve())
    report = analyze(model, openxml, changes)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")

    high_risk = [item for item in report["changes"] if item["riskLevel"] in {"high", "error"}]
    print(
        f"Measure rename impact: {report['changeCount']} changes, "
        f"{sum(item['daxFormulaHitCount'] for item in report['changes'])} DAX hits, "
        f"{sum(item['cubeFormulaHitCount'] for item in report['changes'])} CUBE hits"
    )
    return 1 if args.fail_on_high_risk and high_risk else 0


if __name__ == "__main__":
    raise SystemExit(main())
