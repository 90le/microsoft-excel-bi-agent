#!/usr/bin/env python3
"""Build a reviewable rewrite plan for Power Pivot measure renames.

The planner combines model JSON and OpenXML workbook inspection JSON to propose
static formula rewrites for dependent DAX measures and report-layer CUBE
formulas. It does not edit a workbook, rename a model object, refresh Excel, or
prove calculation equivalence. Use the output as an auditable change plan before
applying workbook edits.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

from mdx_references import cube_measure_refs as extract_cube_measure_refs
from mdx_references import replace_cube_measure_ref

BRACKET_RE = re.compile(r"\[([^\]]+)\]")
TABLE_COLUMN_RE = re.compile(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)\s*\[([^\]]+)\]")
CELL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])(?:'[^']+'!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![A-Za-z0-9_])"
)


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
        sheet = str(item.get("sheet", ""))
        cell = str(item.get("cell", ""))
        helpers = helper_cell_refs(formula)
        records.append(
            {
                "sheet": sheet,
                "cell": cell,
                "address": f"{sheet}!{cell}" if sheet or cell else "",
                "formula": formula,
                "cachedValue": item.get("cachedValue"),
                "measureReferences": extract_cube_measure_refs(formula),
                "helperCellReferences": helpers,
                "normalizedHelperCellReferences": sorted(
                    {
                        normalized
                        for ref in helpers
                        for normalized in [normalize_cell_ref(ref, sheet)]
                        if normalized
                    }
                ),
            }
        )
    return records


def helper_cell_refs(formula: str) -> list[str]:
    return sorted(set(CELL_REF_RE.findall(formula)))


def normalize_cell_ref(ref: str, default_sheet: str) -> str:
    if ":" in ref:
        return ""
    sheet = default_sheet
    cell = ref
    if "!" in ref:
        sheet, cell = ref.rsplit("!", 1)
        sheet = sheet.strip("'")
    normalized_cell = cell.replace("$", "").upper()
    if not normalized_cell:
        return ""
    return f"{sheet.lower()}!{normalized_cell}"


def record_cell_key(record: dict[str, Any]) -> str:
    sheet = str(record.get("sheet", ""))
    cell = str(record.get("cell", "")).replace("$", "").upper()
    return f"{sheet.lower()}!{cell}" if sheet and cell else ""


def dax_table_column_spans(formula: str) -> list[tuple[int, int]]:
    cleaned = strip_comments_and_strings(formula)
    return [match.span() for match in TABLE_COLUMN_RE.finditer(cleaned)]


def span_inside(span: tuple[int, int], containers: list[tuple[int, int]]) -> bool:
    start, end = span
    return any(container_start <= start and end <= container_end for container_start, container_end in containers)


def replace_dax_measure_ref(formula: str, old: str, new: str) -> tuple[str, bool]:
    old_key = old.lower()
    protected_spans = dax_table_column_spans(formula)
    changed = False

    def repl(match: re.Match[str]) -> str:
        nonlocal changed
        if span_inside(match.span(), protected_spans):
            return match.group(0)
        if match.group(1).strip().lower() == old_key:
            changed = True
            return f"[{new}]"
        return match.group(0)

    return BRACKET_RE.sub(repl, formula), changed


def parse_rename(value: str) -> dict[str, str]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("rename must use Old=New")
    old, new = value.split("=", 1)
    old = old.strip()
    new = new.strip()
    if not old or not new:
        raise argparse.ArgumentTypeError("rename must include both old and new measure names")
    return {"operation": "rename", "old": old, "new": new}


def parse_delete(value: str) -> dict[str, str]:
    value = value.strip()
    if not value:
        raise argparse.ArgumentTypeError("measure name cannot be empty")
    return {"operation": "delete", "old": value, "new": ""}


def build_change_plan(
    measures: list[dict[str, Any]],
    formulas: list[dict[str, Any]],
    change: dict[str, str],
    measure_names: dict[str, str],
) -> dict[str, Any]:
    old = change["old"]
    new = change["new"]
    old_key = old.lower()
    new_key = new.lower()
    source_exists = old_key in measure_names
    target_exists = bool(new) and new_key in measure_names

    dax_rewrites: list[dict[str, Any]] = []
    cube_rewrites: list[dict[str, Any]] = []
    downstream_impacts: list[dict[str, Any]] = []
    manual_review: list[dict[str, Any]] = []
    directly_affected_cube_keys: set[str] = set()

    if not source_exists:
        manual_review.append(
            {
                "kind": "measure",
                "measure": old,
                "reason": "source measure was not found in the model JSON",
            }
        )
    if change["operation"] == "rename" and target_exists:
        manual_review.append(
            {
                "kind": "measure",
                "measure": new,
                "reason": "target measure already exists in the model JSON",
            }
        )

    for measure in measures:
        formula = str(measure.get("formula", ""))
        rewritten, changed = replace_dax_measure_ref(formula, old, new) if change["operation"] == "rename" else (formula, False)
        references_old = changed or old_key in {
            match.group(1).strip().lower()
            for match in BRACKET_RE.finditer(strip_comments_and_strings(formula))
            if not span_inside(match.span(), dax_table_column_spans(formula))
        }
        if change["operation"] == "rename" and changed:
            dax_rewrites.append(
                {
                    "measure": measure["name"],
                    "associatedTable": measure["associatedTable"],
                    "oldFormula": formula,
                    "newFormula": rewritten,
                }
            )
        elif change["operation"] == "delete" and references_old:
            manual_review.append(
                {
                    "kind": "dax",
                    "measure": measure["name"],
                    "associatedTable": measure["associatedTable"],
                    "reason": "deleted measure is referenced by this DAX formula",
                    "formula": formula,
                }
            )

    for record in formulas:
        formula = str(record.get("formula", ""))
        refs = {str(ref).lower() for ref in record.get("measureReferences", [])}
        if old_key not in refs:
            continue
        direct_key = record_cell_key(record)
        if direct_key:
            directly_affected_cube_keys.add(direct_key)
        if change["operation"] == "rename":
            rewritten, changed = replace_cube_measure_ref(formula, old, new)
            if changed:
                cube_rewrites.append(
                    {
                        "sheet": record["sheet"],
                        "cell": record["cell"],
                        "address": record["address"],
                        "cachedValue": record.get("cachedValue"),
                        "helperCellReferences": record.get("helperCellReferences", []),
                        "oldFormula": formula,
                        "newFormula": rewritten,
                    }
                )
        else:
            manual_review.append(
                {
                    "kind": "cube",
                    "sheet": record["sheet"],
                    "cell": record["cell"],
                    "address": record["address"],
                    "reason": "deleted measure is referenced by this CUBE formula",
                    "formula": formula,
                }
            )

    for record in formulas:
        current_key = record_cell_key(record)
        if current_key and current_key in directly_affected_cube_keys:
            continue
        refs = set(record.get("normalizedHelperCellReferences", []))
        affected_refs = sorted(refs & directly_affected_cube_keys)
        if not affected_refs:
            continue
        downstream_impacts.append(
            {
                "sheet": record["sheet"],
                "cell": record["cell"],
                "address": record["address"],
                "cachedValue": record.get("cachedValue"),
                "formula": record["formula"],
                "dependsOnAffectedCells": affected_refs,
                "reason": "formula references a helper cell whose measure path is affected by this change",
            }
        )
        if change["operation"] == "delete":
            manual_review.append(
                {
                    "kind": "cube-downstream",
                    "sheet": record["sheet"],
                    "cell": record["cell"],
                    "address": record["address"],
                    "reason": "formula depends on a helper cell that references the deleted measure",
                    "formula": record["formula"],
                    "dependsOnAffectedCells": affected_refs,
                }
            )

    return {
        "operation": change["operation"],
        "oldMeasure": old,
        "newMeasure": new,
        "sourceMeasureExists": source_exists,
        "targetMeasureExists": target_exists,
        "daxRewriteCount": len(dax_rewrites),
        "cubeRewriteCount": len(cube_rewrites),
        "downstreamFormulaImpactCount": len(downstream_impacts),
        "manualReviewCount": len(manual_review),
        "affectedSheets": sorted(
            {
                item["sheet"]
                for group in [cube_rewrites, downstream_impacts]
                for item in group
                if item.get("sheet")
            }
        ),
        "daxRewrites": dax_rewrites,
        "cubeRewrites": cube_rewrites,
        "downstreamFormulaImpacts": downstream_impacts,
        "manualReview": manual_review,
    }


def build_plan(model: dict[str, Any], openxml: dict[str, Any], changes: list[dict[str, str]]) -> dict[str, Any]:
    measures = model_measures(model)
    formulas = cube_formulas(openxml)
    measure_names = {item["name"].lower(): item["name"] for item in measures}
    plans = [build_change_plan(measures, formulas, change, measure_names) for change in changes]
    return {
        "workbookPath": openxml.get("workbookPath") or model.get("workbookPath"),
        "measureCount": len(measures),
        "cubeFormulaCount": len(formulas),
        "changeCount": len(plans),
        "daxRewriteCount": sum(item["daxRewriteCount"] for item in plans),
        "cubeRewriteCount": sum(item["cubeRewriteCount"] for item in plans),
        "downstreamFormulaImpactCount": sum(item["downstreamFormulaImpactCount"] for item in plans),
        "manualReviewCount": sum(item["manualReviewCount"] for item in plans),
        "changes": plans,
        "notes": [
            "This is a static rewrite plan; it does not edit a workbook.",
            "DAX replacement avoids table-qualified references such as Fact[Amount] to reduce false positives.",
            "Validate renamed measures and report formulas in Excel after applying edits.",
        ],
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
        "# Measure Rename Rewrite Plan",
        "",
        f"- Workbook: `{report.get('workbookPath', '')}`",
        f"- Measures: `{report.get('measureCount')}`",
        f"- CUBE formulas: `{report.get('cubeFormulaCount')}`",
        f"- Changes: `{report.get('changeCount')}`",
        f"- DAX rewrites: `{report.get('daxRewriteCount')}`",
        f"- CUBE rewrites: `{report.get('cubeRewriteCount')}`",
        f"- Downstream formula impacts: `{report.get('downstreamFormulaImpactCount')}`",
        f"- Manual review items: `{report.get('manualReviewCount')}`",
        "",
        "## Summary",
        "",
    ]
    summary_rows = []
    for item in report["changes"]:
        label = f"{item['oldMeasure']} -> {item['newMeasure']}" if item["operation"] == "rename" else f"delete {item['oldMeasure']}"
        summary_rows.append(
            [
                item["operation"],
                label,
                item["sourceMeasureExists"],
                item["targetMeasureExists"],
                item["daxRewriteCount"],
                item["cubeRewriteCount"],
                item["downstreamFormulaImpactCount"],
                item["manualReviewCount"],
                ", ".join(item["affectedSheets"]),
            ]
        )
    lines.extend(
        markdown_table(
            [
                "Operation",
                "Measure",
                "Source exists",
                "Target exists",
                "DAX rewrites",
                "CUBE rewrites",
                "Downstream impacts",
                "Manual review",
                "Sheets",
            ],
            summary_rows,
        )
    )
    lines.append("")

    for item in report["changes"]:
        heading = f"{item['oldMeasure']} -> {item['newMeasure']}" if item["operation"] == "rename" else f"delete {item['oldMeasure']}"
        lines.extend([f"## {heading}", ""])
        if item["daxRewrites"]:
            lines.extend(["### DAX Rewrites", ""])
            lines.extend(
                markdown_table(
                    ["Measure", "Table", "Old formula", "New formula"],
                    [[hit["measure"], hit["associatedTable"], hit["oldFormula"], hit["newFormula"]] for hit in item["daxRewrites"]],
                )
            )
            lines.append("")
        if item["cubeRewrites"]:
            lines.extend(["### CUBE Formula Rewrites", ""])
            lines.extend(
                markdown_table(
                    ["Sheet", "Cell", "Cached", "Helpers", "Old formula", "New formula"],
                    [
                        [
                            hit["sheet"],
                            hit["cell"],
                            hit.get("cachedValue"),
                            ", ".join(hit.get("helperCellReferences", [])),
                            hit["oldFormula"],
                            hit["newFormula"],
                        ]
                        for hit in item["cubeRewrites"]
                    ],
                )
            )
            lines.append("")
        if item["downstreamFormulaImpacts"]:
            lines.extend(["### Downstream Formula Impacts", ""])
            lines.extend(
                markdown_table(
                    ["Sheet", "Cell", "Cached", "Depends on", "Formula", "Reason"],
                    [
                        [
                            hit["sheet"],
                            hit["cell"],
                            hit.get("cachedValue"),
                            ", ".join(hit.get("dependsOnAffectedCells", [])),
                            hit["formula"],
                            hit["reason"],
                        ]
                        for hit in item["downstreamFormulaImpacts"]
                    ],
                )
            )
            lines.append("")
        if item["manualReview"]:
            lines.extend(["### Manual Review", ""])
            lines.extend(
                markdown_table(
                    ["Kind", "Location", "Reason"],
                    [
                        [
                            hit.get("kind", ""),
                            hit.get("measure") or hit.get("address") or "",
                            hit.get("reason", ""),
                        ]
                        for hit in item["manualReview"]
                    ],
                )
            )
            lines.append("")

    lines.extend(["## Notes", ""])
    for note in report.get("notes", []):
        lines.append(f"- {note}")
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--model-json", required=True, type=Path, help="Data Model JSON or model summary JSON")
    parser.add_argument("--openxml-json", required=True, type=Path, help="OpenXML workbook inspection JSON")
    parser.add_argument("--rename", action="append", type=parse_rename, default=[], help="Measure rename mapping, e.g. Revenue=Net Revenue")
    parser.add_argument("--delete", action="append", type=parse_delete, default=[], help="Measure name planned for deletion")
    parser.add_argument("--out-json", type=Path, help="Optional JSON rewrite plan path")
    parser.add_argument("--out-md", type=Path, help="Optional Markdown rewrite plan path")
    parser.add_argument("--fail-on-manual-review", action="store_true", help="Exit with 1 if any manual-review item is present")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    changes = list(args.rename) + list(args.delete)
    if not changes:
        raise SystemExit("At least one --rename Old=New or --delete Measure is required.")

    model = load_json(args.model_json.expanduser().resolve())
    openxml = load_json(args.openxml_json.expanduser().resolve())
    report = build_plan(model, openxml, changes)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(render_markdown(report), encoding="utf-8")

    print(
        f"Measure rewrite plan: {report['changeCount']} changes, "
        f"{report['daxRewriteCount']} DAX rewrites, "
        f"{report['cubeRewriteCount']} CUBE rewrites, "
        f"{report['downstreamFormulaImpactCount']} downstream impacts, "
        f"{report['manualReviewCount']} manual-review items"
    )
    return 1 if args.fail_on_manual_review and report["manualReviewCount"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
