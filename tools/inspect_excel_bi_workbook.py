#!/usr/bin/env python3
"""Inspect Excel BI workbook structure with OpenXML only.

This script is intentionally read-only and cross-platform. It does not execute
VBA, refresh Power Query, evaluate CUBE formulas, or decode Power Pivot models.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "officeRel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

CUBE_FUNCTION_RE = re.compile(
    r"\b(CUBEVALUE|CUBEMEMBER|CUBESET|CUBERANKEDMEMBER|CUBEKPIMEMBER|CUBEMEMBERPROPERTY)\s*\(",
    re.IGNORECASE,
)

BI_FORMULA_RE = re.compile(
    r"\b(CUBEVALUE|CUBEMEMBER|CUBESET|CUBERANKEDMEMBER|CUBEKPIMEMBER|"
    r"CUBEMEMBERPROPERTY|GETPIVOTDATA)\s*\(",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class SheetInfo:
    name: str
    sheet_id: str
    rel_id: str
    target: str
    part: str
    state: str


def qname(local: str, namespace: str = "main") -> str:
    return f"{{{NS[namespace]}}}{local}"


def parse_xml(zf: zipfile.ZipFile, part: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(part))
    except KeyError:
        return None
    except ET.ParseError:
        return None


def read_relationships(zf: zipfile.ZipFile, part: str) -> dict[str, dict[str, str]]:
    root = parse_xml(zf, part)
    if root is None:
        return {}
    rels: dict[str, dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id", "")
        if not rel_id:
            continue
        rels[rel_id] = {
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
            "targetMode": rel.attrib.get("TargetMode", ""),
        }
    return rels


def normalize_target(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_path = Path(base).parent
    parts: list[str] = []
    for part in (base_path / target).as_posix().split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def get_workbook_sheets(zf: zipfile.ZipFile) -> list[SheetInfo]:
    workbook_part = "xl/workbook.xml"
    workbook = parse_xml(zf, workbook_part)
    rels = read_relationships(zf, "xl/_rels/workbook.xml.rels")
    sheets: list[SheetInfo] = []
    if workbook is None:
        return sheets
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rel_id = sheet.attrib.get(f"{{{NS['officeRel']}}}id", "")
        target = rels.get(rel_id, {}).get("target", "")
        part = normalize_target(workbook_part, target) if target else ""
        sheets.append(
            SheetInfo(
                name=sheet.attrib.get("name", ""),
                sheet_id=sheet.attrib.get("sheetId", ""),
                rel_id=rel_id,
                target=target,
                part=part,
                state=sheet.attrib.get("state", "visible") or "visible",
            )
        )
    return sheets


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    root = parse_xml(zf, "xl/sharedStrings.xml")
    if root is None:
        return []
    values: list[str] = []
    for si in root.findall("main:si", NS):
        pieces = [node.text or "" for node in si.findall(".//main:t", NS)]
        values.append("".join(pieces))
    return values


def text_of_cell(cell: ET.Element, shared_strings: list[str]) -> str | None:
    cell_type = cell.attrib.get("t")
    value = cell.find("main:v", NS)
    if value is None or value.text is None:
        return None
    if cell_type == "s":
        try:
            index = int(value.text)
            return shared_strings[index]
        except (ValueError, IndexError):
            return None
    return value.text


def inspect_formulas(
    zf: zipfile.ZipFile,
    sheets: list[SheetInfo],
    shared_strings: list[str],
) -> dict[str, Any]:
    formulas: list[dict[str, Any]] = []
    cube_formulas: list[dict[str, Any]] = []
    formula_counts: dict[str, int] = {}
    sheet_formula_counts: dict[str, int] = {}

    for sheet in sheets:
        if not sheet.part:
            continue
        root = parse_xml(zf, sheet.part)
        if root is None:
            continue
        for cell in root.findall(".//main:c", NS):
            formula = cell.find("main:f", NS)
            if formula is None or formula.text is None:
                continue
            ref = cell.attrib.get("r", "")
            formula_text = formula.text.strip()
            formula_type = "bi" if BI_FORMULA_RE.search(formula_text) else "formula"
            record = {
                "sheet": sheet.name,
                "cell": ref,
                "formula": formula_text,
                "formulaType": formula_type,
                "cachedValue": text_of_cell(cell, shared_strings),
            }
            formulas.append(record)
            sheet_formula_counts[sheet.name] = sheet_formula_counts.get(sheet.name, 0) + 1
            function = formula_text.split("(", 1)[0].split(".")[-1].upper()
            formula_counts[function] = formula_counts.get(function, 0) + 1
            if CUBE_FUNCTION_RE.search(formula_text):
                cube_formulas.append(record)

    return {
        "formulas": formulas,
        "totalFormulaCount": len(formulas),
        "formulaCountsBySheet": sheet_formula_counts,
        "formulaFunctionCounts": dict(sorted(formula_counts.items())),
        "cubeFormulaCount": len(cube_formulas),
        "cubeFormulas": cube_formulas,
        "biFormulas": [item for item in formulas if item["formulaType"] == "bi"],
    }


def inspect_connections(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    root = parse_xml(zf, "xl/connections.xml")
    if root is None:
        return []
    connections: list[dict[str, Any]] = []
    for connection in root.findall("main:connection", NS):
        item: dict[str, Any] = dict(connection.attrib)
        db_pr = connection.find("main:dbPr", NS)
        if db_pr is not None:
            item["dbPr"] = dict(db_pr.attrib)
        olap_pr = connection.find("main:olapPr", NS)
        if olap_pr is not None:
            item["olapPr"] = dict(olap_pr.attrib)
        text_pr = connection.find("main:textPr", NS)
        if text_pr is not None:
            item["textPr"] = dict(text_pr.attrib)
        connections.append(item)
    return connections


def inspect_tables(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    for part in sorted(name for name in zf.namelist() if name.startswith("xl/tables/table") and name.endswith(".xml")):
        root = parse_xml(zf, part)
        if root is None:
            continue
        columns = []
        for column in root.findall("main:tableColumns/main:tableColumn", NS):
            columns.append({"id": column.attrib.get("id", ""), "name": column.attrib.get("name", "")})
        tables.append(
            {
                "part": part,
                "name": root.attrib.get("name", ""),
                "displayName": root.attrib.get("displayName", ""),
                "ref": root.attrib.get("ref", ""),
                "columns": columns,
            }
        )
    return tables


def inspect_pivot_caches(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    caches: list[dict[str, Any]] = []
    for part in sorted(
        name
        for name in zf.namelist()
        if name.startswith("xl/pivotCache/pivotCacheDefinition") and name.endswith(".xml")
    ):
        root = parse_xml(zf, part)
        if root is None:
            continue
        cache_source = root.find("main:cacheSource", NS)
        worksheet_source = root.find("main:cacheSource/main:worksheetSource", NS)
        fields = root.findall("main:cacheFields/main:cacheField", NS)
        caches.append(
            {
                "part": part,
                "recordCount": root.attrib.get("recordCount", ""),
                "refreshOnLoad": root.attrib.get("refreshOnLoad", ""),
                "cacheSourceType": cache_source.attrib.get("type", "") if cache_source is not None else "",
                "worksheetSource": dict(worksheet_source.attrib) if worksheet_source is not None else {},
                "fieldCount": len(fields),
                "fields": [field.attrib.get("name", "") for field in fields],
            }
        )
    return caches


def inspect_workbook_defined_names(zf: zipfile.ZipFile) -> list[dict[str, str]]:
    workbook = parse_xml(zf, "xl/workbook.xml")
    if workbook is None:
        return []
    names: list[dict[str, str]] = []
    for defined_name in workbook.findall("main:definedNames/main:definedName", NS):
        names.append(
            {
                "name": defined_name.attrib.get("name", ""),
                "localSheetId": defined_name.attrib.get("localSheetId", ""),
                "hidden": defined_name.attrib.get("hidden", ""),
                "refersTo": defined_name.text or "",
            }
        )
    return names


def inspect_workbook_protection(zf: zipfile.ZipFile) -> dict[str, Any]:
    workbook = parse_xml(zf, "xl/workbook.xml")
    if workbook is None:
        return {"hasWorkbookProtection": False, "workbookProtection": {}}
    protection = workbook.find("main:workbookProtection", NS)
    if protection is None:
        return {"hasWorkbookProtection": False, "workbookProtection": {}}
    return {"hasWorkbookProtection": True, "workbookProtection": dict(protection.attrib)}


def inspect_sheet_controls(zf: zipfile.ZipFile, sheets: list[SheetInfo]) -> list[dict[str, Any]]:
    controls: list[dict[str, Any]] = []
    for sheet in sheets:
        root = parse_xml(zf, sheet.part) if sheet.part else None
        sheet_protection: dict[str, str] = {}
        auto_filter_ref = ""
        panes: list[dict[str, str]] = []
        data_validation_count = 0
        if root is not None:
            protection = root.find("main:sheetProtection", NS)
            if protection is not None:
                sheet_protection = dict(protection.attrib)
            auto_filter = root.find("main:autoFilter", NS)
            if auto_filter is not None:
                auto_filter_ref = auto_filter.attrib.get("ref", "")
            for pane in root.findall("main:sheetViews/main:sheetView/main:pane", NS):
                panes.append(dict(pane.attrib))
            data_validations = root.find("main:dataValidations", NS)
            if data_validations is not None:
                try:
                    data_validation_count = int(data_validations.attrib.get("count", "0") or "0")
                except ValueError:
                    data_validation_count = len(data_validations.findall("main:dataValidation", NS))
        controls.append(
            {
                "name": sheet.name,
                "state": sheet.state,
                "hasSheetProtection": bool(sheet_protection),
                "sheetProtection": sheet_protection,
                "hasAutoFilter": bool(auto_filter_ref),
                "autoFilterRef": auto_filter_ref,
                "hasFrozenPane": any(pane.get("state") in {"frozen", "frozenSplit"} for pane in panes),
                "panes": panes,
                "dataValidationCount": data_validation_count,
            }
        )
    return controls


def inspect_parts(zf: zipfile.ZipFile) -> dict[str, Any]:
    names = zf.namelist()
    power_pivot_parts = [
        name for name in names
        if name.startswith("xl/model/")
        or "item.data" in name.lower()
        or "datamodel" in name.lower()
    ]
    mashup_parts = [
        name for name in names
        if "mashup" in name.lower()
        or "datamashup" in name.lower()
        or name.startswith("customXml/")
    ]
    vba_parts = [name for name in names if name.endswith("vbaProject.bin")]
    external_links = [name for name in names if name.startswith("xl/externalLinks/")]
    slicers = [name for name in names if "slicer" in name.lower()]
    timelines = [name for name in names if "timeline" in name.lower()]
    charts = [name for name in names if name.startswith("xl/charts/") and name.endswith(".xml")]
    drawings = [name for name in names if name.startswith("xl/drawings/") and name.endswith(".xml")]
    return {
        "hasVbaProject": bool(vba_parts),
        "vbaParts": vba_parts,
        "hasPowerPivotLikeParts": bool(power_pivot_parts),
        "powerPivotLikeParts": power_pivot_parts,
        "hasMashupLikeParts": bool(mashup_parts),
        "mashupLikeParts": mashup_parts,
        "externalLinks": external_links,
        "slicerParts": slicers,
        "timelineParts": timelines,
        "chartParts": charts,
        "drawingParts": drawings,
    }


def inspect_workbook(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path) as zf:
        sheets = get_workbook_sheets(zf)
        shared_strings = load_shared_strings(zf)
        formula_info = inspect_formulas(zf, sheets, shared_strings)
        connections = inspect_connections(zf)
        tables = inspect_tables(zf)
        pivots = inspect_pivot_caches(zf)
        parts = inspect_parts(zf)
        defined_names = inspect_workbook_defined_names(zf)
        protection = inspect_workbook_protection(zf)
        sheet_controls = inspect_sheet_controls(zf, sheets)
        return {
            "workbookPath": str(path),
            "fileType": path.suffix.lower(),
            "sheets": [sheet.__dict__ for sheet in sheets],
            "sheetControls": sheet_controls,
            **protection,
            "connections": connections,
            "tables": tables,
            "pivotCaches": pivots,
            "definedNames": defined_names,
            **formula_info,
            **parts,
            "limitations": [
                "OpenXML inspection does not execute formulas or refresh queries.",
                "Power Pivot DAX model metadata is not fully decoded by this script.",
                "Use Windows Excel COM for VBA execution, Power Query refresh, and Data Model validation.",
            ],
        }


def render_markdown(result: dict[str, Any]) -> str:
    lines = [
        "# Excel BI Workbook Inspection",
        "",
        f"- workbook: `{result['workbookPath']}`",
        f"- file type: `{result['fileType']}`",
        f"- sheets: {len(result['sheets'])}",
        f"- connections: {len(result['connections'])}",
        f"- tables: {len(result['tables'])}",
        f"- pivot caches: {len(result['pivotCaches'])}",
        f"- chart parts: {len(result['chartParts'])}",
        f"- formulas: {result['totalFormulaCount']}",
        f"- CUBE formulas: {result['cubeFormulaCount']}",
        f"- hidden sheets: {sum(1 for item in result.get('sheetControls', []) if item.get('state') != 'visible')}",
        f"- protected sheets: {sum(1 for item in result.get('sheetControls', []) if item.get('hasSheetProtection'))}",
        f"- workbook protection: {result.get('hasWorkbookProtection', False)}",
        f"- has VBA project: {result['hasVbaProject']}",
        f"- has Power Pivot-like parts: {result['hasPowerPivotLikeParts']}",
        f"- has Mashup-like parts: {result['hasMashupLikeParts']}",
        "",
    ]
    if result["connections"]:
        lines.extend(["## Connections", ""])
        for conn in result["connections"]:
            name = conn.get("name", "")
            ctype = conn.get("type", "")
            db = conn.get("dbPr", {})
            connection_string = db.get("connection", "")
            lines.append(f"- `{name}` type `{ctype}` {connection_string}")
        lines.append("")
    if result["cubeFormulas"]:
        lines.extend(["## CUBE Formulas", ""])
        for item in result["cubeFormulas"][:50]:
            lines.append(f"- `{item['sheet']}!{item['cell']}`: `{item['formula']}`")
        if len(result["cubeFormulas"]) > 50:
            lines.append(f"- ... {len(result['cubeFormulas']) - 50} more")
        lines.append("")
    if result["pivotCaches"]:
        lines.extend(["## Pivot Caches", ""])
        for cache in result["pivotCaches"]:
            lines.append(
                f"- `{cache['part']}` source `{cache['cacheSourceType']}` fields {cache['fieldCount']}"
            )
        lines.append("")
    lines.extend(["## Limitations", ""])
    for item in result["limitations"]:
        lines.append(f"- {item}")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="Path to .xlsx/.xlsm/.xlsb OpenXML workbook")
    parser.add_argument("--out-json", type=Path, help="Write JSON report")
    parser.add_argument("--markdown", action="store_true", help="Print Markdown summary")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workbook = args.workbook.expanduser().resolve()
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 2
    try:
        result = inspect_workbook(workbook)
    except zipfile.BadZipFile:
        print(f"Not an OpenXML workbook: {workbook}", file=sys.stderr)
        return 2

    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.markdown:
        print(render_markdown(result))
    elif args.out_json:
        print(f"Wrote JSON report: {args.out_json}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
