#!/usr/bin/env python3
"""Cross-platform OpenXML workbook inventory for .xlsx/.xlsm files.

This script intentionally avoids Excel automation. It can inspect workbook
structure, formulas, connections, links, drawings, and macro binary presence,
but it cannot export VBE modules, compile VBA, run macros, refresh Power Query,
or execute Solver.
"""

from __future__ import annotations

import argparse
import json
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def read_xml(zf: zipfile.ZipFile, name: str):
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None
    except ET.ParseError as exc:
        return {"parse_error": str(exc)}


def rels_for(zf: zipfile.ZipFile, part_name: str) -> dict[str, dict[str, str]]:
    directory = posixpath.dirname(part_name)
    base = posixpath.basename(part_name)
    rels_name = posixpath.join(directory, "_rels", base + ".rels")
    root = read_xml(zf, rels_name)
    if root is None or isinstance(root, dict):
        return {}
    rels = {}
    for rel in root.findall("pkgrel:Relationship", NS):
        rels[rel.attrib.get("Id", "")] = {
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
            "targetMode": rel.attrib.get("TargetMode", ""),
        }
    return rels


def normalize_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target))


def formula_count(root) -> int:
    if root is None or isinstance(root, dict):
        return 0
    return len(root.findall(".//main:f", NS))


def inspect_workbook(path: Path) -> dict:
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return {
            "workbookPath": str(path),
            "error": "OpenXML inspection supports .xlsx/.xlsm/.xltx/.xltm only. Use Excel COM for .xls/.xlsb.",
        }

    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        workbook_root = read_xml(zf, "xl/workbook.xml")
        workbook_rels = rels_for(zf, "xl/workbook.xml")

        sheets = []
        defined_names = []

        if workbook_root is not None and not isinstance(workbook_root, dict):
            for sheet in workbook_root.findall("main:sheets/main:sheet", NS):
                rel_id = sheet.attrib.get(f"{{{NS['rel']}}}id", "")
                rel = workbook_rels.get(rel_id, {})
                target = rel.get("target", "")
                sheet_part = normalize_part("xl/workbook.xml", target) if target else ""
                sheet_root = read_xml(zf, sheet_part) if sheet_part else None
                sheet_rels = rels_for(zf, sheet_part) if sheet_part else {}
                drawing_count = sum(1 for r in sheet_rels.values() if "/drawing" in r.get("type", ""))

                dimension = ""
                rows = 0
                cols = 0
                if sheet_root is not None and not isinstance(sheet_root, dict):
                    dim = sheet_root.find("main:dimension", NS)
                    dimension = dim.attrib.get("ref", "") if dim is not None else ""
                    rows = len(sheet_root.findall(".//main:sheetData/main:row", NS))
                    col_ids = set()
                    for cell in sheet_root.findall(".//main:sheetData/main:row/main:c", NS):
                        ref = cell.attrib.get("r", "")
                        col = "".join(ch for ch in ref if ch.isalpha())
                        if col:
                            col_ids.add(col)
                    cols = len(col_ids)

                sheets.append(
                    {
                        "name": sheet.attrib.get("name", ""),
                        "sheetId": sheet.attrib.get("sheetId", ""),
                        "state": sheet.attrib.get("state", "visible"),
                        "part": sheet_part,
                        "dimension": dimension,
                        "rowCountWithCells": rows,
                        "columnCountWithCells": cols,
                        "formulaCount": formula_count(sheet_root),
                        "drawingRelationshipCount": drawing_count,
                    }
                )

            for dn in workbook_root.findall("main:definedNames/main:definedName", NS):
                defined_names.append(
                    {
                        "name": dn.attrib.get("name", ""),
                        "localSheetId": dn.attrib.get("localSheetId", ""),
                        "hidden": dn.attrib.get("hidden", ""),
                        "text": dn.text or "",
                    }
                )

        connections = []
        conn_root = read_xml(zf, "xl/connections.xml")
        if conn_root is not None and not isinstance(conn_root, dict):
            for conn in conn_root.findall(".//main:connection", NS):
                connections.append(
                    {
                        "id": conn.attrib.get("id", ""),
                        "name": conn.attrib.get("name", ""),
                        "type": conn.attrib.get("type", ""),
                        "refreshedVersion": conn.attrib.get("refreshedVersion", ""),
                    }
                )

        external_links = sorted(n for n in names if n.startswith("xl/externalLinks/") and n.endswith(".xml"))
        query_tables = sorted(n for n in names if "/queryTables/" in n and n.endswith(".xml"))
        table_parts = sorted(n for n in names if n.startswith("xl/tables/") and n.endswith(".xml"))
        drawing_parts = sorted(n for n in names if n.startswith("xl/drawings/") and n.endswith(".xml"))

        return {
            "workbookPath": str(path),
            "fileType": path.suffix.lower(),
            "hasVbaProjectBin": "xl/vbaProject.bin" in names,
            "sheetCount": len(sheets),
            "worksheets": sheets,
            "definedNames": defined_names,
            "connections": connections,
            "externalLinks": external_links,
            "queryTables": query_tables,
            "tables": table_parts,
            "drawings": drawing_parts,
            "limitations": [
                "OpenXML inspection does not compile or run VBA.",
                "OpenXML inspection does not refresh Power Query, data model, cube formulas, Solver, or external links.",
                "Use Windows desktop Excel COM for VBE import/export and macro validation.",
            ],
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect .xlsx/.xlsm workbook structure without Excel.")
    parser.add_argument("workbook", help="Path to .xlsx/.xlsm workbook")
    parser.add_argument("--out-json", help="Optional JSON output path")
    args = parser.parse_args()

    path = Path(args.workbook).expanduser().resolve()
    result = inspect_workbook(path)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out_json:
        out = Path(args.out_json).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0 if "error" not in result else 1


if __name__ == "__main__":
    sys.exit(main())
