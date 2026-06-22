#!/usr/bin/env python3
"""Inspect Power Query-related OpenXML package parts without Excel.

This detects workbook connections, query tables, external links, and package
parts that look related to Power Query/Mashup. It does not fully decode Excel's
mashup binary and cannot refresh or validate queries.
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
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

KEYWORDS = [
    b"Power Query",
    b"Microsoft.Mashup",
    b"DataMashup",
    b"Section1",
    b"shared ",
    b"let",
]


def read_xml(zf: zipfile.ZipFile, name: str):
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None
    except ET.ParseError as exc:
        return {"parse_error": str(exc)}


def inspect_connections(zf: zipfile.ZipFile):
    root = read_xml(zf, "xl/connections.xml")
    if root is None:
        return []
    if isinstance(root, dict):
        return [{"parseError": root["parse_error"]}]
    results = []
    for conn in root.findall(".//main:connection", NS):
        item = dict(conn.attrib)
        dbpr = conn.find("main:dbPr", NS)
        if dbpr is not None:
            item["dbPr"] = dict(dbpr.attrib)
        oledb = conn.find("main:oledbPr", NS)
        if oledb is not None:
            item["oledbPr"] = dict(oledb.attrib)
        results.append(item)
    return results


def inspect_query_tables(zf: zipfile.ZipFile):
    names = sorted(n for n in zf.namelist() if "/queryTables/" in n and n.endswith(".xml"))
    results = []
    for name in names:
        root = read_xml(zf, name)
        item = {"part": name}
        if root is not None and not isinstance(root, dict):
            item.update(root.attrib)
        elif isinstance(root, dict):
            item["parseError"] = root["parse_error"]
        results.append(item)
    return results


def find_mashup_like_parts(zf: zipfile.ZipFile):
    results = []
    for name in sorted(zf.namelist()):
        lower = name.lower()
        candidate_by_name = (
            "mashup" in lower
            or "powerquery" in lower
            or lower.startswith("customxml/")
            or "connections" in lower
        )
        if not candidate_by_name:
            continue
        try:
            data = zf.read(name)
        except KeyError:
            continue
        hits = [kw.decode("ascii", errors="ignore") for kw in KEYWORDS if kw in data]
        if hits or "mashup" in lower or "powerquery" in lower:
            results.append(
                {
                    "part": name,
                    "size": len(data),
                    "keywordHits": hits,
                    "isXml": data.lstrip().startswith(b"<"),
                }
            )
    return results


def rels_for(zf: zipfile.ZipFile, part_name: str):
    directory = posixpath.dirname(part_name)
    base = posixpath.basename(part_name)
    rels_name = posixpath.join(directory, "_rels", base + ".rels")
    root = read_xml(zf, rels_name)
    if root is None or isinstance(root, dict):
        return []
    rels = []
    for rel in root.findall("pkgrel:Relationship", NS):
        rels.append(dict(rel.attrib))
    return rels


def inspect(path: Path):
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return {
            "workbookPath": str(path),
            "error": "OpenXML inspection supports .xlsx/.xlsm/.xltx/.xltm only. Use Excel COM or LibreOffice for other formats.",
        }

    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        workbook_rels = rels_for(zf, "xl/workbook.xml")
        external_links = sorted(n for n in names if n.startswith("xl/externalLinks/") and n.endswith(".xml"))
        custom_xml = sorted(n for n in names if n.startswith("customXml/") and not n.endswith(".rels"))
        return {
            "workbookPath": str(path),
            "fileType": path.suffix.lower(),
            "hasVbaProjectBin": "xl/vbaProject.bin" in names,
            "hasConnectionsXml": "xl/connections.xml" in names,
            "connections": inspect_connections(zf),
            "queryTables": inspect_query_tables(zf),
            "externalLinks": external_links,
            "customXmlParts": custom_xml,
            "workbookRelationships": workbook_rels,
            "mashupLikeParts": find_mashup_like_parts(zf),
            "limitations": [
                "This script does not fully decode Excel DataMashup binaries.",
                "This script cannot refresh Power Query.",
                "Use Windows Excel COM export for exact Workbook.Queries formulas.",
            ],
        }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook")
    parser.add_argument("--out-json")
    args = parser.parse_args()

    result = inspect(Path(args.workbook).expanduser().resolve())
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
