#!/usr/bin/env python3
"""Create a generic Excel workbook fixture containing CUBE formulas.

The fixture is intentionally structural. It does not contain a real Power Pivot
model and does not claim that Excel can calculate the formulas. Its purpose is
to provide a safe, cross-platform input for OpenXML CUBE formula inspection and
dependency-report smoke tests.
"""

from __future__ import annotations

import argparse
import json
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="false"/>
  <sheets>
    <sheet name="Report" sheetId="1" r:id="rId1"/>
    <sheet name="ModelNotes" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="191029" calcMode="auto"/>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""

APP = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel BI Agent Pack Fixture</Application>
</Properties>
"""


def inline_string(address: str, value: str) -> str:
    return f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'


def number_cell(address: str, value: int | float) -> str:
    return f'<c r="{address}"><v>{value}</v></c>'


def formula_cell(address: str, formula: str, cached_value: str = "#N/A") -> str:
    return f'<c r="{address}"><f>{escape(formula)}</f><v>{escape(cached_value)}</v></c>'


def row(row_number: int, cells: list[str]) -> str:
    return f'<row r="{row_number}">' + "".join(cells) + "</row>"


def worksheet(rows: list[str], dimensions: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="{dimensions}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    {''.join(rows)}
  </sheetData>
</worksheet>
"""


def report_sheet() -> str:
    rows = [
        row(
            1,
            [
                inline_string("A1", "Quarter helper"),
                inline_string("B1", "Segment helper"),
                inline_string("C1", "Revenue"),
                inline_string("D1", "Awareness"),
                inline_string("E1", "Missing measure"),
                inline_string("F1", "Hard-coded latest marker"),
                inline_string("G1", "Dynamic member string"),
            ],
        ),
        row(
            2,
            [
                formula_cell("A2", 'CUBEMEMBER("ThisWorkbookDataModel","[Calendar].[Quarter].[All].[2026Q1]","2026Q1")'),
                formula_cell("B2", 'CUBEMEMBER("ThisWorkbookDataModel","[Segment].[Name].[All].[Online]","Online")'),
                formula_cell("C2", 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue]",$A$2,$B$2)', "120"),
                formula_cell("D2", 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Awareness]",$A$2,$B$2)', "42"),
                formula_cell("E2", 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Missing Measure]",$A$2)', "#VALUE!"),
                formula_cell("F2", 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue]","[Calendar].[Latest].[All].[new]")', "121"),
                formula_cell("G2", 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue]","[Calendar].[Quarter].[All].["&$A$5&"]")', "#GETTING_DATA"),
            ],
        ),
        row(5, [inline_string("A5", "2026Q2")]),
    ]
    return worksheet(rows, "A1:G5")


def notes_sheet() -> str:
    rows = [
        row(1, [inline_string("A1", "Fixture purpose"), inline_string("B1", "Structural CUBE formula parser smoke test")]),
        row(2, [inline_string("A2", "Known measures"), inline_string("B2", "Revenue, Awareness")]),
        row(3, [inline_string("A3", "No live model"), inline_string("B3", "OpenXML-only; do not use to validate calculation")]),
    ]
    return worksheet(rows, "A1:B3")


def model_summary() -> dict[str, object]:
    return {
        "workbookPath": "generic_cube_formula_fixture.xlsx",
        "modelAvailable": True,
        "tableCount": 3,
        "relationshipCount": 2,
        "measureCount": 2,
        "connections": [
            {"name": "ThisWorkbookDataModel", "type": "MODEL", "description": "Structural fixture connection"}
        ],
        "measures": [
            {
                "name": "Revenue",
                "associatedTable": {"name": "Fact"},
                "formula": "SUM(Fact[Amount])",
            },
            {
                "name": "Awareness",
                "associatedTable": {"name": "Fact"},
                "formula": "AVERAGE(Fact[AwarenessScore])",
            },
        ],
        "tables": [
            {"name": "Fact", "sourceName": "FixtureFact", "recordCount": 4, "columnCount": 5},
            {"name": "Calendar", "sourceName": "FixtureCalendar", "recordCount": 2, "columnCount": 2},
            {"name": "Segment", "sourceName": "FixtureSegment", "recordCount": 2, "columnCount": 2},
        ],
        "relationships": [
            {
                "foreignKeyTable": "Fact",
                "foreignKeyColumn": "QuarterKey",
                "primaryKeyTable": "Calendar",
                "primaryKeyColumn": "QuarterKey",
                "active": True,
            },
            {
                "foreignKeyTable": "Fact",
                "foreignKeyColumn": "SegmentKey",
                "primaryKeyTable": "Segment",
                "primaryKeyColumn": "SegmentKey",
                "active": True,
            },
        ],
        "fixtureNotes": [
            "This is a structural model summary for CUBE dependency smoke tests.",
            "It is not a live Power Pivot model export.",
        ],
    }


def core_props() -> str:
    stamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Generic CUBE Formula Fixture</dc:title>
  <dc:creator>microsoft-excel-bi-agent-pack</dc:creator>
  <cp:lastModifiedBy>microsoft-excel-bi-agent-pack</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{stamp}</dcterms:modified>
</cp:coreProperties>
"""


def create_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("docProps/core.xml", core_props())
        zf.writestr("docProps/app.xml", APP)
        zf.writestr("xl/workbook.xml", WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        zf.writestr("xl/styles.xml", STYLES)
        zf.writestr("xl/worksheets/sheet1.xml", report_sheet())
        zf.writestr("xl/worksheets/sheet2.xml", notes_sheet())


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", required=True, type=Path, help="Output .xlsx fixture path")
    parser.add_argument("--model-json", type=Path, help="Optional output JSON with known model measures")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workbook = args.workbook.expanduser().resolve()
    create_workbook(workbook)
    print(f"Wrote CUBE formula fixture workbook: {workbook}")
    if args.model_json:
        model_path = args.model_json.expanduser().resolve()
        model_path.parent.mkdir(parents=True, exist_ok=True)
        model_path.write_text(json.dumps(model_summary(), ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Wrote CUBE formula fixture model summary: {model_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
