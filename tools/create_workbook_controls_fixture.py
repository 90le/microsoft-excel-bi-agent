#!/usr/bin/env python3
"""Create a generic workbook-controls OpenXML fixture for smoke tests."""

from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from typing import Any


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookProtection lockStructure="1" workbookPassword="ABCD"/>
  <sheets>
    <sheet name="Inputs" sheetId="1" r:id="rId1"/>
    <sheet name="Calc" sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="Audit" sheetId="3" state="veryHidden" r:id="rId3"/>
  </sheets>
  <calcPr calcId="191029" calcMode="auto"/>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

INPUTS_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B4"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Metric</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Budget</t></is></c><c r="B2"><v>100</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Mode</t></is></c><c r="B3" t="inlineStr"><is><t>Review</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Owner</t></is></c><c r="B4" t="inlineStr"><is><t>Team</t></is></c></row>
  </sheetData>
  <autoFilter ref="A1:B4"/>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" sqref="B3">
      <formula1>"Review,Final"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>
"""

CALC_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetProtection sheet="1" objects="1" scenarios="1"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Calc</t></is></c><c r="B1"><f>Inputs!B2*2</f><v>200</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Status</t></is></c><c r="B2" t="inlineStr"><is><t>Hidden support</t></is></c></row>
  </sheetData>
</worksheet>
"""

AUDIT_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B1"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Audit</t></is></c><c r="B1" t="inlineStr"><is><t>Very hidden support</t></is></c></row>
  </sheetData>
</worksheet>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>
"""


def create_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("xl/workbook.xml", WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        zf.writestr("xl/worksheets/sheet1.xml", INPUTS_SHEET)
        zf.writestr("xl/worksheets/sheet2.xml", CALC_SHEET)
        zf.writestr("xl/worksheets/sheet3.xml", AUDIT_SHEET)
        zf.writestr("xl/styles.xml", STYLES)


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", required=True, type=Path, help="Output .xlsx fixture path")
    parser.add_argument("--out-json", type=Path, help="Optional fixture manifest path")
    args = parser.parse_args()

    workbook = args.workbook.expanduser().resolve()
    create_workbook(workbook)
    manifest = {
        "workbook": str(workbook),
        "expected": {
            "sheetCount": 3,
            "hiddenSheetCount": 1,
            "veryHiddenSheetCount": 1,
            "protectedSheetCount": 1,
            "filteredSheetCount": 1,
            "frozenPaneSheetCount": 1,
            "dataValidationSheetCount": 1,
            "hasWorkbookProtection": True,
            "readiness": "review-required",
            "mediumFindingCount": 3,
            "lowFindingCount": 3,
            "requiredCodes": [
                "workbook-protection",
                "hidden-sheets",
                "very-hidden-sheets",
                "sheet-protection",
                "active-auto-filter",
                "data-validation-rules",
            ],
        },
    }
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), manifest)
    else:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
