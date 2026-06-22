#!/usr/bin/env python3
"""Create a sanitized workbook-backed visual QA fixture.

The fixture is intentionally generic and deterministic. It contains one clean
report sheet, one clipped-text report sheet, and one blank report sheet so the
visual QA report can prove both positive and risk-detection paths without
shipping customer workbooks.
"""

from __future__ import annotations

import argparse
import json
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape


def inline_cell(ref: str, text: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def number_cell(ref: str, value: float | int) -> str:
    return f'<c r="{ref}"><v>{value}</v></c>'


def formula_cell(ref: str, formula: str, value: float | int | str) -> str:
    return f'<c r="{ref}"><f>{escape(formula)}</f><v>{escape(str(value))}</v></c>'


def row_xml(row: int, cells: list[str], *, height: float | None = None) -> str:
    attrs = f' r="{row}"'
    if height is not None:
        attrs += f' ht="{height}" customHeight="1"'
    return f"<row{attrs}>{''.join(cells)}</row>"


def worksheet_xml(
    *,
    dimension: str,
    rows: list[str],
    cols: str = "",
    sheet_views: str = "",
    drawing_rel: str = "",
    merge_cells: str = "",
    page_setup: str = "",
) -> str:
    drawing = f'<drawing r:id="{drawing_rel}"/>' if drawing_rel else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  {sheet_views}
  {cols}
  <sheetData>
    {''.join(rows)}
  </sheetData>
  {merge_cells}
  {drawing}
  {page_setup}
</worksheet>
"""


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Report_OK" sheetId="2" r:id="rId2"/>
    <sheet name="Report_Clipped" sheetId="3" r:id="rId3"/>
    <sheet name="Report_Blank" sheetId="4" r:id="rId4"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="1">Report_OK!$A$1:$F$14</definedName>
  </definedNames>
  <calcPr calcId="191029" calcMode="auto"/>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

DATA_SHEET = worksheet_xml(
    dimension="A1:D5",
    cols='<cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="4" width="14" customWidth="1"/></cols>',
    rows=[
        row_xml(1, [inline_cell("A1", "Channel"), inline_cell("B1", "Amount"), inline_cell("C1", "Score"), inline_cell("D1", "Period")]),
        row_xml(2, [inline_cell("A2", "Search"), number_cell("B2", 120), number_cell("C2", 18.4), inline_cell("D2", "Q1")]),
        row_xml(3, [inline_cell("A3", "Social"), number_cell("B3", 80), number_cell("C3", 16.8), inline_cell("D3", "Q1")]),
        row_xml(4, [inline_cell("A4", "TV"), number_cell("B4", 240), number_cell("C4", 21.2), inline_cell("D4", "Q2")]),
        row_xml(5, [inline_cell("A5", "OOH"), number_cell("B5", 60), number_cell("C5", 14.9), inline_cell("D5", "Q2")]),
    ],
)

REPORT_OK = worksheet_xml(
    dimension="A1:F14",
    sheet_views='<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
    cols='<cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="6" width="14" customWidth="1"/></cols>',
    rows=[
        row_xml(1, [inline_cell("A1", "Visual QA Report Surface")], height=24),
        row_xml(3, [inline_cell("A3", "Metric"), inline_cell("B3", "Value"), inline_cell("C3", "Status")]),
        row_xml(4, [inline_cell("A4", "Total amount"), formula_cell("B4", "SUM(Data!B2:B5)", 500), inline_cell("C4", "OK")]),
        row_xml(5, [inline_cell("A5", "Average score"), formula_cell("B5", "AVERAGE(Data!C2:C5)", 17.825), inline_cell("C5", "OK")]),
        row_xml(6, [inline_cell("A6", "Channel count"), formula_cell("B6", "COUNTA(Data!A2:A5)", 4), inline_cell("C6", "OK")]),
        row_xml(8, [inline_cell("A8", "Notes"), inline_cell("B8", "Readable report area with a chart and stable dimensions.")]),
    ],
    drawing_rel="rId1",
    merge_cells='<mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>',
    page_setup='<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>',
)

CLIPPED_TEXT = (
    "This intentionally long label is placed in a narrow column to simulate a "
    "presentation surface where important text may be clipped or hard to read."
)

REPORT_CLIPPED = worksheet_xml(
    dimension="A1:B8",
    cols='<cols><col min="1" max="1" width="7" customWidth="1"/><col min="2" max="2" width="6" customWidth="1"/></cols>',
    rows=[
        row_xml(1, [inline_cell("A1", "Clipped text example")], height=14),
        row_xml(2, [inline_cell("A2", "Risk"), inline_cell("B2", CLIPPED_TEXT)], height=12),
        row_xml(4, [inline_cell("A4", "Metric"), number_cell("B4", 123)]),
    ],
)

REPORT_BLANK = worksheet_xml(
    dimension="A1:A1",
    rows=[
        row_xml(1, [inline_cell("A1", "")]),
    ],
)

REPORT_OK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"""

DRAWING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Visual QA Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""

DRAWING_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
"""

CHART = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Amount by channel</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/><c:tickLblPos val="nextTo"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/><c:tickLblPos val="nextTo"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>
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
  <TitlesOfParts><vt:vector size="4" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Report_OK</vt:lpstr><vt:lpstr>Report_Clipped</vt:lpstr><vt:lpstr>Report_Blank</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>
"""


def core(created_iso: str) -> str:
    safe_created = escape(created_iso)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Visual QA Fixture</dc:title>
  <dc:creator>Microsoft Excel BI Agent Pack</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">{safe_created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{safe_created}</dcterms:modified>
</cp:coreProperties>
"""


def write_fixture(workbook: Path) -> dict[str, object]:
    workbook.parent.mkdir(parents=True, exist_ok=True)
    created_iso = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": core(created_iso),
        "docProps/app.xml": APP,
        "xl/workbook.xml": WORKBOOK,
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
        "xl/worksheets/sheet1.xml": DATA_SHEET,
        "xl/worksheets/sheet2.xml": REPORT_OK,
        "xl/worksheets/_rels/sheet2.xml.rels": REPORT_OK_RELS,
        "xl/worksheets/sheet3.xml": REPORT_CLIPPED,
        "xl/worksheets/sheet4.xml": REPORT_BLANK,
        "xl/drawings/drawing1.xml": DRAWING,
        "xl/drawings/_rels/drawing1.xml.rels": DRAWING_RELS,
        "xl/charts/chart1.xml": CHART,
        "xl/styles.xml": STYLES,
    }
    with zipfile.ZipFile(workbook, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, text in parts.items():
            zf.writestr(name, text)
    return {
        "workbookPath": str(workbook),
        "createdAt": created_iso,
        "expected": {
            "reportSheets": ["Report_OK", "Report_Clipped", "Report_Blank"],
            "cleanReportSheet": "Report_OK",
            "expectedFindingCodes": ["long-text-narrow-column", "blank-report-sheet", "missing-print-area"],
            "expectedReadiness": "blocked-for-delivery",
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", required=True, type=Path, help="Output .xlsx workbook path")
    parser.add_argument("--out-json", type=Path, help="Optional fixture metadata JSON")
    args = parser.parse_args()

    result = write_fixture(args.workbook.expanduser().resolve())
    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
