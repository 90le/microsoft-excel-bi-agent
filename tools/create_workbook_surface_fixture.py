#!/usr/bin/env python3
"""Create a generic workbook surface fixture for OpenXML smoke tests.

The fixture is intentionally customer-data-free and structural. It contains
normal worksheet formulas, workbook-defined names, a table, and a chart part so
the release gate can verify workbook delivery surfaces without opening Excel or
shipping customer workbooks.
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
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
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
    <sheet name="Report" sheetId="2" r:id="rId2"/>
    <sheet name="Control" sheetId="3" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="CurrentPeriod">Control!$B$2</definedName>
    <definedName name="TotalAmount">Report!$B$2</definedName>
    <definedName name="ChannelData">Data!$A$1:$D$5</definedName>
    <definedName name="ReportMode">Control!$B$3</definedName>
  </definedNames>
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

DATA_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:D5"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Channel</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Amount</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Score</t></is></c>
      <c r="D1" t="inlineStr"><is><t>Period</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Search</t></is></c>
      <c r="B2"><v>120</v></c>
      <c r="C2"><v>18.4</v></c>
      <c r="D2" t="inlineStr"><is><t>Q1</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Social</t></is></c>
      <c r="B3"><v>80</v></c>
      <c r="C3"><v>16.8</v></c>
      <c r="D3" t="inlineStr"><is><t>Q1</t></is></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>TV</t></is></c>
      <c r="B4"><v>240</v></c>
      <c r="C4"><v>21.2</v></c>
      <c r="D4" t="inlineStr"><is><t>Q2</t></is></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>OOH</t></is></c>
      <c r="B5"><v>60</v></c>
      <c r="C5"><v>14.9</v></c>
      <c r="D5" t="inlineStr"><is><t>Q2</t></is></c>
    </row>
  </sheetData>
  <tableParts count="1">
    <tablePart r:id="rId1"/>
  </tableParts>
</worksheet>
"""

DATA_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>
"""

REPORT_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B8"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Workbook surface report</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Total amount</t></is></c>
      <c r="B2"><f>SUM(Data!B2:B5)</f><v>500</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Average score</t></is></c>
      <c r="B3"><f>AVERAGE(Data!C2:C5)</f><v>17.825</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>Channel count</t></is></c>
      <c r="B4"><f>COUNTA(Data!A2:A5)</f><v>4</v></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>High amount rows</t></is></c>
      <c r="B5"><f>COUNTIF(Data!B2:B5,&quot;&gt;=100&quot;)</f><v>2</v></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>Selected score</t></is></c>
      <c r="B6"><f>XLOOKUP(Control!B4,Data!A2:A5,Data!C2:C5,&quot;&quot;)</f><v>18.4</v></c>
    </row>
    <row r="8">
      <c r="A8" t="inlineStr"><is><t>Amount by period</t></is></c>
      <c r="B8"><f>SUMIFS(Data!B2:B5,Data!D2:D5,CurrentPeriod)</f><v>300</v></c>
    </row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>
"""

REPORT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"""

CONTROL_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B4"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Control surface</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Current period</t></is></c>
      <c r="B2" t="inlineStr"><is><t>Q2</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Mode</t></is></c>
      <c r="B3" t="inlineStr"><is><t>Review</t></is></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>Selected channel</t></is></c>
      <c r="B4" t="inlineStr"><is><t>Search</t></is></c>
    </row>
  </sheetData>
</worksheet>
"""

TABLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SurfaceData" displayName="SurfaceData" ref="A1:D5" totalsRowShown="0">
  <autoFilter ref="A1:D5"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Channel"/>
    <tableColumn id="2" name="Amount"/>
    <tableColumn id="3" name="Score"/>
    <tableColumn id="4" name="Period"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
"""

DRAWING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Surface Chart"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
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
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Amount by channel</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
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
  <TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Report</vt:lpstr><vt:lpstr>Control</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>
"""


def core(created_iso: str) -> str:
    safe_created = escape(created_iso)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Workbook Surface Fixture</dc:title>
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
        "xl/worksheets/_rels/sheet1.xml.rels": DATA_RELS,
        "xl/worksheets/sheet2.xml": REPORT_SHEET,
        "xl/worksheets/_rels/sheet2.xml.rels": REPORT_RELS,
        "xl/worksheets/sheet3.xml": CONTROL_SHEET,
        "xl/tables/table1.xml": TABLE,
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
            "sheets": ["Data", "Report", "Control"],
            "formulaCount": 6,
            "formulaFunctions": ["AVERAGE", "COUNTA", "COUNTIF", "SUM", "SUMIFS", "XLOOKUP"],
            "definedNames": ["ChannelData", "CurrentPeriod", "ReportMode", "TotalAmount"],
            "tableNames": ["SurfaceData"],
            "chartPartCount": 1,
            "drawingPartCount": 1,
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
