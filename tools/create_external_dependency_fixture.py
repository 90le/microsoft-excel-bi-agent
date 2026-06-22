#!/usr/bin/env python3
"""Create a generic workbook fixture with external-dependency structures.

The fixture is intentionally structural. It does not connect to any real file,
endpoint, Power Query source, or Power Pivot model. Its purpose is to provide a
safe cross-platform workbook for testing OpenXML detection of connections,
external links, formulas, and mashup-like/custom XML parts.
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
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
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
    <sheet name="ExternalAudit" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="ExternalInput">'[safe-source.xlsx]Data'!$A$1</definedName>
  </definedNames>
  <calcPr calcId="191029" calcMode="manual"/>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>
"""

SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>External dependency fixture</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Safe structural workbook</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>External formula</t></is></c>
      <c r="B2"><f>'[safe-source.xlsx]Data'!$A$1</f><v>0</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Connection name</t></is></c>
      <c r="B3" t="inlineStr"><is><t>SafePowerQueryLikeConnection</t></is></c>
    </row>
  </sheetData>
</worksheet>
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

CONNECTIONS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <connection id="1" name="SafePowerQueryLikeConnection" type="5" refreshedVersion="7" background="0" saveData="1">
    <dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=SafeQuery;Extended Properties=&quot;&quot;" command="SELECT * FROM [SafeQuery]" commandType="2"/>
  </connection>
  <connection id="2" name="CredentialLikeConnection" type="5" refreshedVersion="7" background="0" saveData="0">
    <dbPr connection="Provider=SQLOLEDB;Data Source=warehouse.example.internal;Initial Catalog=FinanceModel;User ID=report_user;Password=REDACTED" command="SELECT Region, Amount FROM dbo.FactSales" commandType="2"/>
  </connection>
</connections>
"""

EXTERNAL_LINK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames>
      <sheetName val="Data"/>
    </sheetNames>
  </externalBook>
</externalLink>
"""

EXTERNAL_LINK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="safe-source.xlsx" TargetMode="External"/>
</Relationships>
"""

CUSTOM_XML = """<?xml version="1.0" encoding="UTF-8"?>
<safeWorkbookMetadata>
  <purpose>OpenXML external dependency detection fixture</purpose>
  <containsRealConnection>false</containsRealConnection>
</safeWorkbookMetadata>
"""

APP = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel BI Agent Pack Fixture</Application>
</Properties>
"""


def core(created_iso: str) -> str:
    safe_created = escape(created_iso)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>External Dependency Fixture</dc:title>
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
        "xl/worksheets/sheet1.xml": SHEET,
        "xl/styles.xml": STYLES,
        "xl/connections.xml": CONNECTIONS,
        "xl/externalLinks/externalLink1.xml": EXTERNAL_LINK,
        "xl/externalLinks/_rels/externalLink1.xml.rels": EXTERNAL_LINK_RELS,
        "customXml/item1.xml": CUSTOM_XML,
    }
    with zipfile.ZipFile(workbook, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, text in parts.items():
            zf.writestr(name, text)

    return {
        "workbookPath": str(workbook),
        "createdAt": created_iso,
        "expected": {
            "connectionCount": 2,
            "credentialLikeConnectionCount": 1,
            "externalLinkCount": 1,
            "formulaCount": 1,
            "hasMashupLikeParts": True,
            "connectionName": "SafePowerQueryLikeConnection",
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
