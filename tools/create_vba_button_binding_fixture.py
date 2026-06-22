#!/usr/bin/env python3
"""Create generic fixture JSON for VBA button binding report tests."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def workbook_inventory() -> dict[str, Any]:
    return {
        "workbookPath": "generic_button_fixture.xlsm",
        "name": "generic_button_fixture.xlsm",
        "fileFormat": 52,
        "hasVBProject": True,
        "worksheets": [
            {
                "name": "Control",
                "codeName": "Sheet1",
                "visible": "Visible",
                "usedAddress": "A1:D6",
                "rows": 6,
                "columns": 4,
                "formulaCount": 0,
                "shapes": [
                    {
                        "name": "btnRefresh",
                        "type": 1,
                        "onAction": "'generic_button_fixture.xlsm'!RefreshReport",
                        "text": "Refresh Report",
                    },
                    {
                        "name": "btnExport",
                        "type": 1,
                        "onAction": "modMain.ExportReport",
                        "text": "Export Report",
                    },
                    {
                        "name": "btnMissing",
                        "type": 1,
                        "onAction": "MissingMacro",
                        "text": "Missing Macro",
                    },
                    {
                        "name": "decorativeShape",
                        "type": 1,
                        "onAction": "",
                        "text": "Decoration",
                    },
                ],
            }
        ],
        "names": [],
        "links": [],
        "connections": [],
        "queries": [],
        "queryAccessError": None,
        "vbaComponents": [{"name": "modMain", "type": 1, "lineCount": 12}],
        "vbaAccessError": None,
    }


def vba_lint() -> dict[str, Any]:
    public_entries = [
        {"module": "modMain.bas", "name": "RefreshReport", "kind": "SUB", "visibility": "Public", "line": 3},
        {"module": "modMain.bas", "name": "ExportReport", "kind": "SUB", "visibility": "Public", "line": 8},
    ]
    return {
        "source": "src/vba",
        "moduleCount": 1,
        "procedureCount": 2,
        "publicEntryCount": 2,
        "publicEntries": public_entries,
        "modules": [
            {
                "path": "src/vba/modMain.bas",
                "file": "modMain.bas",
                "extension": ".bas",
                "encoding": "utf-8-sig",
                "declaredName": "modMain",
                "hasOptionExplicit": True,
                "procedures": public_entries,
                "errors": [],
                "warnings": [],
            }
        ],
        "errors": [],
        "warnings": [],
    }


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def create_fixture(out_dir: Path) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    workbook_json = out_dir / "workbook_inventory.json"
    vba_json = out_dir / "vba_lint.json"
    manifest = {
        "workbookInventoryJson": str(workbook_json),
        "vbaLintJson": str(vba_json),
        "expected": {
            "shapeActionCount": 3,
            "resolvedCount": 2,
            "missingMacroCount": 1,
            "missingMacro": "MissingMacro",
        },
    }
    write_json(workbook_json, workbook_inventory())
    write_json(vba_json, vba_lint())
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out-dir", required=True, type=Path, help="Directory for generated fixture JSON files")
    parser.add_argument("--out-json", type=Path, help="Optional manifest path")
    args = parser.parse_args()
    manifest = create_fixture(args.out_dir.expanduser().resolve())
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), manifest)
    else:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
