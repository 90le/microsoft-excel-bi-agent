#!/usr/bin/env python3
"""Create generic OpenXML-inspection JSON fixtures for formula quality reports."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def safe_openxml_report() -> dict[str, Any]:
    formulas = [
        {"sheet": "Report", "cell": "B2", "formula": "SUM(Data!B2:B5)", "formulaType": "formula", "cachedValue": "500"},
        {"sheet": "Report", "cell": "B3", "formula": 'IFERROR(XLOOKUP(Control!B4,Data!A2:A5,Data!C2:C5,""),"")', "formulaType": "formula", "cachedValue": "18.4"},
        {"sheet": "Report", "cell": "B4", "formula": "ROUND(AVERAGE(Data!C2:C5),2)", "formulaType": "formula", "cachedValue": "17.83"},
        {"sheet": "Report", "cell": "B5", "formula": "SUMIFS(Data!B2:B5,Data!D2:D5,Control!B2)", "formulaType": "formula", "cachedValue": "300"},
    ]
    return {
        "workbookPath": "formula_quality_safe_fixture.xlsx",
        "sheets": [{"name": "Data"}, {"name": "Report"}, {"name": "Control"}],
        "formulas": formulas,
        "totalFormulaCount": len(formulas),
        "formulaFunctionCounts": {"AVERAGE": 1, "IFERROR": 1, "ROUND": 1, "SUM": 1, "SUMIFS": 1, "XLOOKUP": 1},
        "cubeFormulaCount": 0,
        "cubeFormulas": [],
        "connections": [],
        "externalLinks": [],
        "definedNames": [],
        "hasVbaProject": False,
    }


def risky_openxml_report() -> dict[str, Any]:
    formulas = [
        {"sheet": "Report", "cell": "B2", "formula": "SUM(Data!B2:B5)", "formulaType": "formula", "cachedValue": "#VALUE!"},
        {"sheet": "Report", "cell": "B3", "formula": "VLOOKUP(A3,Missing!#REF!,2,FALSE)", "formulaType": "formula", "cachedValue": "#REF!"},
        {"sheet": "Report", "cell": "B4", "formula": "INDIRECT(Control!B2)", "formulaType": "formula", "cachedValue": "10"},
        {"sheet": "Report", "cell": "B5", "formula": "NOW()", "formulaType": "formula", "cachedValue": "45123.5"},
        {"sheet": "Report", "cell": "B6", "formula": "'C:\\Users\\analyst\\Desktop\\[source.xlsx]Data'!A1", "formulaType": "formula", "cachedValue": "1"},
    ]
    return {
        "workbookPath": "formula_quality_risky_fixture.xlsx",
        "sheets": [{"name": "Data"}, {"name": "Report"}, {"name": "Control"}],
        "formulas": formulas,
        "totalFormulaCount": len(formulas),
        "formulaFunctionCounts": {"INDIRECT": 1, "NOW": 1, "SUM": 1, "VLOOKUP": 1},
        "cubeFormulaCount": 0,
        "cubeFormulas": [],
        "connections": [],
        "externalLinks": [],
        "definedNames": [],
        "hasVbaProject": False,
    }


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def create_fixture(out_dir: Path) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_json = out_dir / "formula_quality_safe_openxml.json"
    risky_json = out_dir / "formula_quality_risky_openxml.json"
    write_json(safe_json, safe_openxml_report())
    write_json(risky_json, risky_openxml_report())
    return {
        "safeOpenXmlJson": str(safe_json),
        "riskyOpenXmlJson": str(risky_json),
        "expected": {
            "safe": {
                "readiness": "clean",
                "formulaCount": 4,
                "findingCount": 0,
            },
            "risky": {
                "readiness": "blocked-for-delivery",
                "formulaCount": 5,
                "highFindingCount": 4,
                "mediumFindingCount": 1,
                "lowFindingCount": 1,
                "requiredCodes": [
                    "cached-formula-error",
                    "formula-ref-error",
                    "local-path-formula",
                    "dynamic-reference-function",
                    "volatile-function",
                ],
            },
        },
    }


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
