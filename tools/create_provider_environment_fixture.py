#!/usr/bin/env python3
"""Create generic provider-environment drift fixtures.

The files are synthetic and customer-data-free. They model the shape produced
by probe_excel_bi_providers.ps1 and a prior provider environment report so the
release gate can test baseline comparison logic without depending on the local
machine's real Office/provider state.
"""

from __future__ import annotations

import argparse
import copy
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def current_probe() -> dict[str, Any]:
    return {
        "startedAt": now_iso(),
        "completedAt": now_iso(),
        "machine": {
            "computerName": "GENERIC-FIXTURE",
            "osVersion": "Generic Windows fixture",
            "is64BitOperatingSystem": True,
            "is64BitProcess": True,
            "powershellVersion": "5.1.0.0",
        },
        "providers": [
            {
                "provider": "Microsoft.ACE.OLEDB.16.0",
                "registryDetected": True,
                "registryHits": [{"path": "HKCR\\Microsoft.ACE.OLEDB.16.0", "defaultValue": "", "clsid": "{ACE16}"}],
            },
            {
                "provider": "Microsoft.ACE.OLEDB.12.0",
                "registryDetected": True,
                "registryHits": [{"path": "HKCR\\Microsoft.ACE.OLEDB.12.0", "defaultValue": "", "clsid": "{ACE12}"}],
            },
            {
                "provider": "MSOLAP",
                "registryDetected": True,
                "registryHits": [{"path": "HKCR\\MSOLAP", "defaultValue": "", "clsid": "{MSOLAP}"}],
            },
        ],
        "comProgIds": [
            {"progId": "Excel.Application", "creatable": True, "error": ""},
            {"progId": "ADODB.Connection", "creatable": True, "error": ""},
            {"progId": "ADODB.Recordset", "creatable": True, "error": ""},
            {"progId": "ADOMD.Catalog", "creatable": True, "error": ""},
            {"progId": "ADOMD.Cellset", "creatable": True, "error": ""},
        ],
        "dotNetAssemblies": [
            {
                "assemblyName": "Microsoft.AnalysisServices.AdomdClient",
                "loadable": True,
                "fullName": "Microsoft.AnalysisServices.AdomdClient, Version=19.0.0.0",
                "error": "",
            }
        ],
        "excelComSmoke": {
            "succeeded": True,
            "version": "16.0",
            "build": "17328",
            "operatingSystem": "Windows fixture",
            "hWnd": 123,
            "error": "",
        },
        "adoWorkbookSmoke": {
            "succeeded": True,
            "provider": "Microsoft.ACE.OLEDB.12.0",
            "workbookPath": "provider-fixture.xlsx",
            "rowCount": 2,
            "fields": ["Category", "TotalAmount"],
            "schemaTableCount": 1,
            "error": "",
        },
        "interpretation": [
            "Synthetic fixture for provider environment report tests.",
            "Does not prove local Office runtime availability.",
        ],
    }


def matching_summary() -> dict[str, Any]:
    return {
        "machine": current_probe()["machine"],
        "providerStatus": {
            "Microsoft.ACE.OLEDB.16.0": {"registryDetected": True, "registryHitCount": 1},
            "Microsoft.ACE.OLEDB.12.0": {"registryDetected": True, "registryHitCount": 1},
            "MSOLAP": {"registryDetected": True, "registryHitCount": 1},
        },
        "comStatus": {
            "Excel.Application": {"creatable": True, "error": ""},
            "ADODB.Connection": {"creatable": True, "error": ""},
            "ADODB.Recordset": {"creatable": True, "error": ""},
            "ADOMD.Catalog": {"creatable": True, "error": ""},
            "ADOMD.Cellset": {"creatable": True, "error": ""},
        },
        "assemblyStatus": {
            "Microsoft.AnalysisServices.AdomdClient": {
                "loadable": True,
                "fullName": "Microsoft.AnalysisServices.AdomdClient, Version=19.0.0.0",
                "error": "",
            }
        },
        "excelComSmoke": {
            "succeeded": True,
            "version": "16.0",
            "build": "17328",
            "operatingSystem": "Windows fixture",
        },
        "adoWorkbookSmoke": {
            "succeeded": True,
            "provider": "Microsoft.ACE.OLEDB.12.0",
            "rowCount": 2,
            "fields": ["Category", "TotalAmount"],
            "schemaTableCount": 1,
        },
        "readiness": {
            "excelAutomationReady": True,
            "adodbReady": True,
            "workbookSqlReady": True,
            "aceProviderRegistered": True,
            "msolapRegistered": True,
            "adomdComReady": True,
            "adomdNetLoadable": True,
        },
        "boundaries": [
            "registryDetected proves ProgID registry presence only, not a successful connection.",
            "COM creatable proves object activation only, not endpoint query success.",
            "workbookSqlReady proves generated-workbook ACE/ADO SQL only.",
            "adomdComReady does not prove a real workbook Data Model or external cube endpoint is queryable.",
        ],
    }


def report_with_summary(summary: dict[str, Any], *, status: str = "pass") -> dict[str, Any]:
    return {
        "status": status,
        "generatedAt": now_iso(),
        "projectRoot": "",
        "source": {
            "probeJson": "provider_fixture_probe.json",
            "runProbe": False,
            "excelComSmokeRequested": True,
            "adoWorkbookSmokeRequested": True,
            "adoSmokeProvider": "Microsoft.ACE.OLEDB.12.0",
        },
        "summary": summary,
        "comparison": {"baselineCompared": False},
        "probeStdout": "",
        "probeStderr": "",
        "errors": [],
    }


def drifting_baseline() -> dict[str, Any]:
    summary = copy.deepcopy(matching_summary())
    summary["providerStatus"]["MSOLAP"]["registryDetected"] = False
    summary["providerStatus"]["MSOLAP"]["registryHitCount"] = 0
    summary["comStatus"]["ADOMD.Cellset"]["creatable"] = False
    summary["comStatus"]["ADOMD.Cellset"]["error"] = "Synthetic missing ADOMD.Cellset"
    summary["assemblyStatus"]["Microsoft.AnalysisServices.AdomdClient"]["loadable"] = False
    summary["assemblyStatus"]["Microsoft.AnalysisServices.AdomdClient"]["error"] = "Synthetic missing assembly"
    summary["excelComSmoke"]["version"] = "15.0"
    summary["excelComSmoke"]["build"] = "00000"
    summary["adoWorkbookSmoke"]["succeeded"] = False
    summary["adoWorkbookSmoke"]["rowCount"] = 0
    summary["adoWorkbookSmoke"]["fields"] = []
    summary["adoWorkbookSmoke"]["schemaTableCount"] = 0
    summary["readiness"]["workbookSqlReady"] = False
    summary["readiness"]["msolapRegistered"] = False
    summary["readiness"]["adomdComReady"] = False
    summary["readiness"]["adomdNetLoadable"] = False
    return report_with_summary(summary)


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def create_fixture(out_dir: Path) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    current_probe_path = out_dir / "provider_fixture_probe.json"
    matching_baseline_path = out_dir / "provider_matching_baseline.json"
    drifting_baseline_path = out_dir / "provider_drifting_baseline.json"
    manifest_path = out_dir / "provider_environment_fixture.json"

    write_json(current_probe_path, current_probe())
    write_json(matching_baseline_path, report_with_summary(matching_summary()))
    write_json(drifting_baseline_path, drifting_baseline())

    manifest = {
        "createdAt": now_iso(),
        "currentProbe": str(current_probe_path),
        "matchingBaseline": str(matching_baseline_path),
        "driftingBaseline": str(drifting_baseline_path),
        "expected": {
            "matchingChangedCount": 0,
            "driftingMinimumChangedCount": 6,
            "driftPaths": [
                "providerStatus.MSOLAP",
                "comStatus.ADOMD.Cellset",
                "assemblyStatus.Microsoft.AnalysisServices.AdomdClient",
                "readiness.workbookSqlReady",
                "readiness.msolapRegistered",
                "readiness.adomdComReady",
                "excelComSmoke.version",
                "adoWorkbookSmoke.succeeded",
            ],
        },
    }
    write_json(manifest_path, manifest)
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out-dir", required=True, type=Path, help="Directory for generated fixture JSON files")
    parser.add_argument("--out-json", type=Path, help="Optional manifest JSON path")
    args = parser.parse_args()

    manifest = create_fixture(args.out_dir.expanduser().resolve())
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), manifest)
    else:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
