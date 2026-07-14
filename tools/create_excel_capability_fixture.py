#!/usr/bin/env python3
"""Create synthetic Excel capability probe fixtures for report tests."""

from __future__ import annotations

import argparse
import copy
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


TOOLS_DIR = Path(__file__).resolve().parent
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

from build_excel_compatibility_report import CAPABILITY_IDS  # noqa: E402


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def capability(status: str = "pass", evidence_level: str = "smoke", detail: str = "Synthetic fixture evidence.") -> dict[str, str]:
    return {
        "status": status,
        "evidenceLevel": evidence_level,
        "detail": detail,
        "errorCategory": "" if status == "pass" else "synthetic-fixture",
        "error": "" if status == "pass" else detail,
    }


def all_supported_probe() -> dict[str, Any]:
    return {
        "schemaVersion": "1.0",
        "kind": "excel-capability-probe",
        "generatedAt": now_iso(),
        "probe": {"profile": "runtime", "platform": "windows", "syntheticFixture": True},
        "environment": {
            "osVersion": "Generic Windows fixture",
            "is64BitOperatingSystem": True,
            "is64BitProcess": True,
            "powershellVersion": "5.1.0.0",
            "excelVersion": "16.0",
            "excelBuild": "00000",
        },
        "capabilities": {item: capability() for item in CAPABILITY_IDS},
        "boundaries": ["Synthetic fixture; not evidence of local Office readiness."],
        "errors": [],
    }


def fixture_cases() -> dict[str, dict[str, Any]]:
    supported = all_supported_probe()
    blocked = copy.deepcopy(supported)
    blocked["capabilities"]["excel.com.activation"] = capability("fail", "activation", "Synthetic COM failure.")
    blocked["capabilities"]["excel.workbook.roundtrip"] = capability("skip", "not-tested", "Blocked by COM failure.")

    partial = copy.deepcopy(supported)
    partial["probe"]["profile"] = "inventory"
    for capability_id in [
        "excel.com.activation",
        "excel.workbook.roundtrip",
        "excel.vba.project-access",
        "excel.power-query.object-model",
        "excel.power-query.async-wait",
        "excel.data-model.object-model",
        "excel.pdf-export",
        "ace.workbook-sql",
    ]:
        partial["capabilities"][capability_id] = capability("skip", "not-tested", "Runtime smoke was not requested.")

    malformed = copy.deepcopy(supported)
    malformed.pop("schemaVersion")
    return {
        "allSupported": supported,
        "coreBlocked": blocked,
        "partialEvidence": partial,
        "malformedContract": malformed,
    }


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def create_fixture(out_dir: Path) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    names = {
        "allSupported": "excel_capability_all_supported.json",
        "coreBlocked": "excel_capability_core_blocked.json",
        "partialEvidence": "excel_capability_partial_evidence.json",
        "malformedContract": "excel_capability_malformed_contract.json",
    }
    cases = fixture_cases()
    paths: dict[str, str] = {}
    for case_id, filename in names.items():
        path = out_dir / filename
        write_json(path, cases[case_id])
        paths[case_id] = str(path)
    manifest = {
        "createdAt": now_iso(),
        "cases": paths,
        "expected": {
            "allSupported": {"status": "pass", "passCount": 11},
            "coreBlocked": {"workbookAutomation": "blocked"},
            "partialEvidence": {"workbookAutomation": "unknown"},
            "malformedContract": {"status": "fail"},
        },
        "boundaries": ["Fixtures prove report behavior only, not local Office readiness."],
    }
    write_json(out_dir / "excel_capability_fixture_manifest.json", manifest)
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
