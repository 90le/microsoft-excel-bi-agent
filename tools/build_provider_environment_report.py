#!/usr/bin/env python3
"""Build a provider environment and drift report for Excel BI automation.

The report normalizes the output from probe_excel_bi_providers.ps1 into a
capability matrix for Office COM, ADO/OLEDB, MSOLAP, ADOMD, and workbook SQL.
Use --run-probe on Windows to generate fresh evidence, or --probe-json to
summarize a previously captured probe result.
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


CORE_PROVIDERS = [
    "Microsoft.ACE.OLEDB.16.0",
    "Microsoft.ACE.OLEDB.12.0",
    "MSOLAP",
]
CORE_COM_PROG_IDS = [
    "Excel.Application",
    "ADODB.Connection",
    "ADODB.Recordset",
    "ADOMD.Catalog",
    "ADOMD.Cellset",
]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def find_powershell() -> str | None:
    for name in ["powershell", "powershell.exe", "pwsh", "pwsh.exe"]:
        candidate = shutil.which(name)
        if candidate:
            return candidate
    return None


def run_provider_probe(
    project_root: Path,
    *,
    powershell: str | None,
    excel_com: bool,
    ado_workbook_smoke: bool,
    ado_smoke_provider: str,
) -> tuple[dict[str, Any] | None, list[str], str, str]:
    if platform.system().lower() != "windows":
        return None, ["--run-probe requires Windows"], "", ""
    ps_exe = powershell or find_powershell()
    if not ps_exe:
        return None, ["PowerShell was not found"], "", ""
    script = project_root / "tools" / "probe_excel_bi_providers.ps1"
    if not script.is_file():
        return None, [f"probe script not found: {script}"], "", ""

    with tempfile.TemporaryDirectory(prefix="excel_bi_provider_env_") as tmp:
        tmp_dir = Path(tmp)
        out_json = tmp_dir / "provider_probe.json"
        smoke_workbook = tmp_dir / "provider environment ado smoke.xlsx"
        command = [
            ps_exe,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-OutJson",
            str(out_json),
            "-AdoSmokeProvider",
            ado_smoke_provider,
        ]
        if excel_com:
            command.append("-RunExcelComSmoke")
        if ado_workbook_smoke:
            command.extend(["-RunAdoWorkbookSmoke", "-SmokeWorkbookPath", str(smoke_workbook)])

        env = os.environ.copy()
        env["PYTHONDONTWRITEBYTECODE"] = "1"
        completed = subprocess.run(command, cwd=str(project_root), text=True, capture_output=True, env=env, timeout=180)
        if completed.returncode != 0:
            return None, [f"provider probe failed with exit_code={completed.returncode}"], completed.stdout, completed.stderr
        try:
            return load_json(out_json), [], completed.stdout, completed.stderr
        except (OSError, json.JSONDecodeError) as exc:
            return None, [f"cannot read provider probe JSON: {exc}"], completed.stdout, completed.stderr


def bool_from_item(item: dict[str, Any], key: str) -> bool:
    return item.get(key) is True


def summarize_probe(probe: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    errors: list[str] = []
    providers = probe.get("providers", [])
    com_prog_ids = probe.get("comProgIds", [])
    dotnet = probe.get("dotNetAssemblies", [])
    if not isinstance(providers, list):
        errors.append("providers is not a list")
        providers = []
    if not isinstance(com_prog_ids, list):
        errors.append("comProgIds is not a list")
        com_prog_ids = []
    if not isinstance(dotnet, list):
        errors.append("dotNetAssemblies is not a list")
        dotnet = []

    provider_status: dict[str, dict[str, Any]] = {}
    for item in providers:
        if not isinstance(item, dict):
            continue
        name = str(item.get("provider", "")).strip()
        if not name:
            continue
        hits = item.get("registryHits", [])
        hit_count = len(hits) if isinstance(hits, list) else 0
        provider_status[name] = {
            "registryDetected": item.get("registryDetected") is True,
            "registryHitCount": hit_count,
        }

    com_status: dict[str, dict[str, Any]] = {}
    for item in com_prog_ids:
        if not isinstance(item, dict):
            continue
        name = str(item.get("progId", "")).strip()
        if not name:
            continue
        com_status[name] = {
            "creatable": item.get("creatable") is True,
            "error": str(item.get("error", "") or ""),
        }

    assembly_status: dict[str, dict[str, Any]] = {}
    for item in dotnet:
        if not isinstance(item, dict):
            continue
        name = str(item.get("assemblyName", "")).strip()
        if not name:
            continue
        assembly_status[name] = {
            "loadable": item.get("loadable") is True,
            "fullName": str(item.get("fullName", "") or ""),
            "error": str(item.get("error", "") or ""),
        }

    excel_smoke = probe.get("excelComSmoke")
    if not isinstance(excel_smoke, dict):
        excel_smoke = {}
    ado_smoke = probe.get("adoWorkbookSmoke")
    if not isinstance(ado_smoke, dict):
        ado_smoke = {}

    readiness = {
        "excelAutomationReady": bool_from_item(com_status.get("Excel.Application", {}), "creatable")
        and excel_smoke.get("succeeded") is True,
        "adodbReady": bool_from_item(com_status.get("ADODB.Connection", {}), "creatable"),
        "workbookSqlReady": ado_smoke.get("succeeded") is True,
        "aceProviderRegistered": any(
            provider_status.get(name, {}).get("registryDetected") is True
            for name in ["Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"]
        ),
        "msolapRegistered": any(
            name.upper().startswith("MSOLAP") and data.get("registryDetected") is True
            for name, data in provider_status.items()
        ),
        "adomdComReady": bool_from_item(com_status.get("ADOMD.Catalog", {}), "creatable")
        and bool_from_item(com_status.get("ADOMD.Cellset", {}), "creatable"),
        "adomdNetLoadable": any(item.get("loadable") is True for item in assembly_status.values()),
    }

    for name in CORE_PROVIDERS:
        if name not in provider_status:
            errors.append(f"missing provider row: {name}")
    for name in CORE_COM_PROG_IDS:
        if name not in com_status:
            errors.append(f"missing COM row: {name}")
    if "machine" not in probe or not isinstance(probe.get("machine"), dict):
        errors.append("machine section missing")

    summary = {
        "machine": probe.get("machine", {}),
        "providerStatus": provider_status,
        "comStatus": com_status,
        "assemblyStatus": assembly_status,
        "excelComSmoke": {
            "succeeded": excel_smoke.get("succeeded"),
            "version": str(excel_smoke.get("version", "") or ""),
            "build": str(excel_smoke.get("build", "") or ""),
            "operatingSystem": str(excel_smoke.get("operatingSystem", "") or ""),
        },
        "adoWorkbookSmoke": {
            "succeeded": ado_smoke.get("succeeded"),
            "provider": str(ado_smoke.get("provider", "") or ""),
            "rowCount": ado_smoke.get("rowCount"),
            "fields": ado_smoke.get("fields", []),
            "schemaTableCount": ado_smoke.get("schemaTableCount"),
        },
        "readiness": readiness,
        "boundaries": [
            "registryDetected proves ProgID registry presence only, not a successful connection.",
            "COM creatable proves object activation only, not endpoint query success.",
            "workbookSqlReady proves generated-workbook ACE/ADO SQL only.",
            "adomdComReady does not prove a real workbook Data Model or external cube endpoint is queryable.",
        ],
    }
    return summary, errors


def comparable(summary: dict[str, Any]) -> dict[str, Any]:
    return {
        "providerStatus": {
            key: value.get("registryDetected")
            for key, value in sorted(summary.get("providerStatus", {}).items())
        },
        "comStatus": {
            key: value.get("creatable")
            for key, value in sorted(summary.get("comStatus", {}).items())
        },
        "assemblyStatus": {
            key: value.get("loadable")
            for key, value in sorted(summary.get("assemblyStatus", {}).items())
        },
        "readiness": summary.get("readiness", {}),
        "excelComSmoke": summary.get("excelComSmoke", {}),
        "adoWorkbookSmoke": summary.get("adoWorkbookSmoke", {}),
    }


def baseline_summary(path: Path) -> dict[str, Any]:
    data = load_json(path)
    if isinstance(data.get("summary"), dict):
        return data["summary"]
    raise ValueError(f"baseline does not contain a summary object: {path}")


def compare_summary(current: dict[str, Any], baseline: dict[str, Any]) -> dict[str, Any]:
    current_comp = comparable(current)
    baseline_comp = comparable(baseline)
    changes: list[dict[str, Any]] = []

    def walk(prefix: str, left: Any, right: Any) -> None:
        if isinstance(left, dict) and isinstance(right, dict):
            keys = sorted(set(left) | set(right))
            for key in keys:
                walk(f"{prefix}.{key}" if prefix else str(key), left.get(key), right.get(key))
            return
        if left != right:
            changes.append({"path": prefix, "before": right, "after": left})

    walk("", current_comp, baseline_comp)
    return {
        "baselineCompared": True,
        "changedCount": len(changes),
        "changes": changes,
    }


def build_report(
    project_root: Path,
    *,
    probe_json: Path | None,
    run_probe: bool,
    powershell: str | None,
    excel_com: bool,
    ado_workbook_smoke: bool,
    ado_smoke_provider: str,
    baseline_json: Path | None,
    fail_on_drift: bool,
) -> dict[str, Any]:
    errors: list[str] = []
    stdout = ""
    stderr = ""
    if probe_json:
        try:
            probe = load_json(probe_json)
        except (OSError, json.JSONDecodeError) as exc:
            probe = {}
            errors.append(f"cannot read probe JSON: {exc}")
    elif run_probe:
        probe, probe_errors, stdout, stderr = run_provider_probe(
            project_root,
            powershell=powershell,
            excel_com=excel_com,
            ado_workbook_smoke=ado_workbook_smoke,
            ado_smoke_provider=ado_smoke_provider,
        )
        errors.extend(probe_errors)
        if probe is None:
            probe = {}
    else:
        probe = {}
        errors.append("Supply --probe-json or --run-probe")

    summary, summary_errors = summarize_probe(probe)
    errors.extend(summary_errors)

    if excel_com and summary["excelComSmoke"].get("succeeded") is not True:
        errors.append("Excel COM smoke did not succeed")
    if ado_workbook_smoke and summary["adoWorkbookSmoke"].get("succeeded") is not True:
        errors.append("ADO workbook smoke did not succeed")
    if summary["comStatus"].get("Excel.Application", {}).get("creatable") is not True:
        errors.append("Excel.Application is not creatable")
    if summary["comStatus"].get("ADODB.Connection", {}).get("creatable") is not True:
        errors.append("ADODB.Connection is not creatable")

    comparison: dict[str, Any] = {"baselineCompared": False}
    if baseline_json:
        try:
            comparison = compare_summary(summary, baseline_summary(baseline_json))
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            errors.append(f"cannot read baseline: {exc}")
        if fail_on_drift and comparison.get("changedCount", 0):
            errors.append(f"provider environment drift detected: changed={comparison.get('changedCount')}")

    return {
        "status": "pass" if not errors else "fail",
        "generatedAt": now_iso(),
        "projectRoot": str(project_root),
        "source": {
            "probeJson": str(probe_json) if probe_json else "",
            "runProbe": run_probe,
            "excelComSmokeRequested": excel_com,
            "adoWorkbookSmokeRequested": ado_workbook_smoke,
            "adoSmokeProvider": ado_smoke_provider,
        },
        "summary": summary,
        "comparison": comparison,
        "probeStdout": stdout.strip(),
        "probeStderr": stderr.strip(),
        "errors": errors,
    }


def markdown_report(report: dict[str, Any]) -> str:
    summary = report["summary"]
    readiness = summary["readiness"]
    comparison = report["comparison"]
    lines = [
        "# Provider Environment Report",
        "",
        f"status: **{report['status']}**",
        "",
        "## Readiness",
        "",
        "| Capability | Ready |",
        "|---|---:|",
    ]
    for key, value in readiness.items():
        lines.append(f"| `{key}` | {str(value).lower()} |")

    lines.extend(["", "## Core Providers", "", "| Provider | Registry detected | Hits |", "|---|---:|---:|"])
    for name in CORE_PROVIDERS:
        item = summary["providerStatus"].get(name, {})
        lines.append(f"| `{name}` | {str(item.get('registryDetected')).lower()} | {item.get('registryHitCount', 0)} |")

    lines.extend(["", "## Core COM ProgIDs", "", "| ProgID | Creatable |", "|---|---:|"])
    for name in CORE_COM_PROG_IDS:
        item = summary["comStatus"].get(name, {})
        lines.append(f"| `{name}` | {str(item.get('creatable')).lower()} |")

    excel = summary["excelComSmoke"]
    ado = summary["adoWorkbookSmoke"]
    lines.extend(
        [
            "",
            "## Runtime Smoke",
            "",
            f"- Excel COM smoke: `{excel.get('succeeded')}`; version `{excel.get('version')}` build `{excel.get('build')}`.",
            f"- ADO workbook smoke: `{ado.get('succeeded')}`; provider `{ado.get('provider')}`; rowCount `{ado.get('rowCount')}`.",
            "",
            "## Baseline Comparison",
            "",
        ]
    )
    if comparison.get("baselineCompared"):
        lines.append(f"- Changed fields: `{comparison.get('changedCount')}`.")
    else:
        lines.append("- No baseline supplied.")

    lines.extend(["", "## Boundaries", ""])
    for item in summary["boundaries"]:
        lines.append(f"- {item}")

    lines.extend(["", "## Errors", ""])
    if report["errors"]:
        for error in report["errors"]:
            lines.append(f"- {error}")
    else:
        lines.append("No provider environment errors found.")
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--probe-json", type=Path, help="Existing probe_excel_bi_providers.ps1 JSON output")
    parser.add_argument("--run-probe", action="store_true", help="Run probe_excel_bi_providers.ps1 before building the report")
    parser.add_argument("--powershell", default="", help="Optional PowerShell executable")
    parser.add_argument("--excel-com", action="store_true", help="Request Excel COM smoke when --run-probe is used")
    parser.add_argument("--ado-workbook-smoke", action="store_true", help="Request ADO workbook smoke when --run-probe is used")
    parser.add_argument("--ado-smoke-provider", default="Microsoft.ACE.OLEDB.12.0", help="Provider for ADO workbook smoke")
    parser.add_argument("--baseline-json", type=Path, help="Optional prior provider report JSON to compare")
    parser.add_argument("--fail-on-drift", action="store_true", help="Fail when a supplied baseline differs")
    parser.add_argument("--out-json", default="", help="Write JSON report")
    parser.add_argument("--out-md", default="", help="Write Markdown report")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero when status is not pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    report = build_report(
        project_root,
        probe_json=args.probe_json.expanduser().resolve() if args.probe_json else None,
        run_probe=args.run_probe,
        powershell=args.powershell or None,
        excel_com=args.excel_com,
        ado_workbook_smoke=args.ado_workbook_smoke,
        ado_smoke_provider=args.ado_smoke_provider,
        baseline_json=args.baseline_json.expanduser().resolve() if args.baseline_json else None,
        fail_on_drift=args.fail_on_drift,
    )

    if args.out_json:
        out_json = Path(args.out_json).expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = Path(args.out_md).expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_report(report), encoding="utf-8")

    readiness = report["summary"]["readiness"]
    print(
        "Provider environment {status}: excel={excel}, workbookSql={workbook_sql}, msolap={msolap}, adomdCom={adomd}".format(
            status=report["status"],
            excel=readiness.get("excelAutomationReady"),
            workbook_sql=readiness.get("workbookSqlReady"),
            msolap=readiness.get("msolapRegistered"),
            adomd=readiness.get("adomdComReady"),
        )
    )

    if args.require_pass and report["status"] != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
