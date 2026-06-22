#!/usr/bin/env python3
"""Build an auditable Power Query refresh timing and status report.

The input is normally JSON emitted by refresh_power_queries_excel_com.ps1. This
script does not open Excel, refresh sources, or inspect credentials. It converts
refresh evidence into a stable JSON/Markdown report that can be reviewed before
dependent VBA, Power Pivot, or delivery steps continue.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


PASS = "pass"
WARN = "warn"
FAIL = "fail"


def read_json(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise SystemExit(f"Input is not valid JSON: {path}: {exc}") from exc
    if not isinstance(data, dict):
        raise SystemExit(f"Input JSON must be an object: {path}")
    return data


def as_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def as_float(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and value.strip():
        try:
            return float(value.strip())
        except ValueError:
            return None
    return None


def truthy_refreshing(connection: Any) -> bool:
    return isinstance(connection, dict) and connection.get("refreshing") is True


def add_finding(findings: list[dict[str, Any]], severity: str, code: str, message: str, evidence: Any = None) -> None:
    item: dict[str, Any] = {"severity": severity, "code": code, "message": message}
    if evidence is not None:
        item["evidence"] = evidence
    findings.append(item)


def collect_error_messages(data: dict[str, Any]) -> list[dict[str, str]]:
    messages: list[dict[str, str]] = []
    if isinstance(data.get("error"), str) and data["error"].strip():
        messages.append({"path": "$.error", "message": data["error"].strip()})
    for index, item in enumerate(as_list(data.get("errors"))):
        if isinstance(item, dict):
            message = item.get("message")
            if isinstance(message, str) and message.strip():
                phase = item.get("phase")
                messages.append(
                    {
                        "path": f"$.errors[{index}].message",
                        "phase": str(phase or ""),
                        "message": message.strip(),
                    }
                )
        elif isinstance(item, str) and item.strip():
            messages.append({"path": f"$.errors[{index}]", "message": item.strip()})
    return messages


def determine_refresh_status(data: dict[str, Any]) -> str:
    if data.get("failedAt") or data.get("error"):
        return "failed"
    if data.get("completedAt"):
        return "completed"
    return "unknown"


def build_report(input_path: Path, max_elapsed_seconds: float | None, require_completed: bool) -> dict[str, Any]:
    data = read_json(input_path)
    findings: list[dict[str, Any]] = []
    refresh_status = determine_refresh_status(data)
    elapsed = as_float(data.get("elapsedSeconds"))
    elapsed_status = "missing" if elapsed is None else "ok"

    errors = collect_error_messages(data)
    before_connections = [item for item in as_list(data.get("beforeConnections")) if isinstance(item, dict)]
    after_connections = [item for item in as_list(data.get("afterConnections")) if isinstance(item, dict)]
    refreshing_after = [item.get("name", "") for item in after_connections if truthy_refreshing(item)]
    background_changes = [item for item in as_list(data.get("backgroundChanges")) if isinstance(item, dict)]
    target_refreshes = [item for item in as_list(data.get("targetLoadRefreshes")) if isinstance(item, dict)]

    if refresh_status == "failed":
        add_finding(findings, "error", "refresh-failed", "Refresh report contains a failedAt or top-level error field.", errors[:5])
    if errors and refresh_status != "failed":
        add_finding(findings, "warning", "refresh-errors-captured", "Refresh completed with captured errors that need review.", errors[:5])
    if require_completed and refresh_status != "completed":
        add_finding(findings, "error", "refresh-not-completed", "Refresh report does not prove a completed refresh.", refresh_status)
    if elapsed is None:
        add_finding(findings, "warning", "elapsed-missing", "Refresh report does not contain elapsedSeconds.")
    elif max_elapsed_seconds is not None and elapsed > max_elapsed_seconds:
        elapsed_status = "slow"
        add_finding(
            findings,
            "warning",
            "slow-refresh",
            f"Refresh elapsedSeconds {elapsed:.3f} exceeded threshold {max_elapsed_seconds:.3f}.",
        )
    if refreshing_after:
        add_finding(findings, "error", "connections-still-refreshing", "Connections were still refreshing after wait.", refreshing_after)
    if data.get("disableBackgroundRefresh") is not True:
        add_finding(
            findings,
            "warning",
            "background-refresh-not-disabled",
            "Refresh report does not show background refresh was disabled before dependent steps.",
        )

    error_count = sum(1 for item in findings if item["severity"] == "error")
    warning_count = sum(1 for item in findings if item["severity"] == "warning")
    status = FAIL if error_count else WARN if warning_count else PASS

    return {
        "source": str(input_path),
        "status": status,
        "refreshStatus": refresh_status,
        "workbookPath": data.get("workbookPath", ""),
        "outputWorkbookPath": data.get("outputWorkbookPath", ""),
        "queryName": data.get("queryName", ""),
        "startedAt": data.get("startedAt", ""),
        "completedAt": data.get("completedAt", ""),
        "failedAt": data.get("failedAt", ""),
        "elapsedSeconds": elapsed,
        "elapsedStatus": elapsed_status,
        "maxElapsedSeconds": max_elapsed_seconds,
        "disableBackgroundRefresh": data.get("disableBackgroundRefresh"),
        "calculateFull": data.get("calculateFull"),
        "connectionCountBefore": len(before_connections),
        "connectionCountAfter": len(after_connections),
        "refreshingConnectionCountAfter": len(refreshing_after),
        "backgroundChangeCount": len(background_changes),
        "targetLoadRefreshCount": len(target_refreshes),
        "rawErrorMessageCount": len(errors),
        "findingCount": len(findings),
        "errorFindingCount": error_count,
        "warningFindingCount": warning_count,
        "findings": findings,
        "diagnosticCommand": "classify_power_query_refresh_errors.py <refresh-json> --out-json <diagnosis.json> --out-md <diagnosis.md>",
        "boundaries": [
            "This report summarizes existing refresh evidence; it does not refresh Excel itself.",
            "Use refresh_power_queries_excel_com.ps1 for live Excel refresh and wait-for-completion evidence.",
            "Use classify_power_query_refresh_errors.py for detailed root-cause buckets when errors are present.",
        ],
    }


def markdown_report(report: dict[str, Any]) -> str:
    elapsed = report.get("elapsedSeconds")
    elapsed_text = "missing" if elapsed is None else f"{float(elapsed):.3f}s"
    lines = [
        "# Power Query Refresh Performance Report",
        "",
        f"- Source: `{report['source']}`",
        f"- Status: **{report['status']}**",
        f"- Refresh status: `{report['refreshStatus']}`",
        f"- Query: `{report.get('queryName') or '<all queries>'}`",
        f"- Elapsed: `{elapsed_text}`",
        f"- Elapsed status: `{report['elapsedStatus']}`",
        f"- Background refresh disabled: `{report.get('disableBackgroundRefresh')}`",
        f"- Target loaded refreshes: `{report['targetLoadRefreshCount']}`",
        f"- Connections after refresh: `{report['connectionCountAfter']}`",
        f"- Still refreshing after wait: `{report['refreshingConnectionCountAfter']}`",
        f"- Findings: `{report['findingCount']}`",
        "",
    ]
    if report["findings"]:
        lines.append("## Findings")
        lines.append("")
        for finding in report["findings"]:
            lines.append(f"- `{finding['severity']}` `{finding['code']}`: {finding['message']}")
        lines.append("")
    else:
        lines.append("No refresh status or timing issue was detected.")
        lines.append("")
    lines.extend(["## Boundaries", ""])
    for item in report["boundaries"]:
        lines.append(f"- {item}")
    lines.extend(["", "## Follow-Up Diagnostic Command", "", f"`{report['diagnosticCommand']}`", ""])
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Refresh JSON report from refresh_power_queries_excel_com.ps1")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--out-md", type=Path, help="Optional Markdown report path")
    parser.add_argument("--max-elapsed-seconds", type=float, help="Warn when refresh elapsedSeconds exceeds this threshold")
    parser.add_argument("--require-completed", action="store_true", help="Require completedAt evidence")
    parser.add_argument("--fail-on-error", action="store_true", help="Exit non-zero when the report status is fail")
    parser.add_argument("--fail-on-warning", action="store_true", help="Exit non-zero when the report status is warn or fail")
    parser.add_argument("--fail-on-slow", action="store_true", help="Exit non-zero when elapsedStatus is slow")
    return parser.parse_args()


def write_outputs(report: dict[str, Any], out_json: Path | None, out_md: Path | None) -> None:
    if out_json:
        path = out_json.expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if out_md:
        path = out_md.expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(markdown_report(report), encoding="utf-8")


def main() -> int:
    args = parse_args()
    report = build_report(args.input.expanduser().resolve(), args.max_elapsed_seconds, args.require_completed)
    write_outputs(report, args.out_json, args.out_md)
    print(
        "Power Query refresh report: "
        f"{report['status']}, refresh={report['refreshStatus']}, "
        f"elapsed={report['elapsedSeconds']}, findings={report['findingCount']}"
    )
    if args.fail_on_error and report["status"] == FAIL:
        return 1
    if args.fail_on_warning and report["status"] in {WARN, FAIL}:
        return 1
    if args.fail_on_slow and report["elapsedStatus"] == "slow":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
