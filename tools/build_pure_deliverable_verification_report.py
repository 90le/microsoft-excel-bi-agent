#!/usr/bin/env python3
"""Verify a cleaned pure deliverable against a cleanup plan.

Inputs are:

- cleanup-plan JSON from ``tools/build_pure_deliverable_cleanup_plan.py``
- post-clean readiness JSON from ``tools/build_external_dependency_report.py``

The script is read-only. It compares the plan's post-clean assertions against
the inspected post-clean workbook report and produces JSON/Markdown evidence.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"
MANUAL = "manual-review"


def normalize_bool(value: Any) -> bool | None:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lower = value.strip().lower()
        if lower == "true":
            return True
        if lower == "false":
            return False
    return None


def summary_value(summary: dict[str, Any], check: str) -> Any:
    return summary.get(check)


def evaluate_assertion(assertion: dict[str, Any], summary: dict[str, Any], target: str) -> dict[str, Any]:
    check = str(assertion.get("check", ""))
    expected = str(assertion.get("expected", ""))
    actual = summary_value(summary, check)

    if check not in summary:
        return {
            "check": check,
            "expected": expected,
            "actual": actual,
            "status": MANUAL,
            "reason": "post-clean readiness summary does not contain this evidence field",
        }

    if expected == "0" or expected.startswith("0 for "):
        try:
            actual_number = int(actual)
        except (TypeError, ValueError):
            return {
                "check": check,
                "expected": expected,
                "actual": actual,
                "status": FAIL,
                "reason": "actual value is not numeric",
            }
        return {
            "check": check,
            "expected": expected,
            "actual": actual_number,
            "status": PASS if actual_number == 0 else FAIL,
            "reason": "numeric post-clean count is zero" if actual_number == 0 else "numeric post-clean count is not zero",
        }

    if expected == "false" or expected.startswith("false"):
        actual_bool = normalize_bool(actual)
        if actual_bool is None:
            return {
                "check": check,
                "expected": expected,
                "actual": actual,
                "status": FAIL,
                "reason": "actual value is not boolean",
            }
        if actual_bool is False:
            return {
                "check": check,
                "expected": expected,
                "actual": actual_bool,
                "status": PASS,
                "reason": "boolean post-clean marker is false",
            }
        if "documented as intentionally retained" in expected:
            return {
                "check": check,
                "expected": expected,
                "actual": actual_bool,
                "status": MANUAL,
                "reason": "retained dependency requires explicit delivery documentation",
            }
        return {
            "check": check,
            "expected": expected,
            "actual": actual_bool,
            "status": FAIL,
            "reason": "boolean post-clean marker is still true",
        }

    return {
        "check": check,
        "expected": expected,
        "actual": actual,
        "status": MANUAL,
        "reason": "assertion type requires manual evidence review",
    }


def build_report(cleanup_plan: dict[str, Any], post_readiness: dict[str, Any]) -> dict[str, Any]:
    target = str(cleanup_plan.get("target", ""))
    summary = post_readiness.get("summary", {})
    if not isinstance(summary, dict):
        summary = {}

    assertion_results = [
        evaluate_assertion(assertion, summary, target)
        for assertion in cleanup_plan.get("postCleanupAssertions", [])
        if isinstance(assertion, dict)
    ]
    passed = sum(1 for item in assertion_results if item.get("status") == PASS)
    failed = sum(1 for item in assertion_results if item.get("status") == FAIL)
    manual = sum(1 for item in assertion_results if item.get("status") == MANUAL)

    if failed:
        status = FAIL
    elif manual:
        status = "manual-review-required"
    else:
        status = PASS

    return {
        "workbookPath": post_readiness.get("workbookPath", ""),
        "sourceWorkbookPath": cleanup_plan.get("workbookPath", ""),
        "target": target,
        "status": status,
        "assertionCount": len(assertion_results),
        "passedCount": passed,
        "failedCount": failed,
        "manualReviewCount": manual,
        "postCleanupReadiness": summary,
        "assertions": assertion_results,
        "limitations": [
            "This report verifies static post-clean workbook structures only.",
            "It does not prove Excel recalculation, Power Query refresh, VBA execution, or visual output correctness.",
            "For real customer workbooks, pair this report with value/range checks and, on Windows, Excel COM runtime validation.",
        ],
    }


def clean_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Pure Deliverable Verification Report",
        "",
        f"- workbook: `{report.get('workbookPath', '')}`",
        f"- source workbook: `{report.get('sourceWorkbookPath', '')}`",
        f"- target: `{report.get('target', '')}`",
        f"- status: **{report.get('status', '')}**",
        f"- assertions: {report.get('assertionCount', 0)}",
        f"- passed: {report.get('passedCount', 0)}",
        f"- failed: {report.get('failedCount', 0)}",
        f"- manual review: {report.get('manualReviewCount', 0)}",
        "",
        "| Check | Expected | Actual | Status | Reason |",
        "|---|---|---|---:|---|",
    ]
    for assertion in report.get("assertions", []):
        lines.append(
            "| "
            + " | ".join(
                [
                    clean_markdown(assertion.get("check", "")),
                    clean_markdown(assertion.get("expected", "")),
                    clean_markdown(assertion.get("actual", "")),
                    clean_markdown(assertion.get("status", "")),
                    clean_markdown(assertion.get("reason", "")),
                ]
            )
            + " |"
        )
    lines.extend(["", "## Limitations", ""])
    for item in report.get("limitations", []):
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cleanup-plan-json", required=True, type=Path, help="JSON from build_pure_deliverable_cleanup_plan.py")
    parser.add_argument("--post-readiness-json", required=True, type=Path, help="Post-clean JSON from build_external_dependency_report.py")
    parser.add_argument("--out-json", type=Path, help="Write machine-readable verification report")
    parser.add_argument("--out-md", type=Path, help="Write Markdown verification report")
    parser.add_argument("--fail-on-fail", action="store_true", help="Exit with code 1 when static assertions fail")
    args = parser.parse_args()

    cleanup_plan = json.loads(args.cleanup_plan_json.expanduser().read_text(encoding="utf-8"))
    post_readiness = json.loads(args.post_readiness_json.expanduser().read_text(encoding="utf-8"))
    report = build_report(cleanup_plan, post_readiness)

    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        args.out_md.parent.mkdir(parents=True, exist_ok=True)
        args.out_md.write_text(render_markdown(report), encoding="utf-8")
    if not args.out_json and not args.out_md:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.fail_on_fail and report.get("status") == FAIL:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
