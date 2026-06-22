#!/usr/bin/env python3
"""Build an official-documentation URL drift report.

The default mode is offline and deterministic. It inventories bundled
official-docs indexes, summarizes Microsoft official URLs, and can compare
against a prior report. Use --check-online only when live link drift evidence is
needed and network access is acceptable.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from validate_official_docs_index import (
    DEFAULT_ONLINE_TIMEOUT,
    DEFAULT_USER_AGENT,
    check_url,
    find_indexes,
    iter_official_urls,
    validate_index,
)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def normalize_records(records: list[dict[str, str]]) -> list[dict[str, str]]:
    normalized: list[dict[str, str]] = []
    for record in records:
        url = record.get("url", "").strip()
        parsed = urlparse(url)
        normalized.append(
            {
                "skill": record.get("skill", "").strip(),
                "index": record.get("index", "").strip(),
                "id": record.get("id", "").strip(),
                "title": record.get("title", "").strip(),
                "url": url,
                "host": parsed.netloc.lower(),
                "path": parsed.path,
            }
        )
    return sorted(normalized, key=lambda item: (item["skill"], item["id"], item["url"]))


def inventory(project_root: Path) -> tuple[list[str], dict[str, Any], list[dict[str, str]]]:
    indexes = find_indexes(project_root)
    errors: list[str] = []
    index_summaries: list[dict[str, Any]] = []
    for path in indexes:
        index_errors, summary = validate_index(path, project_root)
        errors.extend(index_errors)
        index_summaries.append(summary)

    records = normalize_records(iter_official_urls(indexes, project_root))
    host_counts = Counter(record["host"] for record in records)
    skill_counts = Counter(record["skill"] for record in records)
    category_counts: Counter[str] = Counter()
    for path in indexes:
        data = load_json(path)
        for entry in data.get("entries", []):
            if isinstance(entry, dict):
                category_counts[str(entry.get("category", "")).strip()] += 1

    summary = {
        "indexCount": len(indexes),
        "entryCount": sum(int(item.get("entryCount", 0)) for item in index_summaries),
        "uniqueUrlCount": len({record["url"] for record in records}),
        "hostCounts": dict(sorted(host_counts.items())),
        "skillCounts": dict(sorted(skill_counts.items())),
        "categoryCounts": dict(sorted((key, value) for key, value in category_counts.items() if key)),
        "indexes": index_summaries,
    }
    return errors, summary, records


def baseline_records(path: Path) -> list[dict[str, str]]:
    data = load_json(path)
    if isinstance(data.get("records"), list):
        return normalize_records([item for item in data["records"] if isinstance(item, dict)])
    if isinstance(data.get("officialUrls"), list):
        return normalize_records([item for item in data["officialUrls"] if isinstance(item, dict)])
    raise ValueError(f"baseline does not contain records or officialUrls: {path}")


def compare_baseline(current: list[dict[str, str]], baseline: list[dict[str, str]]) -> dict[str, Any]:
    current_by_key = {(item["skill"], item["id"]): item for item in current}
    baseline_by_key = {(item["skill"], item["id"]): item for item in baseline}
    current_keys = set(current_by_key)
    baseline_keys = set(baseline_by_key)

    added = [current_by_key[key] for key in sorted(current_keys - baseline_keys)]
    removed = [baseline_by_key[key] for key in sorted(baseline_keys - current_keys)]
    changed: list[dict[str, Any]] = []
    for key in sorted(current_keys & baseline_keys):
        before = baseline_by_key[key]
        after = current_by_key[key]
        if before["url"] != after["url"] or before["title"] != after["title"]:
            changed.append({"key": {"skill": key[0], "id": key[1]}, "before": before, "after": after})

    return {
        "baselineCompared": True,
        "addedCount": len(added),
        "removedCount": len(removed),
        "changedCount": len(changed),
        "added": added,
        "removed": removed,
        "changed": changed,
    }


def run_online_checks(records: list[dict[str, str]], limit: int, timeout: float, user_agent: str) -> tuple[list[str], list[dict[str, Any]]]:
    selected = records[:limit] if limit > 0 else records
    errors: list[str] = []
    results: list[dict[str, Any]] = []
    for record in selected:
        ok, detail, status = check_url(record["url"], timeout, user_agent)
        item: dict[str, Any] = dict(record)
        item.update({"ok": ok, "detail": detail, "status": status})
        results.append(item)
        if not ok:
            errors.append(f"{record['skill']}:{record['id']} {record['url']} ({detail})")
    return errors, results


def build_report(
    project_root: Path,
    *,
    baseline_json: Path | None,
    check_online: bool,
    online_limit: int,
    online_timeout: float,
    online_user_agent: str,
    fail_on_drift: bool,
) -> dict[str, Any]:
    errors, summary, records = inventory(project_root)
    comparison: dict[str, Any] = {"baselineCompared": False}
    if baseline_json:
        try:
            comparison = compare_baseline(records, baseline_records(baseline_json))
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            errors.append(f"cannot read baseline: {exc}")
        if fail_on_drift and (
            comparison.get("addedCount", 0) or comparison.get("removedCount", 0) or comparison.get("changedCount", 0)
        ):
            errors.append(
                "baseline drift detected: "
                f"added={comparison.get('addedCount')}, "
                f"removed={comparison.get('removedCount')}, "
                f"changed={comparison.get('changedCount')}"
            )

    online_results: list[dict[str, Any]] = []
    if check_online:
        online_errors, online_results = run_online_checks(records, online_limit, online_timeout, online_user_agent)
        errors.extend(online_errors)

    return {
        "status": "pass" if not errors else "fail",
        "generatedAt": now_iso(),
        "projectRoot": str(project_root),
        "summary": summary,
        "comparison": comparison,
        "onlineCheckEnabled": check_online,
        "onlineChecks": online_results,
        "records": records,
        "errors": errors,
    }


def markdown_report(report: dict[str, Any]) -> str:
    summary = report["summary"]
    comparison = report["comparison"]
    lines = [
        "# Official Documentation Drift Report",
        "",
        f"status: **{report['status']}**",
        "",
        "## Summary",
        "",
        "| Metric | Value |",
        "|---|---:|",
        f"| Indexes | {summary['indexCount']} |",
        f"| Entries | {summary['entryCount']} |",
        f"| Unique URLs | {summary['uniqueUrlCount']} |",
        f"| Online checks | {len(report['onlineChecks']) if report['onlineCheckEnabled'] else 0} |",
        "",
        "## Host Coverage",
        "",
        "| Host | URLs |",
        "|---|---:|",
    ]
    for host, count in summary["hostCounts"].items():
        lines.append(f"| `{host}` | {count} |")

    lines.extend(["", "## Skill Coverage", "", "| Skill | URLs |", "|---|---:|"])
    for skill, count in summary["skillCounts"].items():
        lines.append(f"| `{skill}` | {count} |")

    lines.extend(["", "## Baseline Comparison", ""])
    if comparison.get("baselineCompared"):
        lines.extend(
            [
                "| Added | Removed | Changed |",
                "|---:|---:|---:|",
                f"| {comparison['addedCount']} | {comparison['removedCount']} | {comparison['changedCount']} |",
            ]
        )
    else:
        lines.append("No baseline supplied.")

    lines.extend(["", "## Errors", ""])
    if report["errors"]:
        for error in report["errors"]:
            lines.append(f"- {error}")
    else:
        lines.append("No official-documentation drift errors found.")

    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--baseline-json", type=Path, help="Optional prior drift report JSON to compare")
    parser.add_argument("--fail-on-drift", action="store_true", help="Fail when a supplied baseline differs")
    parser.add_argument("--check-online", action="store_true", help="Check Microsoft official URL reachability")
    parser.add_argument("--online-limit", type=int, default=0, help="Maximum online URLs to check; 0 checks all")
    parser.add_argument("--online-timeout", type=float, default=DEFAULT_ONLINE_TIMEOUT, help="Per-request timeout")
    parser.add_argument("--online-user-agent", default=DEFAULT_USER_AGENT, help="User-Agent for online checks")
    parser.add_argument("--out-json", default="", help="Write JSON report")
    parser.add_argument("--out-md", default="", help="Write Markdown report")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero if status is not pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    report = build_report(
        project_root,
        baseline_json=args.baseline_json.expanduser().resolve() if args.baseline_json else None,
        check_online=args.check_online,
        online_limit=args.online_limit,
        online_timeout=args.online_timeout,
        online_user_agent=args.online_user_agent,
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

    summary = report["summary"]
    print(
        "Official docs drift {status}: indexes={indexes}, entries={entries}, urls={urls}, online={online}".format(
            status=report["status"],
            indexes=summary["indexCount"],
            entries=summary["entryCount"],
            urls=summary["uniqueUrlCount"],
            online=len(report["onlineChecks"]) if report["onlineCheckEnabled"] else 0,
        )
    )

    if args.require_pass and report["status"] != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
