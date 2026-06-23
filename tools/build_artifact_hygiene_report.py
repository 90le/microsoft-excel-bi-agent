#!/usr/bin/env python3
"""Build an artifact hygiene report for the Excel BI plugin package.

The report is intentionally conservative. It prevents customer workbooks,
local screenshots, generated release reports, lock files, Python bytecode, and
customer-specific markers from being shipped inside the reusable plugin.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path


OFFICE_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
TEXT_EXTENSIONS = {
    ".bas",
    ".cls",
    ".frm",
    ".json",
    ".m",
    ".md",
    ".ps1",
    ".py",
    ".sh",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}
ALLOWED_OFFICE_FILES: set[str] = set()
IGNORED_DIRS = {".git"}


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_rel(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def customer_markers() -> list[str]:
    return [
        "budget" + "_" + "optimizer",
        "\u9884\u7b97",
        "\u79d2\u9488",
        "\u7eafVBA",
        "\u52a8\u6001\u63a8\u8350",
        "\u7acb\u767d",
    ]


def local_path_markers() -> list[str]:
    return [
        "WX" + "Work",
        "codex" + "-" + "clipboard",
        "App" + "Data" + "\\Local\\Temp",
        "Temp" + "State" + "\\ScreenClip",
        "SEI" + "\u62a5\u544a\u81ea\u52a8\u5316\u9700\u6c42",
        "\u4f01\u4e1a\u5fae\u4fe1\u622a\u56fe",
    ]


def all_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        rel_parts = path.relative_to(root).parts
        if any(part in IGNORED_DIRS for part in rel_parts):
            continue
        files.append(path)
    return sorted(files)


def add_issue(
    issues: list[dict[str, object]],
    *,
    code: str,
    severity: str,
    path: str,
    evidence: str,
    recommended_action: str,
    line: int | None = None,
) -> None:
    issue: dict[str, object] = {
        "code": code,
        "severity": severity,
        "path": path,
        "evidence": evidence[:240],
        "recommendedAction": recommended_action,
    }
    if line is not None:
        issue["line"] = line
    issues.append(issue)


def scan_file_inventory(project_root: Path, files: list[Path]) -> tuple[list[dict[str, object]], list[str]]:
    issues: list[dict[str, object]] = []
    allowed_office_seen: list[str] = []

    for path in files:
        rel = normalize_rel(path, project_root)
        rel_lower = rel.lower()
        name = path.name
        suffix = path.suffix.lower()

        if name.startswith("~$"):
            add_issue(
                issues,
                code="excel-lock-file",
                severity="high",
                path=rel,
                evidence=name,
                recommended_action="Close Excel and remove the transient lock file before packaging.",
            )

        if suffix in OFFICE_EXTENSIONS:
            if rel in ALLOWED_OFFICE_FILES:
                allowed_office_seen.append(rel)
            else:
                add_issue(
                    issues,
                    code="unexpected-office-workbook",
                    severity="high",
                    path=rel,
                    evidence=f"{suffix} file, {path.stat().st_size} bytes",
                    recommended_action="Move customer or ad hoc workbooks outside the plugin; keep only documented generic fixtures.",
                )

        if suffix in {".pyc", ".pyo"}:
            add_issue(
                issues,
                code="python-bytecode-artifact",
                severity="high",
                path=rel,
                evidence=suffix,
                recommended_action="Delete compiled Python artifacts; the release gate should run with source files only.",
            )

        if any(part == "__pycache__" for part in path.relative_to(project_root).parts):
            add_issue(
                issues,
                code="python-cache-artifact",
                severity="high",
                path=rel,
                evidence="__pycache__",
                recommended_action="Remove Python cache directories before packaging.",
            )

        for marker in customer_markers():
            if marker.lower() in rel_lower:
                add_issue(
                    issues,
                    code="sensitive-filename-marker",
                    severity="high",
                    path=rel,
                    evidence=marker,
                    recommended_action="Rename or remove customer-specific files from the reusable plugin package.",
                )

        if re.search(r"(^|/)(excel_bi_)?release_gate.*\.(json|md)$", rel_lower):
            add_issue(
                issues,
                code="generated-release-report",
                severity="high",
                path=rel,
                evidence=name,
                recommended_action="Write machine-specific release reports to a temp directory, not into the plugin package.",
            )

    return issues, sorted(allowed_office_seen)


def scan_text_markers(project_root: Path, files: list[Path]) -> list[dict[str, object]]:
    issues: list[dict[str, object]] = []
    marker_sets = [
        ("sensitive-content-marker", customer_markers(), "Remove customer-specific text from reusable plugin sources."),
        ("local-customer-path-marker", local_path_markers(), "Move local screenshots, temp paths, or customer paths out of the plugin package."),
    ]

    for path in files:
        suffix = path.suffix.lower()
        if suffix not in TEXT_EXTENSIONS:
            continue
        rel = normalize_rel(path, project_root)
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        except OSError as exc:
            add_issue(
                issues,
                code="text-read-error",
                severity="low",
                path=rel,
                evidence=str(exc),
                recommended_action="Inspect the file encoding if the release audit needs to scan it.",
            )
            continue

        lines = text.splitlines()
        for line_no, line in enumerate(lines, start=1):
            line_lower = line.lower()
            for code, markers, recommendation in marker_sets:
                for marker in markers:
                    if marker.lower() in line_lower:
                        add_issue(
                            issues,
                            code=code,
                            severity="high",
                            path=rel,
                            line=line_no,
                            evidence=marker,
                            recommended_action=recommendation,
                        )
                        break
                if issues and issues[-1].get("path") == rel and issues[-1].get("line") == line_no and issues[-1].get("code") == code:
                    break

    return issues


def build_report(project_root: Path) -> dict[str, object]:
    files = all_files(project_root)
    inventory_issues, allowed_office_seen = scan_file_inventory(project_root, files)
    text_issues = scan_text_markers(project_root, files)
    issues = inventory_issues + text_issues
    high_count = sum(1 for issue in issues if issue.get("severity") == "high")
    office_files = [normalize_rel(path, project_root) for path in files if path.suffix.lower() in OFFICE_EXTENSIONS]

    return {
        "status": "pass" if not issues else "fail",
        "generatedAt": now_iso(),
        "projectRoot": str(project_root),
        "summary": {
            "fileCount": len(files),
            "officeFileCount": len(office_files),
            "allowedOfficeFileCount": len(allowed_office_seen),
            "issueCount": len(issues),
            "highIssueCount": high_count,
            "maxSeverity": "high" if high_count else ("low" if issues else "none"),
        },
        "allowlists": {
            "officeFiles": sorted(ALLOWED_OFFICE_FILES),
        },
        "observed": {
            "officeFiles": sorted(office_files),
            "allowedOfficeFiles": allowed_office_seen,
        },
        "issues": issues,
    }


def markdown_report(report: dict[str, object]) -> str:
    summary = report["summary"]
    lines = [
        "# Artifact Hygiene Report",
        "",
        f"status: **{report['status']}**",
        "",
        "## Summary",
        "",
        "| Metric | Value |",
        "|---|---:|",
        f"| Files scanned | {summary['fileCount']} |",
        f"| Office workbooks | {summary['officeFileCount']} |",
        f"| Allowed Office fixtures | {summary['allowedOfficeFileCount']} |",
        f"| Issues | {summary['issueCount']} |",
        f"| High severity issues | {summary['highIssueCount']} |",
        "",
        "## Allowed Office Fixtures",
        "",
    ]

    for item in report["allowlists"]["officeFiles"]:
        observed = item in report["observed"]["allowedOfficeFiles"]
        lines.append(f"- `{item}` ({'present' if observed else 'not present'})")

    lines.extend(["", "## Issues", ""])
    issues = report["issues"]
    if not issues:
        lines.append("No artifact hygiene issues found.")
    else:
        lines.extend([
            "| Severity | Code | Path | Evidence | Recommended Action |",
            "|---|---|---|---|---|",
        ])
        for issue in issues:
            path = issue["path"]
            if "line" in issue:
                path = f"{path}:{issue['line']}"
            lines.append(
                "| {severity} | `{code}` | `{path}` | `{evidence}` | {action} |".format(
                    severity=issue["severity"],
                    code=issue["code"],
                    path=path,
                    evidence=str(issue["evidence"]).replace("|", "\\|"),
                    action=str(issue["recommendedAction"]).replace("|", "\\|"),
                )
            )

    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--out-json", default="", help="Write JSON report")
    parser.add_argument("--out-md", default="", help="Write Markdown report")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero when hygiene issues are found")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    report = build_report(project_root)

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
        "Artifact hygiene {status}: files={files}, office={office}, issues={issues}".format(
            status=report["status"],
            files=summary["fileCount"],
            office=summary["officeFileCount"],
            issues=summary["issueCount"],
        )
    )

    if args.require_pass and report["status"] != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
