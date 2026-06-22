#!/usr/bin/env python3
"""Validate bundled official-documentation indexes.

The default validation is offline and deterministic. It checks JSON shape,
required entry fields, official Microsoft URL domains, duplicate IDs, keyword
coverage, and representative search queries through search_official_docs.py.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


ALLOWED_DOC_HOSTS = {
    "learn.microsoft.com",
    "support.microsoft.com",
}

REQUIRED_ENTRY_FIELDS = {
    "id",
    "category",
    "title",
    "official_url",
    "local_keywords",
    "online_query",
    "use_when",
}

SEARCH_SMOKES = [
    ("Power Query refresh", "RefreshAll", "power-query-m-engineering"),
    ("DAX context", "CALCULATE", "power-pivot-dax-modeling"),
    ("Excel CUBE formula", "CUBEVALUE", "mdx-cubevalue-extraction"),
    ("ADO connection", "ADODB Connection", "excel-ado-sql-data-access"),
]

DEFAULT_ONLINE_TIMEOUT = 10.0
DEFAULT_USER_AGENT = "microsoft-excel-bi-agent-pack-doc-validator/0.1"


def find_indexes(project_root: Path) -> list[Path]:
    root = project_root / ".agents" / "skills"
    return sorted(root.glob("*/references/official-docs-index.json"))


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def is_allowed_official_url(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme == "https" and parsed.netloc.lower() in ALLOWED_DOC_HOSTS


def validate_index(path: Path, project_root: Path) -> tuple[list[str], dict[str, Any]]:
    errors: list[str] = []
    data = load_json(path)
    rel = path.relative_to(project_root).as_posix()
    skill = path.parents[1].name
    entries = data.get("entries")

    if not data.get("version"):
        errors.append(f"{rel}: missing version")
    if not data.get("source_policy"):
        errors.append(f"{rel}: missing source_policy")
    if not isinstance(entries, list) or not entries:
        errors.append(f"{rel}: entries must be a non-empty list")
        entries = []

    ids: list[str] = []
    categories: Counter[str] = Counter()
    for idx, entry in enumerate(entries, start=1):
        if not isinstance(entry, dict):
            errors.append(f"{rel}: entry {idx} must be an object")
            continue
        missing = sorted(REQUIRED_ENTRY_FIELDS - set(entry))
        if missing:
            errors.append(f"{rel}: entry {idx} missing fields: {', '.join(missing)}")
        entry_id = str(entry.get("id", "")).strip()
        if entry_id:
            ids.append(entry_id)
        else:
            errors.append(f"{rel}: entry {idx} has blank id")
        categories[str(entry.get("category", "")).strip()] += 1
        url = str(entry.get("official_url", "")).strip()
        if not is_allowed_official_url(url):
            errors.append(f"{rel}: entry {entry_id or idx} official_url is not an allowed Microsoft URL: {url}")
        keywords = entry.get("local_keywords")
        if not isinstance(keywords, list) or not any(str(item).strip() for item in keywords):
            errors.append(f"{rel}: entry {entry_id or idx} local_keywords must be a non-empty list")
        online_query = str(entry.get("online_query", "")).strip()
        if "site:learn.microsoft.com" not in online_query and "site:support.microsoft.com" not in online_query:
            errors.append(f"{rel}: entry {entry_id or idx} online_query must be restricted to Microsoft official domains")
        if not str(entry.get("use_when", "")).strip():
            errors.append(f"{rel}: entry {entry_id or idx} use_when is blank")

    duplicate_ids = sorted(item for item, count in Counter(ids).items() if count > 1)
    if duplicate_ids:
        errors.append(f"{rel}: duplicate ids: {', '.join(duplicate_ids)}")

    return errors, {
        "skill": skill,
        "index": rel,
        "entryCount": len(entries),
        "categoryCount": len([key for key in categories if key]),
        "officialHosts": sorted({urlparse(str(entry.get("official_url", ""))).netloc for entry in entries if isinstance(entry, dict)}),
    }


def iter_official_urls(indexes: list[Path], project_root: Path) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    seen: set[str] = set()
    for path in indexes:
        data = load_json(path)
        skill = path.parents[1].name
        for entry in data.get("entries", []):
            if not isinstance(entry, dict):
                continue
            url = str(entry.get("official_url", "")).strip()
            if not url or url in seen:
                continue
            seen.add(url)
            records.append(
                {
                    "skill": skill,
                    "index": path.relative_to(project_root).as_posix(),
                    "id": str(entry.get("id", "")).strip(),
                    "title": str(entry.get("title", "")).strip(),
                    "url": url,
                }
            )
    return records


def check_url(url: str, timeout: float, user_agent: str) -> tuple[bool, str, int | None]:
    headers = {"User-Agent": user_agent}
    last_error = ""
    for method in ("HEAD", "GET"):
        request = Request(url, headers=headers, method=method)
        try:
            with urlopen(request, timeout=timeout) as response:
                status = int(getattr(response, "status", response.getcode()))
                if 200 <= status < 400:
                    return True, f"{method} {status}", status
                last_error = f"{method} returned HTTP {status}"
        except HTTPError as exc:
            status = int(exc.code)
            if method == "HEAD" and status in {403, 405, 501}:
                last_error = f"{method} returned HTTP {status}; retrying GET"
                continue
            return False, f"{method} returned HTTP {status}", status
        except (URLError, TimeoutError, OSError) as exc:
            last_error = f"{method} failed: {exc}"
            if method == "HEAD":
                continue
            return False, last_error, None
    return False, last_error or "request failed", None


def run_online_url_checks(indexes: list[Path], project_root: Path, limit: int, timeout: float, user_agent: str) -> tuple[list[str], list[dict[str, Any]]]:
    errors: list[str] = []
    records = iter_official_urls(indexes, project_root)
    if limit > 0:
        records = records[:limit]

    results: list[dict[str, Any]] = []
    for record in records:
        ok, detail, status = check_url(record["url"], timeout, user_agent)
        result = dict(record)
        result.update({"ok": ok, "detail": detail, "status": status})
        results.append(result)
        if not ok:
            errors.append(f"{record['index']}: {record['id']} online URL check failed: {record['url']} ({detail})")
    return errors, results


def run_search_smoke(project_root: Path, query: str, expected_skill: str) -> tuple[bool, str]:
    command = [
        sys.executable,
        str(project_root / "tools" / "search_official_docs.py"),
        *query.split(),
        "--project-root",
        str(project_root),
        "--limit",
        "5",
        "--json",
    ]
    completed = subprocess.run(command, cwd=str(project_root), text=True, capture_output=True, timeout=60)
    if completed.returncode != 0:
        return False, f"{query}: search command failed with exit_code={completed.returncode}: {completed.stderr.strip()}"
    try:
        matches = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        return False, f"{query}: search output was not JSON: {exc}"
    if not matches:
        return False, f"{query}: no matches"
    if not any(item.get("skill") == expected_skill for item in matches if isinstance(item, dict)):
        found = sorted({str(item.get("skill", "")) for item in matches if isinstance(item, dict)})
        return False, f"{query}: expected skill {expected_skill}, found {found}"
    return True, f"{query}: matched {expected_skill}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--out-json", type=Path, help="Optional output JSON report")
    parser.add_argument("--check-online", action="store_true", help="Also check official_url reachability against Microsoft domains")
    parser.add_argument("--online-limit", type=int, default=0, help="Maximum unique URLs to check online; 0 checks all")
    parser.add_argument("--online-timeout", type=float, default=DEFAULT_ONLINE_TIMEOUT, help="Per-request timeout in seconds")
    parser.add_argument("--online-user-agent", default=DEFAULT_USER_AGENT, help="User-Agent for online URL checks")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    indexes = find_indexes(project_root)
    errors: list[str] = []
    summaries: list[dict[str, Any]] = []

    if not indexes:
        errors.append("No official-docs-index.json files found under .agents/skills")

    for path in indexes:
        index_errors, summary = validate_index(path, project_root)
        errors.extend(index_errors)
        summaries.append(summary)

    search_results: list[dict[str, Any]] = []
    for label, query, expected_skill in SEARCH_SMOKES:
        ok, detail = run_search_smoke(project_root, query, expected_skill)
        search_results.append({"label": label, "query": query, "expectedSkill": expected_skill, "ok": ok, "detail": detail})
        if not ok:
            errors.append(detail)

    online_results: list[dict[str, Any]] = []
    if args.check_online:
        online_errors, online_results = run_online_url_checks(indexes, project_root, args.online_limit, args.online_timeout, args.online_user_agent)
        errors.extend(online_errors)

    report = {
        "projectRoot": str(project_root),
        "indexCount": len(indexes),
        "entryCount": sum(int(item.get("entryCount", 0)) for item in summaries),
        "indexes": summaries,
        "searchSmokes": search_results,
        "onlineCheckEnabled": bool(args.check_online),
        "onlineChecks": online_results,
        "errors": errors,
    }

    if args.out_json:
        args.out_json.expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
        args.out_json.expanduser().resolve().write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if errors:
        print("Official documentation index validation failed:")
        for error in errors:
            print(f"- {error}")
        return 1

    print(f"Official documentation indexes OK: {report['indexCount']} indexes, {report['entryCount']} entries")
    for result in search_results:
        print(f"- {result['detail']}")
    if args.check_online:
        print(f"Online official URL checks OK: {len(online_results)} URLs")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
