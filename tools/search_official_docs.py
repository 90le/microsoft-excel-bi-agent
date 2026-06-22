#!/usr/bin/env python3
"""Search bundled official-documentation indexes across plugin skills."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def candidate_index_roots(project_root: Path) -> list[Path]:
    return [
        project_root / ".agents" / "skills",
        project_root / "skills",
    ]


def find_indexes(project_root: Path) -> list[Path]:
    indexes: list[Path] = []
    for root in candidate_index_roots(project_root):
        if not root.is_dir():
            continue
        indexes.extend(sorted(root.glob("*/references/official-docs-index.json")))
    seen: set[Path] = set()
    unique: list[Path] = []
    for path in indexes:
        resolved = path.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        unique.append(path)
    return unique


def prefer_canonical_entries(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    selected: dict[tuple[str, str], dict[str, Any]] = {}
    for entry in entries:
        key = (str(entry.get("skill", "")), str(entry.get("id", "")))
        current = selected.get(key)
        if current is None:
            selected[key] = entry
            continue
        current_path = str(current.get("indexPath", "")).replace("\\", "/")
        candidate_path = str(entry.get("indexPath", "")).replace("\\", "/")
        if "/.agents/skills/" in candidate_path and "/.agents/skills/" not in current_path:
            selected[key] = entry
    return list(selected.values())


def load_entries(index_path: Path) -> list[dict[str, Any]]:
    data = json.loads(index_path.read_text(encoding="utf-8"))
    entries = data.get("entries")
    if not isinstance(entries, list):
        return []
    skill_name = index_path.parents[1].name
    result: list[dict[str, Any]] = []
    for entry in entries:
        if isinstance(entry, dict):
            item = dict(entry)
            item["skill"] = skill_name
            item["indexPath"] = str(index_path)
            result.append(item)
    return result


def entry_text(entry: dict[str, Any]) -> str:
    parts = [
        str(entry.get("skill", "")),
        str(entry.get("id", "")),
        str(entry.get("category", "")),
        str(entry.get("title", "")),
        str(entry.get("official_url", "")),
        str(entry.get("online_query", "")),
        str(entry.get("use_when", "")),
    ]
    parts.extend(str(item) for item in entry.get("local_keywords", []))
    return " ".join(parts).lower()


def score_entry(entry: dict[str, Any], terms: list[str]) -> int:
    text = entry_text(entry)
    score = 0
    for term in terms:
        if term in text:
            score += 1
        if term == str(entry.get("id", "")).lower():
            score += 3
        if term == str(entry.get("skill", "")).lower():
            score += 3
        if term == str(entry.get("category", "")).lower():
            score += 2
        for keyword in entry.get("local_keywords", []):
            if term == str(keyword).lower():
                score += 2
    return score


def render_markdown(entries: list[dict[str, Any]]) -> str:
    if not entries:
        return "No matching official documentation entries found."
    lines = ["# Official Documentation Matches", ""]
    for entry in entries:
        lines.extend(
            [
                f"## {entry.get('title', entry.get('id'))}",
                "",
                f"- skill: `{entry.get('skill', '')}`",
                f"- id: `{entry.get('id', '')}`",
                f"- category: `{entry.get('category', '')}`",
                f"- official_url: {entry.get('official_url', '')}",
                f"- use_when: {entry.get('use_when', '')}",
                f"- online_query: `{entry.get('online_query', '')}`",
                "",
            ]
        )
    return "\n".join(lines).rstrip()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("query", nargs="+", help="Search terms, for example CALCULATE or CUBEVALUE")
    parser.add_argument(
        "--project-root",
        help="Plugin project root. Defaults to the parent of this script's tools directory.",
    )
    parser.add_argument("--limit", type=int, default=10, help="Maximum number of matches")
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.project_root:
        project_root = Path(args.project_root).expanduser().resolve()
    else:
        project_root = Path(__file__).resolve().parents[1]
    terms = [term.lower() for term in " ".join(args.query).split() if term.strip()]
    entries: list[dict[str, Any]] = []
    for index_path in find_indexes(project_root):
        entries.extend(load_entries(index_path))
    entries = prefer_canonical_entries(entries)
    ranked = [(score_entry(entry, terms), entry) for entry in entries]
    matches = [
        entry
        for score, entry in sorted(
            ranked,
            key=lambda item: (-item[0], item[1].get("skill", ""), item[1].get("id", "")),
        )
        if score > 0
    ][: max(args.limit, 1)]

    if args.json:
        print(json.dumps(matches, ensure_ascii=False, indent=2))
    else:
        print(render_markdown(matches))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
