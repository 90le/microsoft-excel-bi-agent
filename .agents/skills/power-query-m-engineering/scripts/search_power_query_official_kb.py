#!/usr/bin/env python3
"""Search the bundled Power Query official-documentation index."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def default_index_path() -> Path:
    return Path(__file__).resolve().parents[1] / "references" / "official-docs-index.json"


def entry_text(entry: dict[str, Any]) -> str:
    parts: list[str] = [
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
        if term == str(entry.get("category", "")).lower():
            score += 2
        for keyword in entry.get("local_keywords", []):
            if term == str(keyword).lower():
                score += 2
    return score


def load_entries(index_path: Path) -> list[dict[str, Any]]:
    with index_path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    entries = data.get("entries")
    if not isinstance(entries, list):
        raise ValueError(f"Invalid index: {index_path}")
    return entries


def render_markdown(entries: list[dict[str, Any]]) -> str:
    if not entries:
        return "No matching official documentation entries found."

    lines = ["# Power Query Official Documentation Matches", ""]
    for entry in entries:
        lines.extend(
            [
                f"## {entry.get('title', entry.get('id'))}",
                "",
                f"- id: `{entry.get('id', '')}`",
                f"- category: `{entry.get('category', '')}`",
                f"- official_url: {entry.get('official_url', '')}",
                f"- use_when: {entry.get('use_when', '')}",
                f"- online_query: `{entry.get('online_query', '')}`",
                f"- local_keywords: {', '.join(entry.get('local_keywords', []))}",
                "",
            ]
        )
    return "\n".join(lines).rstrip()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("query", nargs="+", help="Search text, for example Table.NestedJoin or RefreshAll")
    parser.add_argument("--index", type=Path, default=default_index_path(), help="Path to official-docs-index.json")
    parser.add_argument("--limit", type=int, default=8, help="Maximum number of matches")
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of Markdown")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    terms = [term.lower() for term in " ".join(args.query).split() if term.strip()]
    entries = load_entries(args.index)
    ranked = [
        (score_entry(entry, terms), entry)
        for entry in entries
    ]
    matches = [
        entry for score, entry in sorted(ranked, key=lambda item: (-item[0], item[1].get("id", "")))
        if score > 0
    ][: max(args.limit, 1)]

    if args.json:
        print(json.dumps(matches, ensure_ascii=False, indent=2))
    else:
        print(render_markdown(matches))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
