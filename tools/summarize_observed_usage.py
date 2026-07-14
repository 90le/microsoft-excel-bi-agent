#!/usr/bin/env python3
"""Summarize validated local observed-usage JSONL evidence without external access."""

from __future__ import annotations

import argparse
import json
import statistics
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from validate_observed_usage import validate_file


def sorted_counts(values: list[str]) -> dict[str, int]:
    """Return deterministic grouping counts for sanitized event values."""
    return dict(sorted(Counter(values).items()))


def summarize(events: list[dict[str, Any]]) -> dict[str, object]:
    """Build an aggregate-only report from already validated events."""
    durations = [event["durationMs"] for event in events]
    median_duration = statistics.median(durations) if durations else None
    if isinstance(median_duration, float) and median_duration.is_integer():
        median_duration = int(median_duration)

    return {
        "eventCount": len(events),
        "requestedSkillCounts": sorted_counts([event["requestedSkill"] for event in events]),
        "selectedSkillCounts": sorted_counts([event["selectedSkill"] for event in events]),
        "outcomeCounts": sorted_counts([event["outcome"] for event in events]),
        "evidenceLevelCounts": sorted_counts([event["evidenceLevel"] for event in events]),
        "durationMs": {
            "median": median_duration,
            "max": max(durations) if durations else None,
        },
    }


def read_validated_events(paths: list[Path]) -> list[dict[str, Any]] | None:
    """Validate all inputs before reading any event for aggregation."""
    seen_event_ids: set[str] = set()
    diagnostics: list[str] = []
    for path in paths:
        diagnostics.extend(validate_file(path, seen_event_ids))
    if diagnostics:
        print("Observed usage validation failed; no summary was produced.", file=sys.stderr)
        return None

    events: list[dict[str, Any]] = []
    for path in paths:
        for line in path.expanduser().resolve().read_text(encoding="utf-8").splitlines():
            if line.strip():
                events.append(json.loads(line))
    return events


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("jsonl_files", nargs="+", type=Path, help="Local JSONL evidence files to summarize")
    args = parser.parse_args(argv)

    events = read_validated_events(args.jsonl_files)
    if events is None:
        return 1
    print(json.dumps(summarize(events), ensure_ascii=False, sort_keys=True, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
