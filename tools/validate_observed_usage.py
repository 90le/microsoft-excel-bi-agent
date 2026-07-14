#!/usr/bin/env python3
"""Validate local, sanitized observed-usage JSONL evidence without external access."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator


REQUIRED_FIELDS = (
    "schemaVersion",
    "eventId",
    "recordedAt",
    "caseId",
    "requestedSkill",
    "selectedSkill",
    "outcome",
    "durationMs",
    "evidenceLevel",
    "evidenceBoundary",
)
EVIDENCE_LEVELS = {"structural", "runtime-capability", "workbook-behavior"}
WORKBOOK_BEHAVIOR_BOUNDARY = "local-user-supplied-sanitized"
UTC_TIMESTAMP_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$")
PATH_RE = re.compile(r"(?:[A-Za-z]:|[\\/]{2}|(?:^|\s)[~/]|[\\/])")
CREDENTIAL_RE = re.compile(
    r"(?:api[_-]?key|access[_-]?key|password|passwd|secret|token|authorization)\s*(?::|=|\s+)\S+|"
    r"\bbearer\s+\S+|\bAKIA[0-9A-Z]{16}\b|\bsk-[A-Za-z0-9_-]{16,}\b|"
    r"\bghp_[A-Za-z0-9]{20,}\b|\bxoxb-[A-Za-z0-9-]{20,}\b|"
    r"\beyJ[A-Za-z0-9_-]{5,}\.[A-Za-z0-9_-]{5,}\.[A-Za-z0-9_-]{5,}\b",
    re.IGNORECASE,
)
CREDENTIAL_KEY_RE = re.compile(
    r"^(?:api[_-]?key|access[_-]?key|password|passwd|secret|token|authorization)$",
    re.IGNORECASE,
)
CUSTOMER_ARTIFACT_RE = re.compile(
    r"\b[^\\/\s]+\.(?:xlsx|xlsm|xlsb|xls|csv|tsv)\b", re.IGNORECASE
)


def string_is_sensitive(value: str, *, is_key: bool = False) -> bool:
    return bool(
        PATH_RE.search(value)
        or CREDENTIAL_RE.search(value)
        or (is_key and CREDENTIAL_KEY_RE.search(value))
        or CUSTOMER_ARTIFACT_RE.search(value)
    )


def iter_strings(value: Any, field: str = "") -> Iterator[tuple[str, str, bool]]:
    """Yield every JSON string value with an actionable dotted field location."""
    if isinstance(value, str):
        yield field, value, False
    elif isinstance(value, dict):
        for key, child in value.items():
            key_text = str(key)
            key_segment = "<key>" if string_is_sensitive(key_text, is_key=True) else key_text
            child_field = f"{field}.{key_segment}" if field else key_segment
            key_field = f"{field}.<key>" if field else "<key>"
            yield key_field, key_text, True
            yield from iter_strings(child, child_field)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            child_field = f"{field}[{index}]"
            yield from iter_strings(child, child_field)


def is_utc_timestamp(value: object) -> bool:
    if not isinstance(value, str) or not UTC_TIMESTAMP_RE.fullmatch(value):
        return False
    try:
        datetime.fromisoformat(value[:-1] + "+00:00")
    except ValueError:
        return False
    return True


def validate_event(event: object, seen_event_ids: set[str]) -> list[tuple[str, str]]:
    """Return (field, message) pairs for one parsed JSONL event."""
    errors: list[tuple[str, str]] = []
    if not isinstance(event, dict):
        return [("$", "event must be a JSON object")]

    for field in REQUIRED_FIELDS:
        if field not in event:
            errors.append((field, "missing required field"))

    if event.get("schemaVersion") != "1.0":
        errors.append(("schemaVersion", "must equal '1.0'"))

    for field in REQUIRED_FIELDS:
        if field in {"durationMs", "schemaVersion"} or field not in event:
            continue
        if not isinstance(event[field], str) or not event[field].strip():
            errors.append((field, "must be a non-empty string"))

    event_id = event.get("eventId")
    if isinstance(event_id, str) and event_id.strip():
        if event_id in seen_event_ids:
            errors.append(("eventId", "duplicate eventId"))
        else:
            seen_event_ids.add(event_id)

    if not is_utc_timestamp(event.get("recordedAt")):
        errors.append(("recordedAt", "must be a UTC ISO-8601 timestamp ending in Z"))

    duration = event.get("durationMs")
    if isinstance(duration, bool) or not isinstance(duration, int) or duration < 0:
        errors.append(("durationMs", "must be a nonnegative integer"))

    evidence_level = event.get("evidenceLevel")
    if evidence_level not in EVIDENCE_LEVELS:
        errors.append(("evidenceLevel", "unsupported value; use structural, runtime-capability, or workbook-behavior"))
    elif evidence_level == "workbook-behavior" and event.get("evidenceBoundary") != WORKBOOK_BEHAVIOR_BOUNDARY:
        errors.append(
            ("evidenceBoundary", "workbook-behavior requires local-user-supplied-sanitized")
        )

    for field, value, is_key in iter_strings(event):
        if PATH_RE.search(value):
            errors.append((field, "contains path-like content"))
        if CREDENTIAL_RE.search(value) or (is_key and CREDENTIAL_KEY_RE.search(value)):
            errors.append((field, "contains credential-like content"))
        if CUSTOMER_ARTIFACT_RE.search(value):
            errors.append((field, "contains a customer artifact name"))
    return errors


def validate_file(path: Path, seen_event_ids: set[str]) -> list[str]:
    """Validate one local JSONL file and return file/line/field diagnostics."""
    diagnostics: list[str] = []
    resolved_path = path.expanduser().resolve()
    try:
        lines = resolved_path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        return [f"{resolved_path}: unable to read local JSONL file: {exc}"]
    except UnicodeError as exc:
        return [f"{resolved_path}: invalid UTF-8 JSONL text: {exc}"]

    for line_number, line in enumerate(lines, start=1):
        if not line.strip():
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError as exc:
            diagnostics.append(f"{resolved_path}:{line_number}: $: malformed JSON ({exc.msg})")
            continue
        for field, message in validate_event(event, seen_event_ids):
            diagnostics.append(f"{resolved_path}:{line_number}: {field}: {message}")
    return diagnostics


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("jsonl_files", nargs="+", type=Path, help="Local JSONL evidence files to validate")
    args = parser.parse_args(argv)

    seen_event_ids: set[str] = set()
    diagnostics: list[str] = []
    for path in args.jsonl_files:
        diagnostics.extend(validate_file(path, seen_event_ids))

    if diagnostics:
        print("Observed usage validation failed:", file=sys.stderr)
        for diagnostic in diagnostics:
            print(f"- {diagnostic}", file=sys.stderr)
        return 1
    print(f"Observed usage validation passed for {len(args.jsonl_files)} local JSONL file(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
