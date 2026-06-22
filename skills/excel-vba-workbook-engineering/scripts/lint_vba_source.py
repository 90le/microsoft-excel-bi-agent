#!/usr/bin/env python3
"""Lint exported or drafted VBA source before importing it into Excel.

This is a static source check. It does not compile VBA, inspect references, run
macros, validate button clicks, or replace Windows Excel COM validation.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


VBA_EXTENSIONS = {".bas", ".cls", ".frm"}
PROC_START_RE = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"([A-Za-z_][A-Za-z0-9_]*)\b",
    re.IGNORECASE,
)
PROC_END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
ATTRIBUTE_RE = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.IGNORECASE)
OPTION_EXPLICIT_RE = re.compile(r"^\s*Option\s+Explicit\b", re.IGNORECASE)


@dataclass
class Procedure:
    module: str
    name: str
    kind: str
    visibility: str
    line: int


def read_text(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8-replace"


def strip_comment(line: str) -> str:
    in_string = False
    index = 0
    while index < len(line):
        char = line[index]
        if char == '"':
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif char == "'" and not in_string:
            return line[:index]
        index += 1
    return line


def normalize_proc_kind(kind: str) -> str:
    upper = kind.upper()
    if upper.startswith("PROPERTY"):
        return "PROPERTY"
    return upper


def expected_end_kind(kind: str) -> str:
    return "PROPERTY" if kind == "PROPERTY" else kind


def vba_files(source: Path) -> list[Path]:
    if source.is_file():
        return [source] if source.suffix.lower() in VBA_EXTENSIONS else []
    return sorted(path for path in source.rglob("*") if path.is_file() and path.suffix.lower() in VBA_EXTENSIONS)


def inspect_module(path: Path) -> dict[str, Any]:
    text, encoding = read_text(path)
    lines = text.splitlines()
    errors: list[str] = []
    warnings: list[str] = []
    procedures: list[Procedure] = []
    stack: list[Procedure] = []
    has_option_explicit = False
    declared_name = ""
    seen_code = False

    for line_number, line in enumerate(lines, start=1):
        code = strip_comment(line).strip()
        if not code:
            continue

        attr_match = ATTRIBUTE_RE.match(code)
        if attr_match:
            declared_name = attr_match.group(1)
            if seen_code:
                warnings.append(f"{path.name}:{line_number}: Attribute VB_Name appears after code")
            continue

        if OPTION_EXPLICIT_RE.match(code):
            has_option_explicit = True
            seen_code = True
            continue

        start_match = PROC_START_RE.match(code)
        end_match = PROC_END_RE.match(code)

        if start_match:
            visibility = (start_match.group(1) or "Public").title()
            kind = normalize_proc_kind(start_match.group(2))
            name = start_match.group(3)
            proc = Procedure(module=path.name, name=name, kind=kind, visibility=visibility, line=line_number)
            if stack:
                errors.append(f"{path.name}:{line_number}: procedure {name} starts before {stack[-1].name} ends")
            stack.append(proc)
            procedures.append(proc)
            seen_code = True
            continue

        if end_match:
            end_kind = normalize_proc_kind(end_match.group(1))
            if not stack:
                errors.append(f"{path.name}:{line_number}: End {end_kind.title()} without a matching procedure")
            else:
                proc = stack.pop()
                expected = expected_end_kind(proc.kind)
                if expected != end_kind:
                    errors.append(f"{path.name}:{line_number}: End {end_kind.title()} closes {proc.kind.title()} {proc.name}")
            seen_code = True
            continue

        seen_code = True

    for proc in stack:
        errors.append(f"{path.name}:{proc.line}: procedure {proc.name} has no matching End {expected_end_kind(proc.kind).title()}")

    if not has_option_explicit:
        warnings.append(f"{path.name}: missing Option Explicit")

    if declared_name and declared_name.lower() != path.stem.lower() and path.suffix.lower() in {".bas", ".cls"}:
        warnings.append(f"{path.name}: Attribute VB_Name '{declared_name}' differs from file stem '{path.stem}'")

    return {
        "path": str(path),
        "file": path.name,
        "extension": path.suffix.lower(),
        "encoding": encoding,
        "declaredName": declared_name,
        "hasOptionExplicit": has_option_explicit,
        "procedures": [proc.__dict__ for proc in procedures],
        "errors": errors,
        "warnings": warnings,
    }


def lint_source(source: Path, strict_option_explicit: bool, allow_duplicate_public: bool) -> dict[str, Any]:
    modules = [inspect_module(path) for path in vba_files(source)]
    errors: list[str] = []
    warnings: list[str] = []
    public_entries: list[dict[str, Any]] = []
    public_name_map: dict[str, list[dict[str, Any]]] = {}

    for module in modules:
        errors.extend(module["errors"])
        module_warnings = list(module["warnings"])
        if strict_option_explicit:
            strict_errors = [item for item in module_warnings if item.endswith("missing Option Explicit")]
            errors.extend(strict_errors)
            module_warnings = [item for item in module_warnings if item not in strict_errors]
        warnings.extend(module_warnings)

        if module["extension"] != ".bas":
            continue
        for proc in module["procedures"]:
            if proc["kind"] == "SUB" and proc["visibility"] != "Private":
                public_entries.append(proc)
                public_name_map.setdefault(proc["name"].lower(), []).append(proc)

    if not allow_duplicate_public:
        for entries in public_name_map.values():
            if len(entries) > 1:
                locations = ", ".join(f"{entry['module']}:{entry['line']}" for entry in entries)
                errors.append(f"duplicate public Sub '{entries[0]['name']}' in standard modules: {locations}")

    return {
        "source": str(source),
        "moduleCount": len(modules),
        "procedureCount": sum(len(module["procedures"]) for module in modules),
        "publicEntryCount": len(public_entries),
        "publicEntries": public_entries,
        "modules": modules,
        "errors": errors,
        "warnings": warnings,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="VBA source file or directory containing .bas/.cls/.frm files")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--strict-option-explicit", action="store_true", help="Treat missing Option Explicit as an error")
    parser.add_argument("--allow-duplicate-public", action="store_true", help="Allow duplicate public Sub names in standard modules")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.expanduser().resolve()
    report = lint_source(source, args.strict_option_explicit, args.allow_duplicate_public)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if report["errors"]:
        print("VBA source lint failed:")
        for error in report["errors"]:
            print(f"- {error}")
        return 1

    print(
        "VBA source lint OK: "
        f"{report['moduleCount']} modules, "
        f"{report['procedureCount']} procedures, "
        f"{report['publicEntryCount']} public entry subs"
    )
    for warning in report["warnings"]:
        print(f"- warning: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
