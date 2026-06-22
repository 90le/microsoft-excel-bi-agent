#!/usr/bin/env python3
"""Helpers for parsing MDX bracketed references in Excel CUBE formulas."""

from __future__ import annotations

from typing import Iterator, TypedDict


class MdxPath(TypedDict):
    raw: str
    start: int
    end: int
    parts: list[str]
    spans: list[tuple[int, int]]


def parse_mdx_bracket_identifier(text: str, start: int) -> tuple[str, int] | None:
    """Parse one MDX bracket identifier, including escaped closing brackets."""
    if start >= len(text) or text[start] != "[":
        return None
    index = start + 1
    chars: list[str] = []
    while index < len(text):
        char = text[index]
        if char == "]":
            if index + 1 < len(text) and text[index + 1] == "]":
                chars.append("]")
                index += 2
                continue
            return "".join(chars), index + 1
        chars.append(char)
        index += 1
    return None


def escape_mdx_identifier(value: str) -> str:
    return value.replace("]", "]]")


def iter_mdx_bracket_paths(formula: str) -> Iterator[MdxPath]:
    index = 0
    while index < len(formula):
        start = formula.find("[", index)
        if start < 0:
            break

        parts: list[str] = []
        spans: list[tuple[int, int]] = []
        position = start
        while position < len(formula) and formula[position] == "[":
            parsed = parse_mdx_bracket_identifier(formula, position)
            if not parsed:
                break
            name, end = parsed
            parts.append(name)
            spans.append((position, end))
            position = end
            if position + 1 < len(formula) and formula[position] == "." and formula[position + 1] == "[":
                position += 1
                continue
            break

        if len(parts) >= 2:
            yield {
                "raw": formula[start:position],
                "start": start,
                "end": position,
                "parts": parts,
                "spans": spans,
            }
            index = position
        else:
            index = start + 1


def cube_measure_refs(formula: str) -> list[str]:
    refs = {
        path["parts"][1]
        for path in iter_mdx_bracket_paths(formula)
        if len(path["parts"]) >= 2 and path["parts"][0].lower() == "measures"
    }
    return sorted(refs, key=str.lower)


def member_refs(formula: str) -> list[str]:
    refs = {
        path["raw"]
        for path in iter_mdx_bracket_paths(formula)
        if path["parts"][0].lower() != "measures"
    }
    return sorted(refs, key=str.lower)


def replace_cube_measure_ref(formula: str, old: str, new: str) -> tuple[str, bool]:
    old_key = old.lower()
    replacements: list[tuple[int, int, str]] = []
    for path in iter_mdx_bracket_paths(formula):
        parts = path["parts"]
        if len(parts) < 2 or parts[0].lower() != "measures":
            continue
        if parts[1].lower() != old_key:
            continue
        start = path["spans"][0][0]
        end = path["spans"][1][1]
        replacements.append((start, end, f"[Measures].[{escape_mdx_identifier(new)}]"))

    if not replacements:
        return formula, False

    output: list[str] = []
    last = 0
    for start, end, replacement in replacements:
        if start < last:
            continue
        output.append(formula[last:start])
        output.append(replacement)
        last = end
    output.append(formula[last:])
    return "".join(output), True
