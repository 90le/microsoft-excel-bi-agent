#!/usr/bin/env python3
"""Validate cross-agent SKILL.md directories without external dependencies."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ALLOWED_KEYS = {"name", "description", "license", "allowed-tools", "metadata"}
NAME_RE = re.compile(r"^[a-z0-9-]+$")
LINK_RE = re.compile(r"\[[^\]]+\]\(([^)]+)\)")
OPENAI_INTERFACE_KEY_RE = re.compile(r'^  (display_name|short_description|default_prompt):\s+"([^"]+)"\s*$')
PLACEHOLDER_RE = re.compile(
    "|".join(
        [
            r"\b" + "TO" + "DO" + r"\b",
            r"\[" + "TO" + "DO" + r"\]",
            "\u5f85\u8865",
            "\u5360\u4f4d",
            "x" + "xx",
        ]
    ),
    re.IGNORECASE,
)


def validate_openai_yaml(skill_dir: Path, skill_name: str) -> list[str]:
    errors: list[str] = []
    path = skill_dir / "agents" / "openai.yaml"
    if not path.exists():
        return [f"{path}: missing agents/openai.yaml"]

    text = path.read_text(encoding="utf-8")
    if PLACEHOLDER_RE.search(text):
        errors.append(f"{path}: placeholder text found")
    if not text.startswith("interface:\n"):
        errors.append(f"{path}: missing top-level interface block")

    values: dict[str, str] = {}
    unexpected_interface_lines: list[str] = []
    for line in text.splitlines()[1:]:
        if not line.strip():
            continue
        if line.startswith("  "):
            match = OPENAI_INTERFACE_KEY_RE.match(line)
            if not match:
                unexpected_interface_lines.append(line)
                continue
            values[match.group(1)] = match.group(2).strip()

    for key in ("display_name", "short_description", "default_prompt"):
        if not values.get(key):
            errors.append(f"{path}: missing interface.{key}")

    default_prompt = values.get("default_prompt", "")
    if f"${skill_name}" not in default_prompt:
        errors.append(f"{path}: default_prompt must mention ${skill_name}")

    for key, value in values.items():
        if len(value) > 180:
            errors.append(f"{path}: interface.{key} too long ({len(value)})")

    if unexpected_interface_lines:
        preview = "; ".join(unexpected_interface_lines[:3])
        errors.append(f"{path}: unsupported or unquoted interface line(s): {preview}")

    return errors


def parse_frontmatter(text: str) -> tuple[dict[str, str], str | None]:
    if not text.startswith("---\n"):
        return {}, "missing YAML frontmatter"
    match = re.match(r"^---\n(.*?)\n---", text, re.S)
    if not match:
        return {}, "invalid YAML frontmatter delimiters"
    data: dict[str, str] = {}
    for line in match.group(1).splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if ":" not in stripped:
            return {}, f"invalid frontmatter line: {line}"
        key, value = stripped.split(":", 1)
        data[key.strip()] = value.strip().strip('"').strip("'")
    return data, None


def validate_skill(skill_dir: Path) -> list[str]:
    errors: list[str] = []
    skill_md = skill_dir / "SKILL.md"
    if not skill_md.exists():
        return [f"{skill_dir}: missing SKILL.md"]

    text = skill_md.read_text(encoding="utf-8")
    frontmatter, error = parse_frontmatter(text)
    if error:
        errors.append(f"{skill_md}: {error}")
        return errors

    unexpected = set(frontmatter) - ALLOWED_KEYS
    if unexpected:
        errors.append(f"{skill_md}: unexpected frontmatter keys: {', '.join(sorted(unexpected))}")

    name = frontmatter.get("name", "").strip()
    description = frontmatter.get("description", "").strip()
    if not name:
        errors.append(f"{skill_md}: missing name")
    elif not NAME_RE.match(name):
        errors.append(f"{skill_md}: invalid name {name!r}")
    elif name.startswith("-") or name.endswith("-") or "--" in name:
        errors.append(f"{skill_md}: invalid hyphen placement in name {name!r}")
    elif name != skill_dir.name:
        errors.append(f"{skill_md}: name {name!r} does not match directory {skill_dir.name!r}")
    if len(name) > 64:
        errors.append(f"{skill_md}: name too long")

    if not description:
        errors.append(f"{skill_md}: missing description")
    if len(description) > 1024:
        errors.append(f"{skill_md}: description too long ({len(description)})")
    if "<" in description or ">" in description:
        errors.append(f"{skill_md}: description contains angle brackets")

    if name:
        errors.extend(validate_openai_yaml(skill_dir, name))

    if PLACEHOLDER_RE.search(text):
        errors.append(f"{skill_md}: placeholder text found")

    for link in LINK_RE.findall(text):
        if re.match(r"^[a-z]+://", link) or link.startswith("#") or link.startswith("mailto:"):
            continue
        target = (skill_dir / link).resolve()
        if not target.exists():
            errors.append(f"{skill_md}: missing relative link target {link}")

    for md in skill_dir.rglob("*.md"):
        if md == skill_md:
            continue
        body = md.read_text(encoding="utf-8")
        if PLACEHOLDER_RE.search(body):
            errors.append(f"{md}: placeholder text found")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", default=".", help="Project root or skills directory")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    skills_dir = root / ".agents" / "skills"
    if not skills_dir.exists():
        skills_dir = root

    skill_dirs = sorted(p for p in skills_dir.iterdir() if p.is_dir() and (p / "SKILL.md").exists())
    if not skill_dirs:
        print(f"No skills found under {skills_dir}", file=sys.stderr)
        return 1

    errors: list[str] = []
    for skill_dir in skill_dirs:
        errors.extend(validate_skill(skill_dir))

    if errors:
        print("Validation failed:")
        for error in errors:
            print(f"- {error}")
        return 1

    print(f"Validation OK: {len(skill_dirs)} skills")
    for skill_dir in skill_dirs:
        print(f"- {skill_dir.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
