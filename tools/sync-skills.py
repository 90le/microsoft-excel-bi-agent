#!/usr/bin/env python3
"""Sync canonical .agents/skills to agent-specific skill directories.

The source of truth is always:

    <project-root>/.agents/skills

Generated mirrors may include:

- <project-root>/skills for Codex plugin packaging.
- ~/.codex/skills for a user's local Codex skill install.
- <project-root>/.claude/skills for Claude-style project skills.
- <project-root>/.opencode/skills for OpenCode-style project skills.
- Any explicit --target directory.
"""

from __future__ import annotations

import argparse
import filecmp
import json
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


IGNORED_DIR_NAMES = {"__pycache__"}
IGNORED_FILE_SUFFIXES = {".pyc", ".pyo"}


@dataclass(frozen=True)
class SyncTarget:
    label: str
    path: Path


def skill_dirs(root: Path) -> list[Path]:
    if not root.is_dir():
        return []
    return sorted(path for path in root.iterdir() if path.is_dir() and (path / "SKILL.md").is_file())


def relative_files(root: Path) -> set[Path]:
    files: set[Path] = set()
    if not root.exists():
        return files
    for path in root.rglob("*"):
        if any(part in IGNORED_DIR_NAMES for part in path.relative_to(root).parts):
            continue
        if path.suffix in IGNORED_FILE_SUFFIXES:
            continue
        if path.is_file():
            files.add(path.relative_to(root))
    return files


def copy_ignore(_directory: str, names: list[str]) -> set[str]:
    ignored: set[str] = set()
    for name in names:
        if name in IGNORED_DIR_NAMES or Path(name).suffix in IGNORED_FILE_SUFFIXES:
            ignored.add(name)
    return ignored


def files_equal(left: Path, right: Path) -> bool:
    if not left.is_file() or not right.is_file():
        return False
    return filecmp.cmp(left, right, shallow=False)


def compare_trees(source: Path, target: Path) -> dict[str, list[str]]:
    source_files = relative_files(source)
    target_files = relative_files(target)
    missing = sorted(str(path).replace("\\", "/") for path in source_files - target_files)
    extra = sorted(str(path).replace("\\", "/") for path in target_files - source_files)
    changed = sorted(
        str(path).replace("\\", "/")
        for path in source_files & target_files
        if not files_equal(source / path, target / path)
    )
    return {"missing": missing, "extra": extra, "changed": changed}


def compare_user_skill_install(source: Path, target: Path) -> dict[str, list[str]]:
    """Compare only this package's skill directories inside a shared user skill root.

    User-level skill folders may already contain system or unrelated personal
    skills. Those are valid neighbors, not drift from this package.
    """
    source_files = relative_files(source)
    source_skill_names = {path.name for path in skill_dirs(source)}
    target_files = {
        path
        for path in relative_files(target)
        if path.parts and path.parts[0] in source_skill_names
    }
    missing = sorted(str(path).replace("\\", "/") for path in source_files - target_files)
    extra = sorted(str(path).replace("\\", "/") for path in target_files - source_files)
    changed = sorted(
        str(path).replace("\\", "/")
        for path in source_files & target_files
        if not files_equal(source / path, target / path)
    )
    return {"missing": missing, "extra": extra, "changed": changed}


def copy_skills(source: Path, target: Path, replace: bool) -> list[str]:
    target.mkdir(parents=True, exist_ok=True)
    copied: list[str] = []
    for skill_dir in skill_dirs(source):
        destination = target / skill_dir.name
        if destination.exists():
            if not replace:
                continue
            shutil.rmtree(destination)
        shutil.copytree(skill_dir, destination, ignore=copy_ignore)
        copied.append(skill_dir.name)
    return copied


def target_manifest(source: Path, targets: Iterable[SyncTarget]) -> dict[str, object]:
    skills = [path.name for path in skill_dirs(source)]
    return {
        "source": str(source),
        "skillCount": len(skills),
        "skills": skills,
        "targets": [{"label": target.label, "path": str(target.path)} for target in targets],
    }


def build_targets(args: argparse.Namespace, project_root: Path) -> list[SyncTarget]:
    targets: list[SyncTarget] = []
    if args.plugin_mirror:
        targets.append(SyncTarget("codex-plugin-mirror", project_root / "skills"))
    if args.codex_user:
        targets.append(SyncTarget("codex-user", Path.home() / ".codex" / "skills"))
    if args.claude_project:
        targets.append(SyncTarget("claude-project", project_root / ".claude" / "skills"))
    if args.opencode_project:
        targets.append(SyncTarget("opencode-project", project_root / ".opencode" / "skills"))
    for index, target in enumerate(args.target or [], start=1):
        targets.append(SyncTarget(f"custom-{index}", Path(target).expanduser().resolve()))
    if args.all_project_mirrors:
        existing = {(target.label, target.path) for target in targets}
        for candidate in [
            SyncTarget("codex-plugin-mirror", project_root / "skills"),
            SyncTarget("claude-project", project_root / ".claude" / "skills"),
            SyncTarget("opencode-project", project_root / ".opencode" / "skills"),
        ]:
            if (candidate.label, candidate.path) not in existing:
                targets.append(candidate)
                existing.add((candidate.label, candidate.path))
    return targets


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Project root containing .agents/skills")
    parser.add_argument("--target", action="append", help="Custom target skills directory; may be passed multiple times")
    parser.add_argument("--plugin-mirror", action="store_true", help="Copy/check <project-root>/skills for Codex plugin packaging")
    parser.add_argument("--codex-user", action="store_true", help="Copy/check ~/.codex/skills")
    parser.add_argument("--claude-project", action="store_true", help="Copy/check <project-root>/.claude/skills")
    parser.add_argument("--opencode-project", action="store_true", help="Copy/check <project-root>/.opencode/skills")
    parser.add_argument("--all-project-mirrors", action="store_true", help="Copy/check plugin, Claude, and OpenCode project mirrors")
    parser.add_argument("--replace", action="store_true", help="Replace existing skill directories during copy")
    parser.add_argument("--check-drift", action="store_true", help="Compare targets with source and fail when they differ")
    parser.add_argument("--list-targets", action="store_true", help="Print resolved source and targets without copying")
    parser.add_argument("--manifest-json", type=Path, help="Write source/target manifest JSON")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    source = project_root / ".agents" / "skills"
    targets = build_targets(args, project_root)

    if not source.exists():
        raise SystemExit(f"Source skills directory not found: {source}")
    if not targets:
        raise SystemExit("Specify --target, --plugin-mirror, --codex-user, --claude-project, --opencode-project, or --all-project-mirrors")

    manifest = target_manifest(source, targets)
    if args.manifest_json:
        args.manifest_json.parent.mkdir(parents=True, exist_ok=True)
        args.manifest_json.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Wrote manifest: {args.manifest_json}")

    if args.list_targets:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return 0

    if args.check_drift:
        failed = False
        for target in targets:
            if target.label == "codex-user":
                diff = compare_user_skill_install(source, target.path)
            else:
                diff = compare_trees(source, target.path)
            has_diff = any(diff.values())
            status = "DRIFT" if has_diff else "OK"
            print(f"{status}: {target.label} -> {target.path}")
            if has_diff:
                failed = True
                for key in ("missing", "extra", "changed"):
                    values = diff[key]
                    if values:
                        preview = ", ".join(values[:10])
                        more = f" (+{len(values) - 10} more)" if len(values) > 10 else ""
                        print(f"  {key}: {preview}{more}")
        return 1 if failed else 0

    for target in targets:
        copied = copy_skills(source, target.path, args.replace)
        print(f"Copied {len(copied)} skills to {target.label}: {target.path}")
        for skill_name in copied:
            print(f"- {skill_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
