#!/usr/bin/env python3
"""Deploy this package as a local Codex plugin."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


DEFAULT_MARKETPLACE = Path.home() / ".agents" / "plugins" / "marketplace.json"
DEFAULT_PLUGIN_PARENT = Path.home() / "plugins"


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError(f"{path} must contain a JSON object")
    return payload


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def plugin_name(project_root: Path) -> str:
    manifest = load_json(project_root / ".codex-plugin" / "plugin.json")
    name = manifest.get("name")
    if not isinstance(name, str) or not name.strip():
        raise ValueError("plugin.json field name must be a non-empty string")
    return name


def update_cachebuster(project_root: Path, token: str | None = None) -> str:
    manifest_path = project_root / ".codex-plugin" / "plugin.json"
    manifest = load_json(manifest_path)
    version = manifest.get("version")
    if not isinstance(version, str) or not version.strip():
        raise ValueError("plugin.json field version must be a non-empty string")
    raw_token = token or datetime.now(timezone.utc).strftime("local-%Y%m%d%H%M%S")
    cachebuster = re.sub(r"[^a-z0-9-]+", "-", raw_token.strip().lower())
    cachebuster = re.sub(r"-{2,}", "-", cachebuster).strip("-")
    if not cachebuster:
        raise ValueError("Cachebuster must contain at least one letter or digit")
    base_version = version.split("+", 1)[0]
    next_version = f"{base_version}+codex.{cachebuster}"
    manifest["version"] = next_version
    write_json(manifest_path, manifest)
    return next_version


def sync_plugin_skills(project_root: Path) -> None:
    subprocess.run(
        [
            sys.executable,
            str(project_root / "tools" / "sync-skills.py"),
            "--project-root",
            str(project_root),
            "--plugin-mirror",
            "--replace",
        ],
        check=True,
    )


def should_skip(path: Path) -> bool:
    skip_names = {".git", "__pycache__", ".pytest_cache", ".mypy_cache"}
    return any(part in skip_names for part in path.parts)


def copy_plugin(project_root: Path, destination: Path, replace: bool) -> None:
    plugin_parent = destination.parent.resolve()
    destination = destination.resolve()
    if not destination.is_relative_to(plugin_parent):
        raise ValueError(f"Refusing to deploy outside plugin parent: {destination}")
    if destination.exists():
        if not replace:
            raise FileExistsError(f"{destination} already exists. Use --replace to overwrite.")
        shutil.rmtree(destination)
    destination.mkdir(parents=True, exist_ok=True)

    for item in project_root.iterdir():
        if should_skip(item):
            continue
        target = destination / item.name
        if item.is_dir():
            ignore = shutil.ignore_patterns("__pycache__", ".pytest_cache", ".mypy_cache")
            shutil.copytree(item, target, ignore=ignore)
        else:
            shutil.copy2(item, target)


def update_marketplace(marketplace_path: Path, name: str, category: str) -> str:
    if marketplace_path.exists():
        marketplace = load_json(marketplace_path)
    else:
        marketplace = {
            "name": "personal",
            "interface": {"displayName": "Personal"},
            "plugins": [],
        }

    marketplace_name = marketplace.get("name")
    if not isinstance(marketplace_name, str) or not marketplace_name.strip():
        raise ValueError(f"{marketplace_path} must contain a non-empty name")

    plugins = marketplace.setdefault("plugins", [])
    if not isinstance(plugins, list):
        raise ValueError(f"{marketplace_path} field plugins must be an array")

    entry = {
        "name": name,
        "source": {
            "source": "local",
            "path": f"./plugins/{name}",
        },
        "policy": {
            "installation": "AVAILABLE",
            "authentication": "ON_INSTALL",
        },
        "category": category,
    }

    for index, existing in enumerate(plugins):
        if isinstance(existing, dict) and existing.get("name") == name:
            plugins[index] = entry
            break
    else:
        plugins.append(entry)

    write_json(marketplace_path, marketplace)
    return marketplace_name


def install_with_codex(name: str, marketplace_name: str) -> None:
    plugin_ref = f"{name}@{marketplace_name}"
    codex_path = shutil.which("codex")
    if os.name == "nt" and codex_path and codex_path.lower().endswith((".cmd", ".bat")):
        subprocess.run(["cmd.exe", "/c", codex_path, "plugin", "add", plugin_ref], check=True)
        return
    if os.name == "nt" and codex_path and codex_path.lower().endswith(".ps1"):
        subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                codex_path,
                "plugin",
                "add",
                plugin_ref,
            ],
            check=True,
        )
        return
    subprocess.run(["codex", "plugin", "add", plugin_ref], check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--plugin-parent", default=str(DEFAULT_PLUGIN_PARENT), help="Local plugin parent directory")
    parser.add_argument("--marketplace", default=str(DEFAULT_MARKETPLACE), help="Personal marketplace.json path")
    parser.add_argument("--category", default="Productivity", help="Marketplace category")
    parser.add_argument("--replace", action="store_true", help="Replace existing local plugin copy")
    parser.add_argument("--install", action="store_true", help="Run codex plugin add after deployment")
    parser.add_argument(
        "--update-cachebuster",
        action="store_true",
        help="Rewrite plugin.json version to <base>+codex.<token> before deployment",
    )
    parser.add_argument("--cachebuster", help="Optional cachebuster token for --update-cachebuster")
    parser.add_argument("--dry-run", action="store_true", help="Print planned paths without writing")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    name = plugin_name(project_root)
    plugin_parent = Path(args.plugin_parent).expanduser().resolve()
    destination = plugin_parent / name
    marketplace_path = Path(args.marketplace).expanduser().resolve()

    if args.dry_run:
        print(json.dumps({
            "projectRoot": str(project_root),
            "pluginName": name,
            "destination": str(destination),
            "marketplace": str(marketplace_path),
            "install": bool(args.install),
        }, indent=2))
        return 0

    next_version = None
    if args.update_cachebuster:
        next_version = update_cachebuster(project_root, args.cachebuster)
    sync_plugin_skills(project_root)
    copy_plugin(project_root, destination, args.replace)
    marketplace_name = update_marketplace(marketplace_path, name, args.category)
    if args.install:
        install_with_codex(name, marketplace_name)

    print(json.dumps({
        "pluginName": name,
        "destination": str(destination),
        "marketplace": str(marketplace_path),
        "marketplaceName": marketplace_name,
        "version": next_version,
        "installed": bool(args.install),
    }, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
