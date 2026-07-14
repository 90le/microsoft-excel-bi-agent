#!/usr/bin/env python3
"""Build a compact, deterministic Codex runtime package from the source repository."""

from __future__ import annotations

import argparse
import ast
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any, Iterable


MANIFEST_NAME = "runtime-package-manifest.json"
REFERENCE_RE = re.compile(
    r"(?P<path>(?:\.agents/skills|skills|tools|fixtures|schemas)/"
    r"[A-Za-z0-9_.\-/]+\.(?:py|ps1|sh|json|md|yaml|yml|m|bas|cls|frm|xml))",
    re.IGNORECASE,
)
FORBIDDEN_SUFFIXES = {".xls", ".xlsx", ".xlsm", ".xlsb", ".pdf", ".pyc", ".pyo"}
FORBIDDEN_PARTS = {".git", ".agents", ".claude", ".opencode", "__pycache__", ".pytest_cache", ".mypy_cache"}
FORBIDDEN_REPORT_RE = re.compile(r"(^|/)(excel_bi_)?release_gate.*\.(json|md)$", re.IGNORECASE)
SAFE_FIXTURE_PREFIXES = ("fixtures/real-sanitized-cases/",)
SAFE_FIXTURE_SUFFIXES = {".json", ".md", ".m", ".txt", ".xml", ".yaml", ".yml"}
SAFE_REFERENCED_FIXTURE_SUFFIXES = SAFE_FIXTURE_SUFFIXES | {".py", ".ps1", ".sh", ".bas", ".cls", ".frm"}
PRIVATE_FIXTURE_SUFFIXES = {
    ".csv",
    ".tsv",
    ".parquet",
    ".feather",
    ".db",
    ".sqlite",
    ".sqlite3",
    ".sql",
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".pdf",
}
RUNTIME_README = """# Microsoft Excel BI Agent Runtime Package

This compact package is generated from the Microsoft Excel BI Agent source repository for local Codex plugin execution.

- Start with the skills under `skills/`.
- Runtime helper scripts referenced by those skills are under `tools/`.
- Sanitized regression inputs are under `fixtures/`.
- `runtime-package-manifest.json` records the sorted payload, byte sizes, and SHA-256 hashes.

This runtime package intentionally excludes canonical authoring sources, cross-agent mirrors, development documentation, release tooling, and private artifacts. Structural checks do not prove Excel COM, VBA, Power Query, Power Pivot, or workbook business behavior.

Source and documentation: https://github.com/90le/microsoft-excel-bi-agent
"""


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def relative_files(root: Path) -> list[Path]:
    if not root.is_dir():
        return []
    return sorted(
        path.relative_to(root)
        for path in root.rglob("*")
        if path.is_file() and not any(part in FORBIDDEN_PARTS for part in path.relative_to(root).parts)
    )


def compare_skill_mirror(project_root: Path) -> dict[str, Any]:
    canonical = project_root / ".agents" / "skills"
    mirror = project_root / "skills"
    if not canonical.is_dir() or not mirror.is_dir():
        return {"status": "not-checked", "missing": [], "extra": [], "changed": []}

    canonical_files = set(relative_files(canonical))
    mirror_files = set(relative_files(mirror))
    missing = sorted(path.as_posix() for path in canonical_files - mirror_files)
    extra = sorted(path.as_posix() for path in mirror_files - canonical_files)
    changed = sorted(
        path.as_posix()
        for path in canonical_files & mirror_files
        if sha256_file(canonical / path) != sha256_file(mirror / path)
    )
    return {
        "status": "drift" if missing or extra or changed else "in-sync",
        "missing": missing,
        "extra": extra,
        "changed": changed,
    }


def normalize_reference(reference: str) -> str:
    normalized = reference.replace("\\", "/")
    if normalized.startswith(".agents/skills/"):
        return "skills/" + normalized[len(".agents/skills/") :]
    return normalized


def text_references(path: Path) -> set[str]:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return set()
    return {normalize_reference(match.group("path")) for match in REFERENCE_RE.finditer(text)}


def skill_references(project_root: Path) -> tuple[set[str], list[str]]:
    skills_root = project_root / "skills"
    references: set[str] = set()
    unresolved: set[str] = set()
    for relative in relative_files(skills_root):
        source = skills_root / relative
        for reference in text_references(source):
            references.add(reference)
            if not (project_root / reference).is_file():
                unresolved.add(reference)
    return references, sorted(unresolved)


def local_tool_modules(project_root: Path) -> dict[str, str]:
    modules: dict[str, str] = {}
    tools_root = project_root / "tools"
    for path in sorted(tools_root.rglob("*.py")) if tools_root.is_dir() else []:
        relative = path.relative_to(tools_root)
        module_parts = list(relative.with_suffix("").parts)
        if module_parts[-1] == "__init__":
            module_parts.pop()
        if not module_parts:
            continue
        module = ".".join(module_parts)
        packaged_path = f"tools/{relative.as_posix()}"
        modules[module] = packaged_path
        modules[f"tools.{module}"] = packaged_path
    return modules


def imported_module_names(project_root: Path, path: Path) -> tuple[set[str], str | None]:
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=path.name)
    except (OSError, UnicodeDecodeError, SyntaxError) as exc:
        return set(), f"cannot parse Python imports for {path.name}: {type(exc).__name__}"

    names: set[str] = set()
    relative = path.relative_to(project_root / "tools").with_suffix("")
    current_parts = list(relative.parts)
    current_package = current_parts if current_parts[-1] == "__init__" else current_parts[:-1]
    if current_package and current_package[-1] == "__init__":
        current_package.pop()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.level:
                keep = max(0, len(current_package) - (node.level - 1))
                base_parts = current_package[:keep]
                if node.module:
                    base_parts.extend(node.module.split("."))
                module = ".".join(base_parts)
            else:
                module = node.module or ""
            if module:
                names.add(module)
                names.update(f"{module}.{alias.name}" for alias in node.names if alias.name != "*")
            elif node.level:
                names.update(alias.name for alias in node.names if alias.name != "*")
    return names, None


def expand_tool_references(project_root: Path, initial: Iterable[str]) -> tuple[set[str], list[str]]:
    selected = {path for path in initial if path.startswith("tools/") and (project_root / path).is_file()}
    pending = list(selected)
    module_paths = local_tool_modules(project_root)
    errors: set[str] = set()
    while pending:
        reference = pending.pop()
        for nested in text_references(project_root / reference):
            if nested.startswith("tools/") and nested not in selected and (project_root / nested).is_file():
                selected.add(nested)
                pending.append(nested)
        if not reference.endswith(".py"):
            continue
        imported, parse_error = imported_module_names(project_root, project_root / reference)
        if parse_error:
            errors.add(f"{reference}: {parse_error}")
            continue
        for module in sorted(imported):
            dependency = module_paths.get(module)
            if dependency and dependency not in selected:
                selected.add(dependency)
                pending.append(dependency)
    return selected, sorted(errors)


def fixture_forbidden_reason(relative: str) -> str | None:
    path = Path(relative)
    fixture_parts = path.parts[1:]
    if any(part.startswith(".") for part in fixture_parts):
        return "fixture dotfile is not distributable"
    if path.suffix.lower() in PRIVATE_FIXTURE_SUFFIXES:
        return f"private/customer-data fixture suffix {path.suffix.lower()}"
    return forbidden_reason(relative)


def select_fixture_files(
    project_root: Path,
    references: set[str],
) -> tuple[set[str], list[str], list[str]]:
    selected: set[str] = set()
    forbidden: list[str] = []
    warnings: list[str] = []
    fixtures_root = project_root / "fixtures"
    for path in relative_files(fixtures_root):
        relative = f"fixtures/{path.as_posix()}"
        reason = fixture_forbidden_reason(relative)
        if reason:
            forbidden.append(f"{relative}: {reason}")
            continue
        explicitly_referenced = relative in references
        structured_sanitized = relative.startswith(SAFE_FIXTURE_PREFIXES)
        safe_suffix = (
            structured_sanitized and path.suffix.lower() in SAFE_FIXTURE_SUFFIXES
        ) or (
            explicitly_referenced and path.suffix.lower() in SAFE_REFERENCED_FIXTURE_SUFFIXES
        )
        if safe_suffix:
            selected.add(relative)
        else:
            warnings.append(f"Excluded unallowlisted fixture: {relative}")
    return selected, forbidden, warnings


def forbidden_reason(relative: str) -> str | None:
    path = Path(relative)
    if any(part in FORBIDDEN_PARTS for part in path.parts):
        return "forbidden directory"
    if path.name.startswith("~$"):
        return "Excel lock file"
    if path.suffix.lower() in FORBIDDEN_SUFFIXES:
        return f"forbidden artifact suffix {path.suffix.lower()}"
    if FORBIDDEN_REPORT_RE.search(relative.replace("\\", "/")):
        return "generated release report"
    return None


def copy_payload_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    try:
        destination.chmod(0o755 if destination.suffix.lower() == ".sh" else 0o644)
    except OSError:
        pass


def prepare_output_dir(project_root: Path, out_dir: Path) -> None:
    project_root = project_root.resolve()
    out_dir = out_dir.resolve()
    if out_dir == project_root or out_dir in project_root.parents or project_root in out_dir.parents:
        raise ValueError(f"Runtime staging directory must be outside the project tree: {out_dir}")
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)


def source_size(project_root: Path, excluded_root: Path | None = None) -> int:
    total = 0
    excluded = excluded_root.resolve() if excluded_root else None
    for path in project_root.rglob("*"):
        if not path.is_file():
            continue
        resolved = path.resolve()
        if excluded is not None and (resolved == excluded or excluded in resolved.parents):
            continue
        relative = path.relative_to(project_root)
        if any(part in {".git", "__pycache__", ".pytest_cache", ".mypy_cache"} for part in relative.parts):
            continue
        total += path.stat().st_size
    return total


def payload_inventory(out_dir: Path) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    paths = [item for item in out_dir.rglob("*") if item.is_file() and item.name != MANIFEST_NAME]
    for path in sorted(paths, key=lambda item: item.relative_to(out_dir).as_posix()):
        entries.append(
            {
                "path": path.relative_to(out_dir).as_posix(),
                "size": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )
    return entries


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def validate_zip_path(project_root: Path, out_dir: Path, zip_path: Path) -> None:
    project_root = project_root.resolve()
    out_dir = out_dir.resolve()
    zip_path = zip_path.resolve()
    if zip_path == project_root or zip_path in project_root.parents or project_root in zip_path.parents:
        raise ValueError(f"Zip output must be outside the project tree: {zip_path}")
    if zip_path == out_dir or out_dir in zip_path.parents:
        raise ValueError("Zip output must be outside the runtime staging directory")


def write_deterministic_zip(out_dir: Path, zip_path: Path) -> None:
    out_dir = out_dir.resolve()
    zip_path = zip_path.resolve()
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(item for item in out_dir.rglob("*") if item.is_file()):
            relative = path.relative_to(out_dir).as_posix()
            info = zipfile.ZipInfo(relative, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            mode = 0o755 if path.suffix.lower() == ".sh" else 0o644
            info.external_attr = (stat.S_IFREG | mode) << 16
            archive.writestr(info, path.read_bytes(), compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)


def smoke_python_tools(out_dir: Path) -> tuple[list[dict[str, Any]], list[str]]:
    results: list[dict[str, Any]] = []
    errors: list[str] = []
    tools_root = out_dir / "tools"
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    environment["PYTHONUTF8"] = "1"
    environment["PYTHONPATH"] = os.pathsep.join(
        item for item in [str(out_dir), str(tools_root), environment.get("PYTHONPATH", "")] if item
    )
    paths = sorted(tools_root.rglob("*.py")) if tools_root.is_dir() else []
    for path in paths:
        relative = path.relative_to(out_dir).as_posix()
        try:
            completed = subprocess.run(
                [sys.executable, str(path), "--help"],
                cwd=str(tools_root),
                env=environment,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=30,
                check=False,
            )
            return_code: int | None = completed.returncode
            status = "pass" if completed.returncode == 0 else "fail"
        except subprocess.TimeoutExpired:
            return_code = None
            status = "timeout"
        results.append({"path": relative, "status": status, "returnCode": return_code})
        if status != "pass":
            detail = "timeout" if return_code is None else f"exit {return_code}"
            errors.append(f"python tool --help smoke failed: {relative} ({detail})")
    return results, errors


def build_runtime_package(
    project_root: Path,
    out_dir: Path,
    *,
    zip_path: Path | None = None,
) -> dict[str, Any]:
    project_root = Path(project_root).expanduser().resolve()
    out_dir = Path(out_dir).expanduser().resolve()
    resolved_zip = Path(zip_path).expanduser().resolve() if zip_path is not None else None
    if resolved_zip is not None:
        validate_zip_path(project_root, out_dir, resolved_zip)
    prepare_output_dir(project_root, out_dir)
    manifest_path = project_root / ".codex-plugin" / "plugin.json"
    license_path = project_root / "LICENSE"
    skills_root = project_root / "skills"

    errors: list[str] = []
    warnings: list[str] = []
    for required in [manifest_path, license_path]:
        if not required.is_file():
            errors.append(f"missing required runtime file: {required.relative_to(project_root).as_posix()}")
    if not skills_root.is_dir() or not list(skills_root.glob("*/SKILL.md")):
        errors.append("missing runtime skills: skills/*/SKILL.md")

    references, unresolved = skill_references(project_root)
    if unresolved:
        errors.append(f"unresolved skill references: {', '.join(unresolved)}")
    runtime_tools, dependency_errors = expand_tool_references(project_root, references)
    errors.extend(dependency_errors)

    candidates: set[str] = set()
    for root_name in [".codex-plugin", "skills", "schemas"]:
        root = project_root / root_name
        candidates.update(f"{root_name}/{path.as_posix()}" for path in relative_files(root))
    fixture_files, forbidden_files, fixture_warnings = select_fixture_files(project_root, references)
    candidates.update(fixture_files)
    warnings.extend(fixture_warnings)
    candidates.update(runtime_tools)
    if license_path.is_file():
        candidates.add("LICENSE")

    for relative in sorted(candidates):
        reason = forbidden_reason(relative)
        if reason:
            forbidden_files.append(f"{relative}: {reason}")
            continue
        source = project_root / relative
        if source.is_file():
            copy_payload_file(source, out_dir / relative)

    (out_dir / "README.md").write_text(RUNTIME_README, encoding="utf-8", newline="\n")
    try:
        (out_dir / "README.md").chmod(0o644)
    except OSError:
        pass

    if forbidden_files:
        errors.append(f"forbidden runtime files: {', '.join(forbidden_files)}")

    python_tool_smoke, smoke_errors = smoke_python_tools(out_dir)
    errors.extend(smoke_errors)

    mirror = compare_skill_mirror(project_root)
    if mirror["status"] == "drift":
        warnings.append("Canonical skill mirror drift detected; package used the existing skills/ mirror without syncing it.")

    files = payload_inventory(out_dir)
    total_bytes = sum(int(item["size"]) for item in files)
    source_bytes = source_size(project_root, out_dir)
    manifest: dict[str, Any] = {
        "schemaVersion": "1.0",
        "kind": "codex-runtime-package",
        "status": "pass" if not errors else "fail",
        "files": files,
        "fileCount": len(files),
        "totalBytes": total_bytes,
        "sourceBytes": source_bytes,
        "reductionPercent": round((1 - total_bytes / source_bytes) * 100, 2) if source_bytes else 0.0,
        "allowedRoots": [".codex-plugin/", "skills/", "tools/", "fixtures/", "schemas/", "LICENSE", "README.md"],
        "skillMirror": mirror,
        "pythonToolSmoke": python_tool_smoke,
        "unresolvedReferences": unresolved,
        "forbiddenFiles": forbidden_files,
        "warnings": warnings,
        "errors": errors,
    }
    write_json(out_dir / MANIFEST_NAME, manifest)
    if resolved_zip is not None:
        write_deterministic_zip(out_dir, resolved_zip)
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Source repository root")
    parser.add_argument("--out-dir", required=True, type=Path, help="Runtime staging directory")
    parser.add_argument("--zip", dest="zip_path", type=Path, help="Optional deterministic zip output")
    parser.add_argument("--require-pass", action="store_true", help="Exit non-zero for unresolved references or forbidden files")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = build_runtime_package(
        Path(args.project_root),
        args.out_dir,
        zip_path=args.zip_path,
    )
    print(
        json.dumps(
            {
                "status": manifest["status"],
                "fileCount": manifest["fileCount"],
                "totalBytes": manifest["totalBytes"],
                "sourceBytes": manifest["sourceBytes"],
                "reductionPercent": manifest["reductionPercent"],
                "skillMirrorStatus": manifest["skillMirror"]["status"],
            },
            ensure_ascii=False,
        )
    )
    return 1 if args.require_pass and manifest["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
