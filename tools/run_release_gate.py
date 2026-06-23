#!/usr/bin/env python3
"""Run the Microsoft Excel BI Agent Pack release gate.

This script is intentionally conservative: it gathers repeatable validation
evidence, emits a machine-readable report, and avoids claiming Excel runtime
behavior that was not actually tested.
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path


PASS = "pass"
FAIL = "fail"
SKIP = "skip"
WARN = "warn"


PLACEHOLDER_RE = re.compile(
    "|".join([r"\b" + "TO" + "DO" + r"\b", r"\[" + "TO" + "DO" + r"\]", r"FIX" + "ME"]),
    re.IGNORECASE,
)

DEFAULT_SENSITIVE_MARKERS = [
    "budget" + "_" + "optimizer",
    "\u9884\u7b97",
    "\u79d2\u9488",
    "\u7eafVBA",
    "\u52a8\u6001\u63a8\u8350",
    "\u7acb\u767d",
]


CODEX_CACHEBUSTER_RE = re.compile(r"^\d+\.\d+\.\d+\+codex\.[A-Za-z0-9][A-Za-z0-9._-]*$")


@dataclass
class CheckResult:
    name: str
    status: str
    detail: str = ""
    command: list[str] | None = None
    stdout: str = ""
    stderr: str = ""
    metadata: dict[str, object] = field(default_factory=dict)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def run_command(command: list[str], cwd: Path, name: str, ok_codes: set[int] | None = None) -> CheckResult:
    ok_codes = ok_codes or {0}
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    try:
        completed = subprocess.run(
            command,
            cwd=str(cwd),
            text=True,
            capture_output=True,
            env=env,
            timeout=180,
        )
        status = PASS if completed.returncode in ok_codes else FAIL
        detail = f"exit_code={completed.returncode}"
        return CheckResult(
            name=name,
            status=status,
            detail=detail,
            command=command,
            stdout=completed.stdout.strip(),
            stderr=completed.stderr.strip(),
        )
    except FileNotFoundError as exc:
        return CheckResult(name=name, status=SKIP, detail=str(exc), command=command)
    except PermissionError as exc:
        return CheckResult(name=name, status=SKIP, detail=f"permission denied: {exc}", command=command)
    except OSError as exc:
        return CheckResult(name=name, status=FAIL, detail=str(exc), command=command)
    except subprocess.TimeoutExpired as exc:
        return CheckResult(
            name=name,
            status=FAIL,
            detail=f"timed out after {exc.timeout} seconds",
            command=command,
            stdout=(exc.stdout or "").strip() if isinstance(exc.stdout, str) else "",
            stderr=(exc.stderr or "").strip() if isinstance(exc.stderr, str) else "",
        )


def default_plugin_validator() -> Path:
    return Path.home() / ".codex" / "skills" / ".system" / "plugin-creator" / "scripts" / "validate_plugin.py"


def read_plugin_json(project_root: Path) -> dict[str, object]:
    manifest = project_root / ".codex-plugin" / "plugin.json"
    if not manifest.is_file():
        return {}
    return json.loads(manifest.read_text(encoding="utf-8"))


def find_bash() -> str | None:
    for candidate in [
        shutil.which("bash"),
        r"C:\Program Files\Git\bin\bash.exe",
        r"C:\Program Files\Git\usr\bin\bash.exe",
    ]:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def find_powershell() -> str | None:
    for candidate in [shutil.which("powershell"), shutil.which("powershell.exe"), shutil.which("pwsh"), shutil.which("pwsh.exe")]:
        if candidate:
            return candidate
    return None


def find_codex() -> str | None:
    for name in ["codex.cmd", "codex.exe", "codex"]:
        candidate = shutil.which(name)
        if candidate:
            return candidate
    return None


def remove_pycache(root: Path) -> None:
    for path in root.rglob("__pycache__"):
        if path.is_dir():
            shutil.rmtree(path)


def relative_file_list(root: Path) -> list[Path]:
    ignored_dirs = {".git", "__pycache__"}
    files: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if any(part in ignored_dirs for part in path.relative_to(root).parts):
            continue
        files.append(path)
    return files


def scan_regex(root: Path, pattern: re.Pattern[str], name: str, encoding: str = "utf-8") -> CheckResult:
    hits: list[dict[str, object]] = []
    for path in relative_file_list(root):
        try:
            text = path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
        except OSError:
            continue
        for line_no, line in enumerate(text.splitlines(), start=1):
            if pattern.search(line):
                hits.append({"file": str(path.relative_to(root)), "line": line_no, "text": line[:200]})
                if len(hits) >= 50:
                    break
        if len(hits) >= 50:
            break
    status = PASS if not hits else FAIL
    detail = "no matches" if not hits else f"{len(hits)} matches"
    return CheckResult(name=name, status=status, detail=detail, metadata={"hits": hits})


def scan_sensitive_markers(root: Path, markers: list[str]) -> CheckResult:
    if not markers:
        return CheckResult(name="sensitive marker scan", status=SKIP, detail="no markers supplied")
    escaped = [re.escape(marker) for marker in markers if marker]
    if not escaped:
        return CheckResult(name="sensitive marker scan", status=SKIP, detail="no markers supplied")
    return scan_regex(root, re.compile("|".join(escaped), re.IGNORECASE), "sensitive marker scan")


def source_script_files(project_root: Path, pattern: str) -> list[Path]:
    files = list((project_root / "tools").glob(pattern))
    skill_root = project_root / ".agents" / "skills"
    if skill_root.is_dir():
        files.extend(skill_root.glob(f"*/scripts/{pattern}"))
    return sorted({path.resolve() for path in files if path.is_file()})


def powershell_parse_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    files = source_script_files(project_root, "*.ps1")
    if not files:
        return CheckResult(name="PowerShell syntax", status=SKIP, detail="no .ps1 files")
    if not ps_exe:
        return CheckResult(name="PowerShell syntax", status=SKIP, detail="PowerShell not found")

    script = (
        "$failed = @(); "
        "foreach ($p in $args) { "
        "$tokens=$null; $errs=$null; "
        "$null=[System.Management.Automation.Language.Parser]::ParseFile($p,[ref]$tokens,[ref]$errs); "
        "if ($errs.Count -gt 0) { $failed += ($p + ': ' + (($errs | ForEach-Object { $_.Message }) -join ' | ')) } "
        "} "
        "if ($failed.Count -gt 0) { $failed | ForEach-Object { Write-Error $_ }; exit 1 } "
        "Write-Output ('PowerShell syntax OK: ' + $args.Count + ' files')"
    )
    result = run_command([ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script, *[str(p) for p in files]], project_root, "PowerShell syntax")
    if result.status == PASS:
        result.detail = f"{len(files)} files parsed"
    return result


def bash_syntax_check(project_root: Path, bash: str | None) -> CheckResult:
    files = source_script_files(project_root, "*.sh")
    if not files:
        return CheckResult(name="Bash syntax", status=SKIP, detail="no .sh files")
    if not bash:
        return CheckResult(name="Bash syntax", status=SKIP, detail="bash not found")
    failures: list[str] = []
    outputs: list[str] = []
    for path in files:
        result = run_command([bash, "-n", str(path)], project_root, f"Bash syntax: {path.name}")
        outputs.append(f"{path.name}: {result.status}")
        if result.status != PASS:
            failures.append(f"{path.name}: {result.stderr or result.stdout or result.detail}")
    return CheckResult(
        name="Bash syntax",
        status=PASS if not failures else FAIL,
        detail=f"{len(files)} files checked",
        stdout="\n".join(outputs),
        stderr="\n".join(failures),
    )


def portable_structural_wrapper_fixture_check(project_root: Path, bash: str | None) -> CheckResult:
    name = "Portable structural wrapper fixture smoke"
    if not bash:
        return CheckResult(name=name, status=SKIP, detail="bash not found")
    wrapper = project_root / "tools" / "excel_bi_structural.sh"
    if not wrapper.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {wrapper}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_structural_wrapper_") as tmp:
        tmp_dir = Path(tmp)
        bundle_dir = tmp_dir / "sanitized_bundle"
        provider_dir = tmp_dir / "provider_baseline"
        result = run_command(
            [
                bash,
                str(wrapper),
                "sanitized-bundle",
                "--out-dir",
                str(bundle_dir),
                "--clean",
                "--validate",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            result.name = name
            return result

        provider_result = run_command(
            [
                bash,
                str(wrapper),
                "provider-baseline-fixture",
                "--out-dir",
                str(provider_dir),
                "--clean",
            ],
            project_root,
            f"{name}: provider baseline",
        )
        if provider_result.status != PASS:
            return CheckResult(
                name=name,
                status=FAIL,
                detail=f"provider baseline wrapper failed: {provider_result.detail}",
                command=provider_result.command,
                stdout="\n".join(part for part in [result.stdout, provider_result.stdout] if part),
                stderr="\n".join(part for part in [result.stderr, provider_result.stderr] if part),
            )

        summary_path = bundle_dir / "_validation" / "structural_wrapper_summary.json"
        provider_summary_path = provider_dir / "reports" / "provider_baseline_wrapper_summary.json"
        try:
            summary = json.loads(summary_path.read_text(encoding="utf-8"))
            provider_summary = json.loads(provider_summary_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(
                name=name,
                status=FAIL,
                detail=f"cannot read wrapper summary: {exc}",
                command=result.command,
                stdout="\n".join(part for part in [result.stdout, provider_result.stdout] if part),
                stderr="\n".join(part for part in [result.stderr, provider_result.stderr] if part),
            )

        failures: list[str] = []
        expected_ids = {"cube-formula", "external-dependency", "pure-deliverable", "power-query-lineage"}
        actual_ids = set(summary.get("fixtureIds", []))
        if summary.get("fixtureCount") != 4:
            failures.append(f"fixtureCount={summary.get('fixtureCount')}")
        if actual_ids != expected_ids:
            failures.append(f"fixtureIds={sorted(actual_ids)}")
        if summary.get("cubeFormulaCount") != 7:
            failures.append(f"cubeFormulaCount={summary.get('cubeFormulaCount')}")
        if summary.get("externalReadiness") != "blocked-for-pure-deliverable":
            failures.append(f"externalReadiness={summary.get('externalReadiness')}")
        if summary.get("externalCredentialLikeConnectionCount") != 1:
            failures.append(f"externalCredentialLikeConnectionCount={summary.get('externalCredentialLikeConnectionCount')}")
        if summary.get("pureReadiness") != "clean":
            failures.append(f"pureReadiness={summary.get('pureReadiness')}")
        if summary.get("pqSafeQueryCount") != 4 or summary.get("pqSafeFindingCount") != 0:
            failures.append(f"pq safe summary={summary.get('pqSafeQueryCount')}/{summary.get('pqSafeFindingCount')}")
        if summary.get("pqRiskyQueryCount") != 11 or summary.get("pqRiskyHighFindingCount") != 3:
            failures.append(f"pq risky summary={summary.get('pqRiskyQueryCount')}/{summary.get('pqRiskyHighFindingCount')}")
        required_codes = {
            "hard-coded-local-path",
            "web-source",
            "database-source",
            "cloud-service-source",
            "native-query-review",
            "credential-like-literal",
            "mixed-source-lineage",
            "query-dependency-cycle",
        }
        missing_codes = required_codes - set(summary.get("pqRiskyCodes", []))
        if missing_codes:
            failures.append(f"pq risky codes missing={sorted(missing_codes)}")
        if provider_summary.get("matchingStatus") != PASS:
            failures.append(f"provider matchingStatus={provider_summary.get('matchingStatus')}")
        if provider_summary.get("matchingChangedCount") != 0:
            failures.append(f"provider matchingChangedCount={provider_summary.get('matchingChangedCount')}")
        minimum_drift = int(provider_summary.get("expectedMinimumDriftCount") or 1)
        if provider_summary.get("driftStatus") != FAIL:
            failures.append(f"provider driftStatus={provider_summary.get('driftStatus')}")
        if int(provider_summary.get("driftChangedCount") or 0) < minimum_drift:
            failures.append(f"provider driftChangedCount={provider_summary.get('driftChangedCount')}")
        missing_provider_paths = set(provider_summary.get("requiredDriftPaths", [])) - set(provider_summary.get("driftPaths", []))
        if missing_provider_paths:
            failures.append(f"provider drift paths missing={sorted(missing_provider_paths)}")

        metadata = dict(summary)
        metadata.update(
            {
                "providerMatchingChangedCount": provider_summary.get("matchingChangedCount"),
                "providerDriftChangedCount": provider_summary.get("driftChangedCount"),
                "providerDriftPaths": provider_summary.get("driftPaths"),
            }
        )

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="portable structural wrapper validated sanitized workbook/query-source bundle and provider baseline drift fixture" if not failures else "; ".join(failures),
            command=result.command,
            stdout="\n".join(part for part in [result.stdout, provider_result.stdout] if part),
            stderr="\n".join(part for part in [result.stderr, provider_result.stderr] if part),
            metadata=metadata,
        )


def python_compile_check(project_root: Path) -> CheckResult:
    files = source_script_files(project_root, "*.py")
    if not files:
        return CheckResult(name="Python compile", status=SKIP, detail="no .py files")
    result = run_command([sys.executable, "-m", "py_compile", *[str(p) for p in files]], project_root, "Python compile")
    remove_pycache(project_root)
    if result.status == PASS:
        result.detail = f"{len(files)} files compiled"
    return result


def pycache_check(root: Path) -> CheckResult:
    hits = [str(path.relative_to(root)) for path in root.rglob("__pycache__") if path.is_dir()]
    return CheckResult(
        name="Python cache cleanup",
        status=PASS if not hits else FAIL,
        detail="no __pycache__ directories" if not hits else f"{len(hits)} directories found",
        metadata={"hits": hits},
    )


def sanitized_fixture_bundle_check(project_root: Path) -> CheckResult:
    name = "Sanitized fixture bundle smoke"
    bundle_script = project_root / "tools" / "build_sanitized_fixture_bundle.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    cube_report_script = project_root / "tools" / "build_cube_dependency_report.py"
    readiness_script = project_root / "tools" / "build_external_dependency_report.py"
    pq_lineage_script = project_root / "tools" / "build_power_query_lineage_report.py"
    for script in [bundle_script, inspect_script, cube_report_script, readiness_script, pq_lineage_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_fixture_bundle_") as tmp:
        tmp_dir = Path(tmp)
        bundle_dir = tmp_dir / "bundle"
        manifest_json = bundle_dir / "fixture-bundle.json"
        readme_md = bundle_dir / "README.md"
        cube_openxml = tmp_dir / "cube_openxml.json"
        cube_report = tmp_dir / "cube_report.json"
        external_openxml = tmp_dir / "external_openxml.json"
        external_readiness = tmp_dir / "external_readiness.json"
        pure_openxml = tmp_dir / "pure_openxml.json"
        pure_readiness = tmp_dir / "pure_readiness.json"
        pq_safe_report = tmp_dir / "pq_lineage_safe.json"
        pq_risky_report = tmp_dir / "pq_lineage_risky.json"

        steps = [
            [
                sys.executable,
                str(bundle_script),
                "--out-dir",
                str(bundle_dir),
                "--clean",
                "--out-json",
                str(manifest_json),
                "--out-md",
                str(readme_md),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(bundle_dir / "cube_formula_fixture.xlsx"),
                "--out-json",
                str(cube_openxml),
            ],
            [
                sys.executable,
                str(cube_report_script),
                "--openxml-json",
                str(cube_openxml),
                "--model-json",
                str(bundle_dir / "cube_model_summary.json"),
                "--out-json",
                str(cube_report),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(bundle_dir / "external_dependency_fixture.xlsx"),
                "--out-json",
                str(external_openxml),
            ],
            [
                sys.executable,
                str(readiness_script),
                "--openxml-json",
                str(external_openxml),
                "--out-json",
                str(external_readiness),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(bundle_dir / "pure_deliverable_fixture.xlsx"),
                "--out-json",
                str(pure_openxml),
            ],
            [
                sys.executable,
                str(readiness_script),
                "--openxml-json",
                str(pure_openxml),
                "--out-json",
                str(pure_readiness),
            ],
            [
                sys.executable,
                str(pq_lineage_script),
                str(bundle_dir / "power_query_lineage" / "safe"),
                "--out-json",
                str(pq_safe_report),
                "--fail-on-high-risk",
            ],
            [
                sys.executable,
                str(pq_lineage_script),
                str(bundle_dir / "power_query_lineage" / "risky"),
                "--out-json",
                str(pq_risky_report),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            manifest = json.loads(manifest_json.read_text(encoding="utf-8"))
            readme = readme_md.read_text(encoding="utf-8")
            cube = json.loads(cube_report.read_text(encoding="utf-8"))
            external_meta = json.loads((bundle_dir / "external_dependency_fixture.json").read_text(encoding="utf-8"))
            external = json.loads(external_readiness.read_text(encoding="utf-8"))
            pure = json.loads(pure_readiness.read_text(encoding="utf-8"))
            pq_lineage_meta = json.loads((bundle_dir / "power_query_lineage_fixture.json").read_text(encoding="utf-8"))
            pq_safe = json.loads(pq_safe_report.read_text(encoding="utf-8"))
            pq_risky = json.loads(pq_risky_report.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read generated bundle reports: {exc}")

        fixtures = manifest.get("fixtures", [])
        fixture_ids = {str(item.get("id", "")) for item in fixtures if isinstance(item, dict)}
        expected_ids = {"cube-formula", "external-dependency", "pure-deliverable", "power-query-lineage"}
        expected_files = [
            bundle_dir / "cube_formula_fixture.xlsx",
            bundle_dir / "cube_model_summary.json",
            bundle_dir / "external_dependency_fixture.xlsx",
            bundle_dir / "external_dependency_fixture.json",
            bundle_dir / "pure_deliverable_fixture.xlsx",
            bundle_dir / "pure_deliverable_fixture.json",
            bundle_dir / "power_query_lineage_fixture.json",
            bundle_dir / "power_query_lineage" / "safe" / "power_queries.json",
            bundle_dir / "power_query_lineage" / "risky" / "power_queries.json",
            manifest_json,
            readme_md,
        ]
        failures: list[str] = []
        missing_files = [path.name for path in expected_files if not path.is_file()]
        if missing_files:
            failures.append(f"missing generated files={missing_files}")
        if manifest.get("fixtureCount") != 4:
            failures.append(f"fixtureCount={manifest.get('fixtureCount')}")
        if fixture_ids != expected_ids:
            failures.append(f"fixture ids={sorted(fixture_ids)}")
        external_fixture_entries = [
            item
            for item in fixtures
            if isinstance(item, dict) and item.get("id") == "external-dependency"
        ]
        external_fixture_entry = external_fixture_entries[0] if external_fixture_entries else {}
        external_expected_evidence = set(external_fixture_entry.get("expectedEvidence", [])) if isinstance(external_fixture_entry, dict) else set()
        for expected_phrase in ["2 workbook connections", "1 redacted credential-like connection indicator"]:
            if expected_phrase not in external_expected_evidence:
                failures.append(f"external fixture manifest missing expected evidence={expected_phrase}")
        if "Sanitized Excel BI Fixture Bundle" not in readme:
            failures.append("README title missing")
        if "2 workbook connections" not in readme:
            failures.append("README missing updated external connection evidence")
        if "1 redacted credential-like connection indicator" not in readme:
            failures.append("README missing credential indicator evidence")
        if "power-query-lineage" not in readme:
            failures.append("README missing Power Query lineage fixture")
        if "safe exported-query set with 4 queries and 0 findings" not in readme:
            failures.append("README missing Power Query safe evidence")
        if "native SQL, credential-like literal, mixed-source lineage, and query cycle diagnostics" not in readme:
            failures.append("README missing Power Query risky evidence")
        if cube.get("cubeFormulaCount") != 7:
            failures.append(f"cubeFormulaCount={cube.get('cubeFormulaCount')}")
        if cube.get("missingModelMeasures") != ["Missing Measure"]:
            failures.append("missing model-measure diagnostic mismatch")
        external_expected = external_meta.get("expected", {}) if isinstance(external_meta, dict) else {}
        if external_expected.get("connectionCount") != 2:
            failures.append(f"external metadata connectionCount={external_expected.get('connectionCount')}")
        if external_expected.get("credentialLikeConnectionCount") != 1:
            failures.append(f"external metadata credentialLikeConnectionCount={external_expected.get('credentialLikeConnectionCount')}")
        external_summary = external.get("summary", {})
        if external.get("summary", {}).get("readiness") != "blocked-for-pure-deliverable":
            failures.append(f"external readiness={external.get('summary', {}).get('readiness')}")
        if external_summary.get("connectionCount") != 2:
            failures.append(f"external readiness connectionCount={external_summary.get('connectionCount')}")
        if external_summary.get("credentialLikeConnectionCount") != 1:
            failures.append(f"external readiness credentialLikeConnectionCount={external_summary.get('credentialLikeConnectionCount')}")
        external_codes = {finding.get("code") for finding in external.get("findings", [])}
        if "connection-credential-like-literal" not in external_codes:
            failures.append(f"external readiness codes missing connection-credential-like-literal: {sorted(external_codes)}")
        if pure.get("summary", {}).get("readiness") != "clean":
            failures.append(f"pure readiness={pure.get('summary', {}).get('readiness')}")
        pq_safe_expected = pq_lineage_meta.get("safe", {}).get("expected", {}) if isinstance(pq_lineage_meta, dict) else {}
        pq_risky_expected = pq_lineage_meta.get("risky", {}).get("expected", {}) if isinstance(pq_lineage_meta, dict) else {}
        pq_safe_summary = pq_safe.get("summary", {})
        pq_risky_summary = pq_risky.get("summary", {})
        pq_risky_codes = {finding.get("code") for finding in pq_risky.get("findings", [])}
        if pq_safe_expected.get("queryCount") != 4:
            failures.append(f"pq safe metadata queryCount={pq_safe_expected.get('queryCount')}")
        if pq_risky_expected.get("queryCount") != 11:
            failures.append(f"pq risky metadata queryCount={pq_risky_expected.get('queryCount')}")
        for label, summary, expected in [
            ("pq safe", pq_safe_summary, pq_safe_expected),
            ("pq risky", pq_risky_summary, pq_risky_expected),
        ]:
            for key in ["readiness", "queryCount", "dependencyCount"]:
                if summary.get(key) != expected.get(key):
                    failures.append(f"{label} {key}={summary.get(key)}")
        if pq_safe_summary.get("findingCount") != pq_safe_expected.get("findingCount"):
            failures.append(f"pq safe findingCount={pq_safe_summary.get('findingCount')}")
        if pq_risky_summary.get("highFindingCount") != pq_risky_expected.get("highFindingCount"):
            failures.append(f"pq risky highFindingCount={pq_risky_summary.get('highFindingCount')}")
        if pq_risky_summary.get("mediumFindingCount") != pq_risky_expected.get("mediumFindingCount"):
            failures.append(f"pq risky mediumFindingCount={pq_risky_summary.get('mediumFindingCount')}")
        for kind, expected_count in pq_safe_expected.get("sourceKindCounts", {}).items():
            actual_count = pq_safe_summary.get("sourceKindCounts", {}).get(kind)
            if actual_count != expected_count:
                failures.append(f"pq safe sourceKindCounts[{kind}]={actual_count}")
        for kind, expected_count in pq_risky_expected.get("sourceKindCounts", {}).items():
            actual_count = pq_risky_summary.get("sourceKindCounts", {}).get(kind)
            if actual_count != expected_count:
                failures.append(f"pq risky sourceKindCounts[{kind}]={actual_count}")
        missing_pq_codes = set(pq_risky_expected.get("requiredCodes", [])) - pq_risky_codes
        if missing_pq_codes:
            failures.append(f"pq risky fixture missing codes={sorted(missing_pq_codes)}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="4 sanitized fixtures generated and workbook/query-source evidence verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "fixtureIds": sorted(fixture_ids),
                "cubeFormulaCount": cube.get("cubeFormulaCount"),
                "externalReadiness": external.get("summary", {}).get("readiness"),
                "externalConnectionCount": external_summary.get("connectionCount"),
                "externalCredentialLikeConnectionCount": external_summary.get("credentialLikeConnectionCount"),
                "pureReadiness": pure.get("summary", {}).get("readiness"),
                "pqSafeQueryCount": pq_safe_summary.get("queryCount"),
                "pqRiskyHighFindingCount": pq_risky_summary.get("highFindingCount"),
                "pqRiskyMediumFindingCount": pq_risky_summary.get("mediumFindingCount"),
                "pqRiskyCodes": sorted(pq_risky_codes),
            },
        )


def release_evidence_bundle_check(project_root: Path) -> CheckResult:
    name = "Release evidence bundle smoke"
    script = project_root / "tools" / "build_release_evidence_bundle.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_release_evidence_") as tmp:
        tmp_dir = Path(tmp)
        out_json = tmp_dir / "release_evidence.json"
        out_md = tmp_dir / "release_evidence.md"
        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
            markdown = out_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read evidence bundle: {exc}")

        failures: list[str] = []
        statuses = report.get("statuses", {})
        expected_statuses = {
            "projectDocs": PASS,
            "taskRecipes": PASS,
            "officialDocs": PASS,
            "goalCoverage": PASS,
            "releaseGate": "not-supplied",
        }
        for key, expected in expected_statuses.items():
            if statuses.get(key) != expected:
                failures.append(f"{key}={statuses.get(key)}")
        plugin = report.get("plugin", {})
        if plugin.get("name") != "microsoft-excel-bi-agent-pack":
            failures.append(f"plugin name={plugin.get('name')}")
        if not CODEX_CACHEBUSTER_RE.match(str(plugin.get("version", ""))):
            failures.append(f"plugin version={plugin.get('version')}")
        if "Release Evidence Bundle" not in markdown:
            failures.append("Markdown title missing")
        if "expected cache" not in markdown:
            failures.append("Markdown expected cache path missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="project docs, recipes, official docs, and goal coverage summarized" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "statuses": statuses,
                "version": plugin.get("version"),
                "releaseGateIncluded": report.get("releaseGate", {}).get("included"),
            },
        )


def real_sanitized_case_regression_check(project_root: Path) -> CheckResult:
    name = "Real/sanitized case regression library smoke"
    script = project_root / "tools" / "run_case_regression.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_case_regression_") as tmp:
        tmp_dir = Path(tmp)
        out_json = tmp_dir / "case_regression.json"
        out_md = tmp_dir / "case_regression.md"
        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read case regression report: {exc}")

        failures: list[str] = []
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        if int(report.get("caseCount", 0) or 0) < 6:
            failures.append(f"caseCount={report.get('caseCount')}")
        covered_layers = set(report.get("coveredLayers") or [])
        expected_layers = {"power-query", "dax", "cube-mdx", "vba", "deliverable", "visual-qa"}
        missing_layers = sorted(expected_layers - covered_layers)
        if missing_layers:
            failures.append(f"missingLayers={','.join(missing_layers)}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="6 sanitized regression case specs cover Power Query, DAX, CUBE/MDX, VBA, deliverable, and visual QA layers"
            if not failures
            else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "caseCount": report.get("caseCount"),
                "coveredLayers": report.get("coveredLayers"),
                "evidenceModes": report.get("evidenceModes"),
            },
        )


def excel_bi_router_fixture_check(project_root: Path) -> CheckResult:
    name = "Excel BI router fixture smoke"
    script = project_root / ".agents" / "skills" / "excel-bi-router" / "scripts" / "route_excel_bi_task.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    cases = [
        {
            "id": "workbook-vba",
            "text": "Workbook has xlsm buttons and VBA macro OnAction stopped working after copying sheets.",
            "skill": "excel-vba-workbook-engineering",
            "layer": "Workbook/VBA",
        },
        {
            "id": "power-query",
            "text": "Power Query Folder.Files combine query refresh fails with Formula.Firewall and missing column after source schema drift.",
            "skill": "power-query-m-engineering",
            "layer": "Power Query M",
        },
        {
            "id": "power-pivot-dax",
            "text": "Power Pivot Data Model DAX measure using CALCULATE and REMOVEFILTERS must work in Excel Power Pivot.",
            "skill": "power-pivot-dax-modeling",
            "layer": "Power Pivot DAX",
        },
        {
            "id": "mdx-cube",
            "text": "CUBEVALUE formulas and CUBEMEMBER helper cells need MDX dependency mapping for [Measures].[Revenue].",
            "skill": "mdx-cubevalue-extraction",
            "layer": "MDX/CUBE",
        },
        {
            "id": "ado-sql",
            "text": "VBA ADODB SQL connection string queries workbook tables with ACE OLEDB provider.",
            "skill": "excel-ado-sql-data-access",
            "layer": "ADO/SQL",
        },
        {
            "id": "mixed",
            "text": "The workbook has VBA buttons, Power Query refresh, Data Model DAX measures, and CUBEVALUE report formulas.",
            "skill": "excel-bi-router",
            "layer": "Mixed",
        },
    ]

    with tempfile.TemporaryDirectory(prefix="excel_bi_router_") as tmp:
        tmp_dir = Path(tmp)
        failures: list[str] = []
        reports: list[dict[str, object]] = []
        for case in cases:
            out_json = tmp_dir / f"{case['id']}.json"
            out_md = tmp_dir / f"{case['id']}.md"
            result = run_command(
                [
                    sys.executable,
                    str(script),
                    "--text",
                    str(case["text"]),
                    "--out-json",
                    str(out_json),
                    "--out-md",
                    str(out_md),
                    "--expect-skill",
                    str(case["skill"]),
                    "--expect-layer",
                    str(case["layer"]),
                ],
                project_root,
                f"{name}: {case['id']}",
            )
            if result.status != PASS:
                failures.append(f"{case['id']} command failed: {result.detail}")
                continue
            try:
                report = json.loads(out_json.read_text(encoding="utf-8"))
                markdown = out_md.read_text(encoding="utf-8")
            except Exception as exc:
                failures.append(f"{case['id']} output unreadable: {exc}")
                continue
            reports.append(report)
            if report.get("skill") != case["skill"]:
                failures.append(f"{case['id']} skill={report.get('skill')}")
            if report.get("layer") != case["layer"]:
                failures.append(f"{case['id']} layer={report.get('layer')}")
            if not report.get("validationNeeded"):
                failures.append(f"{case['id']} missing validationNeeded")
            if not report.get("recommendedScripts"):
                failures.append(f"{case['id']} missing recommendedScripts")
            if "# Excel BI Task Route" not in markdown:
                failures.append(f"{case['id']} markdown heading missing")
            if "Boundaries" not in markdown:
                failures.append(f"{case['id']} markdown boundaries missing")

        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))

        return CheckResult(
            name=name,
            status=PASS,
            detail="router selected expected layer and skill for Workbook/VBA, Power Query, DAX, MDX/CUBE, ADO/SQL, and mixed tasks",
            metadata={
                "caseCount": len(cases),
                "skills": sorted({str(item.get("skill")) for item in reports}),
                "layers": sorted({str(item.get("layer")) for item in reports}),
            },
        )


def capability_catalog_fixture_check(project_root: Path) -> CheckResult:
    name = "Capability catalog fixture smoke"
    script = project_root / "tools" / "build_capability_catalog.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_capability_catalog_") as tmp:
        tmp_dir = Path(tmp)
        report_json = tmp_dir / "capability_catalog.json"
        report_md = tmp_dir / "capability_catalog.md"
        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        try:
            catalog = json.loads(report_json.read_text(encoding="utf-8"))
            markdown = report_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read capability catalog: {exc}")

        skills = {str(item.get("name", "")) for item in catalog.get("skills", []) if isinstance(item, dict)}
        tools = {str(item.get("name", "")) for item in catalog.get("tools", []) if isinstance(item, dict)}
        workflows = {str(item.get("id", "")) for item in catalog.get("workflows", []) if isinstance(item, dict)}
        summary = catalog.get("summary", {})
        failures: list[str] = []
        required_skills = {
            "excel-ado-sql-data-access",
            "excel-bi-router",
            "excel-deliverable-publisher",
            "excel-report-builder",
            "excel-testing-fixtures",
            "excel-vba-workbook-engineering",
            "excel-workbook-qa-auditor",
            "mdx-cubevalue-extraction",
            "office-environment-diagnostics",
            "power-bi-semantic-model",
            "power-pivot-dax-modeling",
            "power-query-m-engineering",
        }
        missing_skills = sorted(required_skills - skills)
        if missing_skills:
            failures.append(f"missing skills={missing_skills}")
        for required_tool in [
            "inspect_excel_bi_workbook.py",
            "build_workbook_triage_report.py",
            "build_power_query_refresh_report.py",
            "run_release_gate.py",
            "deploy-local-plugin.py",
        ]:
            if required_tool not in tools:
                failures.append(f"missing tool={required_tool}")
        for required_workflow in [
            "route-then-triage",
            "power-query-lifecycle",
            "power-pivot-cube",
            "pure-deliverable",
            "workbook-qa-audit",
            "office-environment-diagnostics",
            "excel-report-build",
            "power-bi-semantic-model-review",
            "sanitized-fixture-regression",
            "cross-agent-release",
        ]:
            if required_workflow not in workflows:
                failures.append(f"missing workflow={required_workflow}")
        if int(summary.get("officialDocEntryCount") or 0) < 20:
            failures.append(f"officialDocEntryCount={summary.get('officialDocEntryCount')}")
        if int(summary.get("releaseGateCheckCount") or 0) < 20:
            failures.append(f"releaseGateCheckCount={summary.get('releaseGateCheckCount')}")
        if "# Excel BI Capability Catalog" not in markdown or "## Boundary" not in markdown:
            failures.append("markdown heading or boundary missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="skills, tools, workflows, docs indexes, and release-gate check inventory verified" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "skillCount": summary.get("skillCount"),
                "toolCount": summary.get("toolCount"),
                "officialDocEntryCount": summary.get("officialDocEntryCount"),
                "releaseGateCheckCount": summary.get("releaseGateCheckCount"),
                "workflowCount": len(workflows),
            },
        )


def agent_bootstrap_bundle_fixture_check(project_root: Path) -> CheckResult:
    name = "Agent bootstrap bundle fixture smoke"
    script = project_root / "tools" / "build_agent_bootstrap_bundle.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_agent_bootstrap_") as tmp:
        tmp_dir = Path(tmp)
        bundle_dir = tmp_dir / "bundle"
        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-dir",
                str(bundle_dir),
                "--clean",
                "--zip",
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        manifest_path = bundle_dir / "bootstrap-manifest.json"
        bootstrap_path = bundle_dir / "BOOTSTRAP.md"
        zip_path = bundle_dir / "agent-bootstrap-bundle.zip"
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            bootstrap = bootstrap_path.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read bootstrap bundle: {exc}")

        failures: list[str] = []
        if manifest.get("status") != PASS:
            failures.append(f"manifest status={manifest.get('status')}")
        plugin = manifest.get("plugin", {})
        if plugin.get("name") != "microsoft-excel-bi-agent-pack":
            failures.append(f"plugin name={plugin.get('name')}")
        summary = manifest.get("capabilityCatalog", {}).get("summary", {})
        if int(summary.get("skillCount") or 0) < 12:
            failures.append(f"skillCount={summary.get('skillCount')}")
        if int(summary.get("toolCount") or 0) < 60:
            failures.append(f"toolCount={summary.get('toolCount')}")
        if int(summary.get("officialDocEntryCount") or 0) < 20:
            failures.append(f"officialDocEntryCount={summary.get('officialDocEntryCount')}")
        release_statuses = manifest.get("releaseEvidence", {}).get("statuses", {})
        for required_status in ["projectDocs", "taskRecipes", "officialDocs", "goalCoverage"]:
            if release_statuses.get(required_status) != PASS:
                failures.append(f"{required_status}={release_statuses.get(required_status)}")
        for required in [
            "Agent Bootstrap Bundle",
            "onboarding infrastructure",
            "not proof of external-agent behavior",
            "validation-commands.md",
            "capability-catalog.md",
        ]:
            if required not in bootstrap:
                failures.append(f"BOOTSTRAP.md missing {required}")
        for required_file in [
            "BOOTSTRAP.md",
            "bootstrap-manifest.json",
            "capability-catalog.json",
            "capability-catalog.md",
            "release-evidence.json",
            "release-evidence.md",
            "task-recipes.md",
            "validation-commands.md",
        ]:
            if not (bundle_dir / required_file).is_file():
                failures.append(f"missing file={required_file}")
        if not zip_path.is_file():
            failures.append("zip archive missing")
        else:
            with zipfile.ZipFile(zip_path) as archive:
                names = set(archive.namelist())
            for required_file in [
                "BOOTSTRAP.md",
                "bootstrap-manifest.json",
                "capability-catalog.json",
                "release-evidence.json",
                "task-recipes.md",
                "validation-commands.md",
            ]:
                if required_file not in names:
                    failures.append(f"zip archive missing {required_file}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="fresh-agent bootstrap bundle manifest, docs, validation commands, and archive verified" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "skillCount": summary.get("skillCount"),
                "toolCount": summary.get("toolCount"),
                "officialDocEntryCount": summary.get("officialDocEntryCount"),
                "releaseGateCheckCount": summary.get("releaseGateCheckCount"),
                "fileCount": len(manifest.get("files", [])),
            },
        )


def completion_readiness_audit_check(project_root: Path) -> CheckResult:
    name = "Completion readiness audit smoke"
    script = project_root / "tools" / "build_completion_readiness_audit.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_completion_readiness_") as tmp:
        tmp_dir = Path(tmp)
        report_json = tmp_dir / "completion_readiness.json"
        report_md = tmp_dir / "completion_readiness.md"
        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result
        try:
            report = json.loads(report_json.read_text(encoding="utf-8"))
            markdown = report_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read completion readiness report: {exc}")

        failures: list[str] = []
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        coverage = report.get("coverage", {})
        if coverage.get("status") != PASS or int(coverage.get("passedAreaCount") or 0) < 11:
            failures.append(f"coverage={coverage}")
        blocker_codes = {str(item.get("code", "")) for item in report.get("blockers", []) if isinstance(item, dict)}
        completion_ready = report.get("completionReady")
        if completion_ready is True:
            if report.get("readinessStatus") != "ready":
                failures.append(f"readinessStatus={report.get('readinessStatus')}")
            if blocker_codes:
                failures.append(f"unexpected blockers={sorted(blocker_codes)}")
            if "# Completion Readiness Audit" not in markdown or "completion ready: `True`" not in markdown:
                failures.append("Markdown heading or ready-state marker missing")
        elif completion_ready is False:
            if report.get("readinessStatus") != "in-progress":
                failures.append(f"readinessStatus={report.get('readinessStatus')}")
            if not blocker_codes:
                failures.append("in-progress audit has no blockers")
            allowed_in_progress_blockers = {
                "completion-evidence-first-pass",
                "master-goal-not-complete",
                "latest-progress-missing-goal-state",
                "validation-version-mismatch",
            }
            unexpected_blockers = sorted(code for code in blocker_codes if code not in allowed_in_progress_blockers)
            if unexpected_blockers:
                failures.append(f"unexpected blockers={', '.join(unexpected_blockers)}")
            if "# Completion Readiness Audit" not in markdown or "completion ready: `False`" not in markdown:
                failures.append("Markdown heading or in-progress marker missing")
        else:
            failures.append(f"completionReady={completion_ready}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="readiness audit proves public maintenance coverage and active backlog state" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "readinessStatus": report.get("readinessStatus"),
                "completionReady": report.get("completionReady"),
                "blockerCount": len(report.get("blockers", [])),
                "passedAreaCount": coverage.get("passedAreaCount"),
                "areaCount": coverage.get("areaCount"),
            },
        )


def cross_agent_forward_test_pack_check(project_root: Path) -> CheckResult:
    name = "Cross-agent forward-test pack smoke"
    script = project_root / "tools" / "build_cross_agent_forward_test_pack.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_forward_test_pack_") as tmp:
        tmp_dir = Path(tmp)
        pack_dir = tmp_dir / "pack"
        out_json = tmp_dir / "forward_test_pack.json"
        out_md = tmp_dir / "forward_test_pack.md"
        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-dir",
                str(pack_dir),
                "--clean",
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        try:
            manifest = json.loads(out_json.read_text(encoding="utf-8"))
            markdown = out_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read forward-test pack: {exc}")

        failures: list[str] = []
        expected_agents = {"codex", "claude", "opencode", "generic"}
        expected_skills = {
            "excel-ado-sql-data-access",
            "excel-bi-router",
            "excel-deliverable-publisher",
            "excel-report-builder",
            "excel-testing-fixtures",
            "excel-vba-workbook-engineering",
            "excel-workbook-qa-auditor",
            "mdx-cubevalue-extraction",
            "office-environment-diagnostics",
            "power-bi-semantic-model",
            "power-pivot-dax-modeling",
            "power-query-m-engineering",
        }
        agent_ids = {str(item.get("id")) for item in manifest.get("agentTargets", []) if isinstance(item, dict)}
        skill_ids = {str(item.get("name")) for item in manifest.get("skills", []) if isinstance(item, dict)}
        prompts = [item for item in manifest.get("prompts", []) if isinstance(item, dict)]
        if manifest.get("status") != PASS:
            failures.append(f"status={manifest.get('status')}")
        if agent_ids != expected_agents:
            failures.append(f"agent ids={sorted(agent_ids)}")
        if skill_ids != expected_skills:
            failures.append(f"skill ids={sorted(skill_ids)}")
        if manifest.get("agentTargetCount") != len(expected_agents):
            failures.append(f"agentTargetCount={manifest.get('agentTargetCount')}")
        if manifest.get("skillCount") != len(expected_skills):
            failures.append(f"skillCount={manifest.get('skillCount')}")
        if manifest.get("promptCount") != len(expected_agents) * len(expected_skills):
            failures.append(f"promptCount={manifest.get('promptCount')}")
        if len(prompts) != len(expected_agents) * len(expected_skills):
            failures.append(f"prompt entries={len(prompts)}")
        if "Cross-Agent Forward Test Pack" not in markdown:
            failures.append("Markdown title missing")
        if "customer-data-free" not in markdown:
            failures.append("Markdown customer-data boundary missing")

        sample_prompt = pack_dir / "prompts" / "generic" / "power-pivot-dax-modeling.md"
        if not sample_prompt.is_file():
            failures.append(f"sample prompt missing={sample_prompt}")
        else:
            sample_text = sample_prompt.read_text(encoding="utf-8")
            for required in ["REMOVEFILTERS", "DIVIDE", "Expected Evidence", "Excel COM"]:
                if required not in sample_text:
                    failures.append(f"sample prompt missing {required}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="48 forward-test prompts generated for Codex, Claude, OpenCode, and generic agents" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "agentTargetCount": manifest.get("agentTargetCount"),
                "skillCount": manifest.get("skillCount"),
                "promptCount": manifest.get("promptCount"),
                "agentIds": sorted(agent_ids),
                "skillIds": sorted(skill_ids),
            },
        )


def cross_agent_forward_test_result_scorer_check(project_root: Path) -> CheckResult:
    name = "Cross-agent forward-test result scorer smoke"
    pack_script = project_root / "tools" / "build_cross_agent_forward_test_pack.py"
    scorer_script = project_root / "tools" / "score_cross_agent_forward_test_results.py"
    missing = [str(path) for path in [pack_script, scorer_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_forward_test_score_") as tmp:
        tmp_dir = Path(tmp)
        pack_dir = tmp_dir / "pack"
        manifest_json = tmp_dir / "forward_test_pack.json"
        pack_md = tmp_dir / "forward_test_pack.md"
        pass_json = tmp_dir / "score_pass.json"
        pass_md = tmp_dir / "score_pass.md"
        fail_json = tmp_dir / "score_fail.json"
        fail_md = tmp_dir / "score_fail.md"

        build_result = run_command(
            [
                sys.executable,
                str(pack_script),
                "--project-root",
                str(project_root),
                "--out-dir",
                str(pack_dir),
                "--clean",
                "--out-json",
                str(manifest_json),
                "--out-md",
                str(pack_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if build_result.status != PASS:
            return build_result

        pass_result = run_command(
            [
                sys.executable,
                str(scorer_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(tmp_dir / "responses_pass"),
                "--clean-responses",
                "--write-passing-fixture",
                "--out-json",
                str(pass_json),
                "--out-md",
                str(pass_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if pass_result.status != PASS:
            return pass_result

        fail_result = run_command(
            [
                sys.executable,
                str(scorer_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(tmp_dir / "responses_fail"),
                "--clean-responses",
                "--write-failing-fixture",
                "--out-json",
                str(fail_json),
                "--out-md",
                str(fail_md),
                "--require-pass",
            ],
            project_root,
            name,
            ok_codes={1},
        )
        if fail_result.status != PASS:
            return fail_result

        failures: list[str] = []
        try:
            pass_report = json.loads(pass_json.read_text(encoding="utf-8"))
            fail_report = json.loads(fail_json.read_text(encoding="utf-8"))
            pass_markdown = pass_md.read_text(encoding="utf-8")
            fail_markdown = fail_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read scorer report: {exc}")

        if pass_report.get("status") != PASS:
            failures.append(f"passing fixture status={pass_report.get('status')}")
        if pass_report.get("expectedResponseCount") != 48:
            failures.append(f"passing expectedResponseCount={pass_report.get('expectedResponseCount')}")
        if pass_report.get("passedCount") != 48:
            failures.append(f"passing passedCount={pass_report.get('passedCount')}")
        if pass_report.get("failedCount") != 0:
            failures.append(f"passing failedCount={pass_report.get('failedCount')}")
        if fail_report.get("status") != FAIL:
            failures.append(f"failing fixture status={fail_report.get('status')}")
        if fail_report.get("passedCount") != 47:
            failures.append(f"failing passedCount={fail_report.get('passedCount')}")
        if fail_report.get("failedCount") != 1:
            failures.append(f"failing failedCount={fail_report.get('failedCount')}")
        for label, markdown in [("pass", pass_markdown), ("fail", fail_markdown)]:
            if "Cross-Agent Forward-Test Result Score" not in markdown:
                failures.append(f"{label} Markdown title missing")
            if "fresh agent-session outputs" not in markdown:
                failures.append(f"{label} Markdown boundary missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="response scorer accepts complete synthetic evidence and blocks incomplete evidence" if not failures else "; ".join(failures),
            stdout="\n".join([build_result.stdout, pass_result.stdout, fail_result.stdout]).strip(),
            stderr="\n".join([build_result.stderr, pass_result.stderr, fail_result.stderr]).strip(),
            metadata={
                "passingExpectedResponseCount": pass_report.get("expectedResponseCount"),
                "passingPassedCount": pass_report.get("passedCount"),
                "failingPassedCount": fail_report.get("passedCount"),
                "failingFailedCount": fail_report.get("failedCount"),
            },
        )


def cross_agent_forward_test_runbook_check(project_root: Path) -> CheckResult:
    name = "Cross-agent forward-test runbook smoke"
    pack_script = project_root / "tools" / "build_cross_agent_forward_test_pack.py"
    runbook_script = project_root / "tools" / "build_cross_agent_forward_test_runbook.py"
    missing = [str(path) for path in [pack_script, runbook_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_forward_test_runbook_") as tmp:
        tmp_dir = Path(tmp)
        pack_dir = tmp_dir / "pack"
        manifest_json = pack_dir / "forward-test-pack.json"
        pack_md = pack_dir / "README.md"
        responses_dir = tmp_dir / "responses"
        runbook_dir = tmp_dir / "runbook"
        out_json = runbook_dir / "runbook.json"
        out_md = runbook_dir / "RUNBOOK.copy.md"

        build_result = run_command(
            [
                sys.executable,
                str(pack_script),
                "--project-root",
                str(project_root),
                "--out-dir",
                str(pack_dir),
                "--clean",
                "--out-json",
                str(manifest_json),
                "--out-md",
                str(pack_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if build_result.status != PASS:
            return build_result

        runbook_result = run_command(
            [
                sys.executable,
                str(runbook_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(responses_dir),
                "--out-dir",
                str(runbook_dir),
                "--clean",
                "--write-response-stubs",
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if runbook_result.status != PASS:
            return runbook_result

        failures: list[str] = []
        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
            markdown = (runbook_dir / "RUNBOOK.md").read_text(encoding="utf-8")
            copied_markdown = out_md.read_text(encoding="utf-8")
            assignment_lines = (runbook_dir / "assignment-matrix.csv").read_text(encoding="utf-8").splitlines()
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read runbook output: {exc}")

        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        if report.get("assignmentCount") != 48:
            failures.append(f"assignmentCount={report.get('assignmentCount')}")
        if report.get("existingResponseCount") != 0:
            failures.append(f"existingResponseCount={report.get('existingResponseCount')}")
        if report.get("missingResponseCount") != 48:
            failures.append(f"missingResponseCount={report.get('missingResponseCount')}")
        if report.get("stubCount") != 48:
            failures.append(f"stubCount={report.get('stubCount')}")
        if len(assignment_lines) != 49:
            failures.append(f"assignment matrix lines={len(assignment_lines)}")
        for required in [
            "Cross-Agent Forward-Test Runbook",
            "responses/<agent>/<skill>.md",
            "score_cross_agent_forward_test_results.py",
            "does not prove external-agent behavior",
        ]:
            if required not in markdown:
                failures.append(f"runbook missing {required}")
            if required not in copied_markdown:
                failures.append(f"copied runbook missing {required}")
        sample_stub = responses_dir / "generic" / "power-pivot-dax-modeling.md"
        if not sample_stub.is_file():
            failures.append(f"sample response stub missing={sample_stub}")
        else:
            sample_text = sample_stub.read_text(encoding="utf-8")
            if "Fresh-agent response" not in sample_text or "runtime boundaries" not in sample_text:
                failures.append("sample response stub missing guidance")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="runbook, assignment matrix, response stubs, and scoring command generated" if not failures else "; ".join(failures),
            stdout="\n".join([build_result.stdout, runbook_result.stdout]).strip(),
            stderr="\n".join([build_result.stderr, runbook_result.stderr]).strip(),
            metadata={
                "assignmentCount": report.get("assignmentCount"),
                "missingResponseCount": report.get("missingResponseCount"),
                "stubCount": report.get("stubCount"),
                "assignmentMatrixLines": len(assignment_lines),
            },
        )


def cross_agent_forward_test_handoff_bundle_check(project_root: Path) -> CheckResult:
    name = "Cross-agent forward-test handoff bundle smoke"
    script = project_root / "tools" / "build_cross_agent_forward_test_handoff_bundle.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_forward_test_handoff_") as tmp:
        tmp_dir = Path(tmp)
        bundle_dir = tmp_dir / "handoff"
        zip_path = tmp_dir / "handoff.zip"
        out_json = tmp_dir / "handoff.copy.json"
        out_md = tmp_dir / "HANDOFF.copy.md"

        result = run_command(
            [
                sys.executable,
                str(script),
                "--project-root",
                str(project_root),
                "--out-dir",
                str(bundle_dir),
                "--clean",
                "--zip-path",
                str(zip_path),
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        failures: list[str] = []
        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
            markdown = out_md.read_text(encoding="utf-8")
            manifest = json.loads((bundle_dir / "handoff-manifest.json").read_text(encoding="utf-8"))
            with zipfile.ZipFile(zip_path) as archive:
                zip_names = set(archive.namelist())
        except (OSError, json.JSONDecodeError, zipfile.BadZipFile) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read handoff output: {exc}")

        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        if manifest.get("status") != PASS:
            failures.append(f"manifest status={manifest.get('status')}")
        if report.get("promptCount") != 48:
            failures.append(f"promptCount={report.get('promptCount')}")
        if report.get("assignmentCount") != 48:
            failures.append(f"assignmentCount={report.get('assignmentCount')}")
        if report.get("stubCount") != 48:
            failures.append(f"stubCount={report.get('stubCount')}")
        if report.get("missingResponseCount") != 48:
            failures.append(f"missingResponseCount={report.get('missingResponseCount')}")
        for required in [
            "Cross-Agent Fresh-Session Handoff Bundle",
            "fresh-session responses",
            "responses/<agent>/<skill>.md",
            "score_cross_agent_forward_test_results.py",
        ]:
            if required not in markdown:
                failures.append(f"handoff Markdown missing {required}")
        for required in [
            "pack/forward-test-pack.json",
            "runbook/RUNBOOK.md",
            "responses/codex/excel-bi-router.md",
            "HANDOFF.md",
            "handoff-manifest.json",
        ]:
            if required not in zip_names:
                failures.append(f"zip missing {required}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="fresh-session handoff bundle, response stubs, scoring command, and zip archive generated"
            if not failures
            else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "promptCount": report.get("promptCount"),
                "assignmentCount": report.get("assignmentCount"),
                "stubCount": report.get("stubCount"),
                "zipEntryCount": len(zip_names),
            },
        )


def cross_agent_response_collection_report_check(project_root: Path) -> CheckResult:
    name = "Cross-agent response collection report smoke"
    handoff_script = project_root / "tools" / "build_cross_agent_forward_test_handoff_bundle.py"
    scorer_script = project_root / "tools" / "score_cross_agent_forward_test_results.py"
    report_script = project_root / "tools" / "build_cross_agent_response_collection_report.py"
    missing = [str(path) for path in [handoff_script, scorer_script, report_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_response_collection_") as tmp:
        tmp_dir = Path(tmp)
        bundle_dir = tmp_dir / "handoff"
        handoff_json = tmp_dir / "handoff.json"
        handoff_md = tmp_dir / "HANDOFF.md"
        stub_report_json = tmp_dir / "stub_collection.json"
        stub_report_md = tmp_dir / "stub_collection.md"
        fixture_responses_dir = tmp_dir / "fixture_responses"
        fixture_score_json = tmp_dir / "fixture_score.json"
        fixture_score_md = tmp_dir / "fixture_score.md"
        fixture_report_json = tmp_dir / "fixture_collection.json"
        fixture_report_md = tmp_dir / "fixture_collection.md"

        handoff_result = run_command(
            [
                sys.executable,
                str(handoff_script),
                "--project-root",
                str(project_root),
                "--out-dir",
                str(bundle_dir),
                "--clean",
                "--out-json",
                str(handoff_json),
                "--out-md",
                str(handoff_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if handoff_result.status != PASS:
            return handoff_result

        manifest_json = bundle_dir / "pack" / "forward-test-pack.json"
        responses_dir = bundle_dir / "responses"
        stub_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(responses_dir),
                "--out-json",
                str(stub_report_json),
                "--out-md",
                str(stub_report_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if stub_result.status != PASS:
            return stub_result

        score_result = run_command(
            [
                sys.executable,
                str(scorer_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(fixture_responses_dir),
                "--clean-responses",
                "--write-passing-fixture",
                "--out-json",
                str(fixture_score_json),
                "--out-md",
                str(fixture_score_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if score_result.status != PASS:
            return score_result

        fixture_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(fixture_responses_dir),
                "--score-json",
                str(fixture_score_json),
                "--out-json",
                str(fixture_report_json),
                "--out-md",
                str(fixture_report_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if fixture_result.status != PASS:
            return fixture_result

        negative_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--manifest-json",
                str(manifest_json),
                "--responses-dir",
                str(fixture_responses_dir),
                "--score-json",
                str(fixture_score_json),
                "--require-external-proof",
            ],
            project_root,
            name,
            ok_codes={1},
        )
        if negative_result.status != PASS:
            return CheckResult(
                name=name,
                status=FAIL,
                detail="fixture evidence unexpectedly satisfied --require-external-proof",
                command=negative_result.command,
                stdout=negative_result.stdout,
                stderr=negative_result.stderr,
            )

        failures: list[str] = []
        try:
            stub_report = json.loads(stub_report_json.read_text(encoding="utf-8"))
            stub_markdown = stub_report_md.read_text(encoding="utf-8")
            fixture_report = json.loads(fixture_report_json.read_text(encoding="utf-8"))
            fixture_markdown = fixture_report_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read collection report output: {exc}")

        stub_counts = stub_report.get("counts", {})
        fixture_counts = fixture_report.get("counts", {})
        if stub_report.get("status") != PASS:
            failures.append(f"stub status={stub_report.get('status')}")
        if stub_report.get("evidenceStatus") != "collecting":
            failures.append(f"stub evidenceStatus={stub_report.get('evidenceStatus')}")
        if stub_report.get("externalProofReady") is not False:
            failures.append("stub externalProofReady is not false")
        if stub_counts.get("stubResponseCount") != 48:
            failures.append(f"stubResponseCount={stub_counts.get('stubResponseCount')}")
        if stub_counts.get("candidateFreshResponseCount") != 0:
            failures.append(f"stub candidateFreshResponseCount={stub_counts.get('candidateFreshResponseCount')}")

        if fixture_report.get("status") != PASS:
            failures.append(f"fixture status={fixture_report.get('status')}")
        if fixture_report.get("evidenceStatus") != "fixture-only":
            failures.append(f"fixture evidenceStatus={fixture_report.get('evidenceStatus')}")
        if fixture_report.get("externalProofReady") is not False:
            failures.append("fixture externalProofReady is not false")
        if fixture_counts.get("generatedFixtureResponseCount") != 48:
            failures.append(f"generatedFixtureResponseCount={fixture_counts.get('generatedFixtureResponseCount')}")
        if fixture_report.get("score", {}).get("allExpectedPassed") is not True:
            failures.append("fixture scorer did not pass all expected responses")

        for label, markdown in [("stub", stub_markdown), ("fixture", fixture_markdown)]:
            for required in [
                "Cross-Agent Response Collection Report",
                "responses/<agent>/<skill>.md",
                "not external proof",
            ]:
                if required not in markdown:
                    failures.append(f"{label} Markdown missing {required}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="response collection report distinguishes stubs, generated fixtures, scorer status, and external-proof boundary"
            if not failures
            else "; ".join(failures),
            stdout="\n".join([handoff_result.stdout, stub_result.stdout, score_result.stdout, fixture_result.stdout]).strip(),
            stderr="\n".join([handoff_result.stderr, stub_result.stderr, score_result.stderr, fixture_result.stderr]).strip(),
            metadata={
                "stubEvidenceStatus": stub_report.get("evidenceStatus"),
                "stubResponseCount": stub_counts.get("stubResponseCount"),
                "fixtureEvidenceStatus": fixture_report.get("evidenceStatus"),
                "generatedFixtureResponseCount": fixture_counts.get("generatedFixtureResponseCount"),
                "externalProofReady": fixture_report.get("externalProofReady"),
            },
        )


def official_docs_drift_report_check(project_root: Path) -> CheckResult:
    name = "Official documentation drift report"
    drift_script = project_root / "tools" / "build_official_docs_drift_report.py"
    if not drift_script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {drift_script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_official_docs_drift_") as tmp:
        tmp_root = Path(tmp)
        out_json = tmp_root / "official_docs_drift.json"
        out_md = tmp_root / "official_docs_drift.md"
        result = run_command(
            [
                sys.executable,
                str(drift_script),
                "--project-root",
                str(project_root),
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
            markdown = out_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read official docs drift report: {exc}")

        failures: list[str] = []
        summary = report.get("summary", {})
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        if summary.get("indexCount") != 4:
            failures.append(f"indexCount={summary.get('indexCount')}")
        if summary.get("entryCount") != 52:
            failures.append(f"entryCount={summary.get('entryCount')}")
        if summary.get("uniqueUrlCount") != 52:
            failures.append(f"uniqueUrlCount={summary.get('uniqueUrlCount')}")
        if report.get("onlineCheckEnabled"):
            failures.append("default drift report should be offline")
        if "Official Documentation Drift Report" not in markdown:
            failures.append("Markdown title missing")
        if "status: **pass**" not in markdown:
            failures.append("Markdown status missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="4 indexes, 52 entries, and 52 official Microsoft URLs inventoried offline" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "summary": summary,
                "onlineCheckEnabled": report.get("onlineCheckEnabled"),
            },
        )


def artifact_hygiene_report_check(project_root: Path) -> CheckResult:
    name = "Artifact hygiene report"
    hygiene_script = project_root / "tools" / "build_artifact_hygiene_report.py"
    if not hygiene_script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {hygiene_script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_artifact_hygiene_") as tmp:
        tmp_root = Path(tmp)
        out_json = tmp_root / "artifact_hygiene.json"
        out_md = tmp_root / "artifact_hygiene.md"
        result = run_command(
            [
                sys.executable,
                str(hygiene_script),
                "--project-root",
                str(project_root),
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
            markdown = out_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read artifact hygiene report: {exc}")

        failures: list[str] = []
        summary = report.get("summary", {})
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        if summary.get("issueCount") != 0:
            failures.append(f"issueCount={summary.get('issueCount')}")
        if summary.get("officeFileCount") != 0:
            failures.append(f"officeFileCount={summary.get('officeFileCount')}")
        if "Artifact Hygiene Report" not in markdown:
            failures.append("Markdown title missing")
        if "status: **pass**" not in markdown:
            failures.append("Markdown status missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="no customer artifacts, generated reports, locks, caches, or Office workbooks found in the public package" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "summary": summary,
            },
        )


def cube_formula_fixture_check(project_root: Path) -> CheckResult:
    name = "CUBE formula fixture smoke"
    with tempfile.TemporaryDirectory(prefix="excel_bi_cube_fixture_") as tmp:
        tmp_root = Path(tmp)
        workbook = tmp_root / "cube_formula_fixture.xlsx"
        model_json = tmp_root / "cube_model_summary.json"
        openxml_json = tmp_root / "openxml.json"
        report_json = tmp_root / "cube_report.json"
        report_md = tmp_root / "cube_report.md"

        steps = [
            [
                sys.executable,
                str(project_root / "tools" / "create_cube_formula_fixture.py"),
                "--workbook",
                str(workbook),
                "--model-json",
                str(model_json),
            ],
            [
                sys.executable,
                str(project_root / "tools" / "inspect_excel_bi_workbook.py"),
                str(workbook),
                "--out-json",
                str(openxml_json),
            ],
            [
                sys.executable,
                str(project_root / "tools" / "build_cube_dependency_report.py"),
                "--openxml-json",
                str(openxml_json),
                "--model-json",
                str(model_json),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            report = json.loads(report_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read report: {exc}")

        flags = report.get("byDiagnosticFlag", {})
        failures: list[str] = []
        if report.get("cubeFormulaCount") != 7:
            failures.append(f"expected 7 CUBE formulas, got {report.get('cubeFormulaCount')}")
        if report.get("missingModelMeasures") != ["Missing Measure"]:
            failures.append("expected exactly one missing model measure")
        if flags.get("measure_not_found_in_model") != 1:
            failures.append("expected one missing-measure diagnostic flag")
        if flags.get("hard_coded_period_marker") != 1:
            failures.append("expected one hard-coded-period diagnostic flag")
        if flags.get("dynamic_mdx_string", 0) < 1:
            failures.append("expected at least one dynamic-MDX diagnostic flag")
        if "$A$5" not in report.get("byHelperCellReference", {}):
            failures.append("expected dynamic helper cell $A$5")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="7 formulas and diagnostics verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "cubeFormulaCount": report.get("cubeFormulaCount"),
                "missingModelMeasures": report.get("missingModelMeasures"),
                "byDiagnosticFlag": flags,
            },
        )


def power_pivot_model_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Power Pivot model report fixture smoke"
    with tempfile.TemporaryDirectory(prefix="excel_bi_model_report_fixture_") as tmp:
        tmp_root = Path(tmp)
        workbook = tmp_root / "cube_formula_fixture.xlsx"
        model_json = tmp_root / "cube_model_summary.json"
        openxml_json = tmp_root / "openxml.json"
        report_json = tmp_root / "model_report.json"
        report_md = tmp_root / "model_report.md"

        steps = [
            [
                sys.executable,
                str(project_root / "tools" / "create_cube_formula_fixture.py"),
                "--workbook",
                str(workbook),
                "--model-json",
                str(model_json),
            ],
            [
                sys.executable,
                str(project_root / "tools" / "inspect_excel_bi_workbook.py"),
                str(workbook),
                "--out-json",
                str(openxml_json),
            ],
            [
                sys.executable,
                str(project_root / "tools" / "build_excel_bi_model_report.py"),
                "--model-json",
                str(model_json),
                "--openxml-json",
                str(openxml_json),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            report = json.loads(report_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read report: {exc}")

        measures = {item.get("name"): item for item in report.get("measures", [])}
        failures: list[str] = []
        if report.get("modelAvailable") is not True:
            failures.append("expected modelAvailable=True")
        if report.get("tableCount") != 3:
            failures.append(f"expected 3 model tables, got {report.get('tableCount')}")
        if report.get("relationshipCount") != 2:
            failures.append(f"expected 2 relationships, got {report.get('relationshipCount')}")
        if report.get("measureCount") != 2:
            failures.append(f"expected 2 measures, got {report.get('measureCount')}")
        if report.get("cubeFormulaCount") != 7:
            failures.append(f"expected 7 CUBE formulas, got {report.get('cubeFormulaCount')}")
        if measures.get("Revenue", {}).get("cubeFormulaReferenceCount") != 3:
            failures.append("expected Revenue to be referenced by 3 CUBE formulas")
        if measures.get("Awareness", {}).get("cubeFormulaReferenceCount") != 1:
            failures.append("expected Awareness to be referenced by 1 CUBE formula")
        if report.get("cubeFormulaReferencesMissingModelMeasure") != ["Missing Measure"]:
            failures.append("expected Missing Measure in CUBE references not found in model")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="model summary and CUBE usage verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "tableCount": report.get("tableCount"),
                "relationshipCount": report.get("relationshipCount"),
                "measureCount": report.get("measureCount"),
                "cubeFormulaCount": report.get("cubeFormulaCount"),
                "missingMeasureReferences": report.get("cubeFormulaReferencesMissingModelMeasure"),
            },
        )


def external_dependency_fixture_check(project_root: Path) -> CheckResult:
    name = "External dependency OpenXML fixture smoke"
    create_script = project_root / "tools" / "create_external_dependency_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    generic_inspect_script = project_root / ".agents" / "skills" / "excel-vba-workbook-engineering" / "scripts" / "inspect_openxml.py"
    for script in [create_script, inspect_script, generic_inspect_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_external_dependency_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "external_dependency_fixture.xlsx"
        fixture_json = tmp_dir / "fixture.json"
        bi_json = tmp_dir / "bi_inspect.json"
        generic_json = tmp_dir / "generic_inspect.json"

        steps = [
            [
                sys.executable,
                str(create_script),
                "--workbook",
                str(workbook),
                "--out-json",
                str(fixture_json),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(workbook),
                "--out-json",
                str(bi_json),
            ],
            [
                sys.executable,
                str(generic_inspect_script),
                str(workbook),
                "--out-json",
                str(generic_json),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            bi_report = json.loads(bi_json.read_text(encoding="utf-8"))
            generic_report = json.loads(generic_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read external dependency reports: {exc}")

        failures: list[str] = []
        connections = bi_report.get("connections", [])
        connection_names = {connection.get("name") for connection in connections if isinstance(connection, dict)}
        if len(connections) != 2:
            failures.append(f"BI inspector connection count={len(connections)}")
        if "SafePowerQueryLikeConnection" not in connection_names:
            failures.append(f"BI inspector missing SafePowerQueryLikeConnection, names={sorted(connection_names)}")
        if "CredentialLikeConnection" not in connection_names:
            failures.append(f"BI inspector missing CredentialLikeConnection, names={sorted(connection_names)}")
        if bi_report.get("totalFormulaCount") != 1:
            failures.append(f"BI inspector formula count={bi_report.get('totalFormulaCount')}")
        if len(bi_report.get("externalLinks", [])) < 1:
            failures.append("BI inspector did not detect external links")
        if bi_report.get("hasMashupLikeParts") is not True:
            failures.append("BI inspector did not detect mashup-like/custom XML parts")

        generic_connections = generic_report.get("connections", [])
        if len(generic_connections) != 2:
            failures.append(f"generic inspector connection count={len(generic_connections)}")
        if len(generic_report.get("externalLinks", [])) < 1:
            failures.append("generic inspector did not detect external links")
        if generic_report.get("worksheets", [{}])[0].get("formulaCount") != 1:
            failures.append("generic inspector did not detect worksheet formula")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="connections, external links, formulas, and mashup-like parts verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "biConnections": len(connections),
                "biExternalLinks": len(bi_report.get("externalLinks", [])),
                "biHasMashupLikeParts": bi_report.get("hasMashupLikeParts"),
                "genericConnections": len(generic_connections),
                "genericExternalLinks": len(generic_report.get("externalLinks", [])),
            },
        )


def workbook_surface_fixture_check(project_root: Path) -> CheckResult:
    name = "Workbook surface OpenXML fixture smoke"
    create_script = project_root / "tools" / "create_workbook_surface_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    for script in [create_script, inspect_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_workbook_surface_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "workbook_surface_fixture.xlsx"
        fixture_json = tmp_dir / "fixture.json"
        openxml_json = tmp_dir / "openxml.json"

        steps = [
            [
                sys.executable,
                str(create_script),
                "--workbook",
                str(workbook),
                "--out-json",
                str(fixture_json),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(workbook),
                "--out-json",
                str(openxml_json),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            fixture = json.loads(fixture_json.read_text(encoding="utf-8"))
            report = json.loads(openxml_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read workbook surface reports: {exc}")

        expected = fixture.get("expected", {})
        if not isinstance(expected, dict):
            expected = {}
        sheet_names = [item.get("name") for item in report.get("sheets", []) if isinstance(item, dict)]
        defined_names = {item.get("name") for item in report.get("definedNames", []) if isinstance(item, dict)}
        table_names = {item.get("displayName") for item in report.get("tables", []) if isinstance(item, dict)}
        formula_functions = set(report.get("formulaFunctionCounts", {}).keys())
        failures: list[str] = []

        if sheet_names != expected.get("sheets"):
            failures.append(f"sheets={sheet_names}")
        if report.get("totalFormulaCount") != expected.get("formulaCount"):
            failures.append(f"formulaCount={report.get('totalFormulaCount')}")
        expected_functions = set(expected.get("formulaFunctions", []))
        missing_functions = sorted(expected_functions - formula_functions)
        if missing_functions:
            failures.append(f"missing formula functions={missing_functions}")
        missing_names = sorted(set(expected.get("definedNames", [])) - defined_names)
        if missing_names:
            failures.append(f"missing defined names={missing_names}")
        missing_tables = sorted(set(expected.get("tableNames", [])) - table_names)
        if missing_tables:
            failures.append(f"missing tables={missing_tables}")
        if len(report.get("chartParts", [])) != expected.get("chartPartCount"):
            failures.append(f"chartParts={report.get('chartParts')}")
        if len(report.get("drawingParts", [])) != expected.get("drawingPartCount"):
            failures.append(f"drawingParts={report.get('drawingParts')}")
        if report.get("hasVbaProject") is not False:
            failures.append(f"hasVbaProject={report.get('hasVbaProject')}")
        if report.get("connections"):
            failures.append(f"connections={report.get('connections')}")
        if report.get("externalLinks"):
            failures.append(f"externalLinks={report.get('externalLinks')}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="sheets, formulas, names, table, and chart/drawing parts verified"
            if not failures
            else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "sheets": sheet_names,
                "formulaCount": report.get("totalFormulaCount"),
                "formulaFunctions": sorted(formula_functions),
                "definedNames": sorted(str(item) for item in defined_names),
                "tables": sorted(str(item) for item in table_names),
                "chartPartCount": len(report.get("chartParts", [])),
                "drawingPartCount": len(report.get("drawingParts", [])),
            },
        )


def visual_qa_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Visual QA report fixture smoke"
    create_script = project_root / "tools" / "create_visual_qa_fixture.py"
    report_script = project_root / "tools" / "build_visual_qa_report.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    missing = [str(path) for path in [create_script, report_script, inspect_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_visual_qa_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "visual_qa_fixture.xlsx"
        fixture_json = tmp_dir / "fixture.json"
        report_json = tmp_dir / "visual_qa.json"
        report_md = tmp_dir / "visual_qa.md"
        openxml_json = tmp_dir / "openxml.json"
        steps = [
            [
                sys.executable,
                str(create_script),
                "--workbook",
                str(workbook),
                "--out-json",
                str(fixture_json),
            ],
            [
                sys.executable,
                str(report_script),
                "--workbook",
                str(workbook),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(workbook),
                "--out-json",
                str(openxml_json),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            fixture = json.loads(fixture_json.read_text(encoding="utf-8"))
            report = json.loads(report_json.read_text(encoding="utf-8"))
            openxml = json.loads(openxml_json.read_text(encoding="utf-8"))
            markdown = report_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read visual QA reports: {exc}")

        expected = fixture.get("expected", {}) if isinstance(fixture.get("expected"), dict) else {}
        summary = report.get("summary", {}) if isinstance(report.get("summary"), dict) else {}
        finding_codes = {item.get("code") for item in report.get("findings", []) if isinstance(item, dict)}
        expected_codes = set(expected.get("expectedFindingCodes", []))
        failures: list[str] = []
        if summary.get("readiness") != expected.get("expectedReadiness"):
            failures.append(f"readiness={summary.get('readiness')}")
        if summary.get("reportSheetCount") != 3:
            failures.append(f"reportSheetCount={summary.get('reportSheetCount')}")
        if summary.get("blankReportSheetCount") != 1:
            failures.append(f"blankReportSheetCount={summary.get('blankReportSheetCount')}")
        if summary.get("reportSheetWithChartCount") != 1:
            failures.append(f"reportSheetWithChartCount={summary.get('reportSheetWithChartCount')}")
        missing_codes = sorted(expected_codes - finding_codes)
        if missing_codes:
            failures.append(f"missingFindingCodes={missing_codes}")
        if len(openxml.get("chartParts", [])) != 1:
            failures.append(f"chartParts={openxml.get('chartParts')}")
        if "# Visual QA Report" not in markdown or "blocked-for-delivery" not in markdown:
            failures.append("Markdown heading or readiness missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="workbook-backed visual QA fixture detected clean, clipped, and blank report surfaces"
            if not failures
            else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "readiness": summary.get("readiness"),
                "reportSheetCount": summary.get("reportSheetCount"),
                "findingCodes": sorted(str(code) for code in finding_codes),
                "chartPartCount": len(openxml.get("chartParts", [])),
            },
        )


def visual_qa_render_evidence_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "Visual QA render evidence smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")

    script = project_root / "tools" / "export_visual_qa_render_evidence.ps1"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    unavailable_markers = [
        "cannot create activex component",
        "class not registered",
        "excel.application",
        "retrieving the com class factory",
    ]

    with tempfile.TemporaryDirectory(prefix="excel_bi_visual_render_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "visual_render_fixture.xlsx"
        pdf_dir = tmp_dir / "pdf"
        report_json = tmp_dir / "render.json"
        report_md = tmp_dir / "render.md"
        command = [
            ps_exe,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-WorkbookPath",
            str(workbook),
            "-CreateFixture",
            "-OutDir",
            str(pdf_dir),
            "-OutJson",
            str(report_json),
            "-OutMd",
            str(report_md),
        ]
        result = run_command(command, project_root, name)
        if result.status != PASS:
            combined = f"{result.stdout}\n{result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                result.status = SKIP
                result.detail = "Excel COM runtime unavailable for PDF render evidence"
            return result

        try:
            report = json.loads(report_json.read_text(encoding="utf-8-sig"))
            markdown = report_md.read_text(encoding="utf-8-sig")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read render evidence report: {exc}")

        summary = report.get("summary", {}) if isinstance(report.get("summary"), dict) else {}
        workbook_info = report.get("workbook", {}) if isinstance(report.get("workbook"), dict) else {}
        exports = report.get("exports", []) if isinstance(report.get("exports"), list) else []
        sheet_names = {item.get("sheetName") for item in exports if isinstance(item, dict)}
        failures: list[str] = []
        if summary.get("readiness") != "rendered":
            failures.append(f"readiness={summary.get('readiness')}")
        if workbook_info.get("createdFixture") is not True:
            failures.append(f"createdFixture={workbook_info.get('createdFixture')}")
        if summary.get("exportedSheetCount") != 3 or summary.get("exportedFileCount") != 3:
            failures.append(f"export count={summary.get('exportedSheetCount')}/{summary.get('exportedFileCount')}")
        if int(summary.get("totalBytes", 0) or 0) <= 0:
            failures.append(f"totalBytes={summary.get('totalBytes')}")
        for required in ["Report_OK", "Report_Clipped", "Report_Blank"]:
            if required not in sheet_names:
                failures.append(f"missing export sheet={required}")
        for item in exports:
            if not isinstance(item, dict):
                failures.append("non-dict export item")
                continue
            path = Path(str(item.get("path", "")))
            if not path.is_file():
                failures.append(f"missing pdf={path}")
            if int(item.get("bytes", 0) or 0) <= 0:
                failures.append(f"empty pdf={item.get('fileName')}")
        if "# Visual QA Render Evidence" not in markdown or "Windows Excel COM PDF export" not in markdown:
            failures.append("Markdown heading or source missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="Excel COM exported sanitized Report* sheets to PDF and recorded render evidence"
            if not failures
            else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "readiness": summary.get("readiness"),
                "exportedSheetCount": summary.get("exportedSheetCount"),
                "exportedFileCount": summary.get("exportedFileCount"),
                "totalBytes": summary.get("totalBytes"),
                "sheetNames": sorted(str(name) for name in sheet_names),
            },
        )


def formula_quality_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Formula quality report fixture smoke"
    fixture_script = project_root / "tools" / "create_formula_quality_fixture.py"
    report_script = project_root / "tools" / "build_formula_quality_report.py"
    missing = [str(path) for path in [fixture_script, report_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_formula_quality_") as tmp:
        tmp_dir = Path(tmp)
        fixture_dir = tmp_dir / "fixture"
        manifest_json = tmp_dir / "fixture_manifest.json"
        safe_report_json = tmp_dir / "safe_formula_quality.json"
        risky_report_json = tmp_dir / "risky_formula_quality.json"
        risky_report_md = tmp_dir / "risky_formula_quality.md"

        fixture_result = run_command(
            [
                sys.executable,
                str(fixture_script),
                "--out-dir",
                str(fixture_dir),
                "--out-json",
                str(manifest_json),
            ],
            project_root,
            name,
        )
        if fixture_result.status != PASS:
            return fixture_result

        try:
            manifest = json.loads(manifest_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read formula quality fixture manifest: {exc}")

        expected = manifest.get("expected", {})
        if not isinstance(expected, dict):
            expected = {}
        safe_expected = expected.get("safe", {})
        risky_expected = expected.get("risky", {})
        if not isinstance(safe_expected, dict):
            safe_expected = {}
        if not isinstance(risky_expected, dict):
            risky_expected = {}

        safe_json = Path(str(manifest.get("safeOpenXmlJson", "")))
        risky_json = Path(str(manifest.get("riskyOpenXmlJson", "")))
        safe_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--openxml-json",
                str(safe_json),
                "--out-json",
                str(safe_report_json),
                "--fail-on-high-risk",
            ],
            project_root,
            name,
        )
        if safe_result.status != PASS:
            return safe_result

        risky_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--openxml-json",
                str(risky_json),
                "--out-json",
                str(risky_report_json),
                "--out-md",
                str(risky_report_md),
            ],
            project_root,
            name,
        )
        if risky_result.status != PASS:
            return risky_result

        strict_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--openxml-json",
                str(risky_json),
                "--fail-on-high-risk",
            ],
            project_root,
            name,
            ok_codes={1},
        )
        if strict_result.status != PASS:
            return strict_result

        try:
            safe_report = json.loads(safe_report_json.read_text(encoding="utf-8"))
            risky_report = json.loads(risky_report_json.read_text(encoding="utf-8"))
            risky_markdown = risky_report_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read formula quality reports: {exc}")

        safe_summary = safe_report.get("summary", {})
        risky_summary = risky_report.get("summary", {})
        risky_codes = {str(item.get("code", "")) for item in risky_report.get("findings", []) if isinstance(item, dict)}
        failures: list[str] = []

        for field in ["readiness", "formulaCount", "findingCount"]:
            if safe_summary.get(field) != safe_expected.get(field):
                failures.append(f"safe expected {field}={safe_expected.get(field)}, got {safe_summary.get(field)}")
        for field in ["readiness", "formulaCount", "highFindingCount", "mediumFindingCount", "lowFindingCount"]:
            if risky_summary.get(field) != risky_expected.get(field):
                failures.append(f"risky expected {field}={risky_expected.get(field)}, got {risky_summary.get(field)}")
        missing_codes = sorted(set(risky_expected.get("requiredCodes", [])) - risky_codes)
        if missing_codes:
            failures.append(f"missing finding codes={missing_codes}")
        if "Formula Quality Report" not in risky_markdown:
            failures.append("markdown report heading missing")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))

        return CheckResult(
            name=name,
            status=PASS,
            detail="clean and high-risk formula quality reports verified",
            metadata={
                "safeFormulaCount": safe_summary.get("formulaCount"),
                "riskyFormulaCount": risky_summary.get("formulaCount"),
                "riskyHighFindingCount": risky_summary.get("highFindingCount"),
                "strictFailurePathVerified": True,
            },
        )


def workbook_controls_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Workbook controls report fixture smoke"
    create_script = project_root / "tools" / "create_workbook_controls_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    report_script = project_root / "tools" / "build_workbook_controls_report.py"
    missing = [str(path) for path in [create_script, inspect_script, report_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_workbook_controls_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "workbook_controls_fixture.xlsx"
        fixture_json = tmp_dir / "fixture.json"
        openxml_json = tmp_dir / "openxml.json"
        report_json = tmp_dir / "controls.json"
        report_md = tmp_dir / "controls.md"

        steps = [
            [
                sys.executable,
                str(create_script),
                "--workbook",
                str(workbook),
                "--out-json",
                str(fixture_json),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(workbook),
                "--out-json",
                str(openxml_json),
            ],
            [
                sys.executable,
                str(report_script),
                "--openxml-json",
                str(openxml_json),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        strict_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--openxml-json",
                str(openxml_json),
                "--fail-on-review",
            ],
            project_root,
            name,
            ok_codes={1},
        )
        if strict_result.status != PASS:
            return strict_result

        try:
            fixture = json.loads(fixture_json.read_text(encoding="utf-8"))
            report = json.loads(report_json.read_text(encoding="utf-8"))
            markdown = report_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read workbook controls reports: {exc}")

        expected = fixture.get("expected", {})
        if not isinstance(expected, dict):
            expected = {}
        summary = report.get("summary", {})
        codes = {str(item.get("code", "")) for item in report.get("findings", []) if isinstance(item, dict)}
        failures: list[str] = []
        for field in [
            "readiness",
            "sheetCount",
            "hiddenSheetCount",
            "veryHiddenSheetCount",
            "protectedSheetCount",
            "filteredSheetCount",
            "frozenPaneSheetCount",
            "dataValidationSheetCount",
            "hasWorkbookProtection",
            "mediumFindingCount",
            "lowFindingCount",
        ]:
            if summary.get(field) != expected.get(field):
                failures.append(f"expected {field}={expected.get(field)}, got {summary.get(field)}")
        missing_codes = sorted(set(expected.get("requiredCodes", [])) - codes)
        if missing_codes:
            failures.append(f"missing finding codes={missing_codes}")
        if "Workbook Controls Report" not in markdown:
            failures.append("markdown report heading missing")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))

        return CheckResult(
            name=name,
            status=PASS,
            detail="workbook protection, sheet visibility, filters, frozen panes, and validation controls verified",
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "sheetCount": summary.get("sheetCount"),
                "hiddenSheetCount": summary.get("hiddenSheetCount"),
                "veryHiddenSheetCount": summary.get("veryHiddenSheetCount"),
                "protectedSheetCount": summary.get("protectedSheetCount"),
                "strictFailurePathVerified": True,
            },
        )


def external_dependency_report_fixture_check(project_root: Path) -> CheckResult:
    name = "External dependency readiness report fixture smoke"
    create_script = project_root / "tools" / "create_external_dependency_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    report_script = project_root / "tools" / "build_external_dependency_report.py"
    for script in [create_script, inspect_script, report_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_external_readiness_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "external_dependency_fixture.xlsx"
        openxml_json = tmp_dir / "openxml.json"
        report_json = tmp_dir / "readiness.json"
        report_md = tmp_dir / "readiness.md"

        steps = [
            [
                sys.executable,
                str(create_script),
                "--workbook",
                str(workbook),
            ],
            [
                sys.executable,
                str(inspect_script),
                str(workbook),
                "--out-json",
                str(openxml_json),
            ],
            [
                sys.executable,
                str(report_script),
                "--openxml-json",
                str(openxml_json),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            report = json.loads(report_json.read_text(encoding="utf-8"))
            markdown = report_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read readiness report: {exc}")

        summary = report.get("summary", {})
        codes = {str(item.get("code", "")) for item in report.get("findings", []) if isinstance(item, dict)}
        expected_codes = {
            "workbook-connections",
            "connection-credential-like-literal",
            "external-link-parts",
            "external-formula-references",
            "external-defined-names",
            "mashup-like-parts",
        }
        failures: list[str] = []
        if summary.get("readiness") != "blocked-for-pure-deliverable":
            failures.append(f"readiness={summary.get('readiness')}")
        if summary.get("maxSeverity") != "high":
            failures.append(f"maxSeverity={summary.get('maxSeverity')}")
        if summary.get("credentialLikeConnectionCount") != 1:
            failures.append(f"credentialLikeConnectionCount={summary.get('credentialLikeConnectionCount')}")
        missing = expected_codes - codes
        if missing:
            failures.append(f"missing finding codes={sorted(missing)}")
        if "External Dependency Readiness Report" not in markdown:
            failures.append("Markdown report title missing")
        if "blocked-for-pure-deliverable" not in markdown:
            failures.append("Markdown report readiness missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="high-risk external dependency findings, redacted credential-like connection indicators, and Markdown report verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "readiness": summary.get("readiness"),
                "maxSeverity": summary.get("maxSeverity"),
                "findingCodes": sorted(codes),
            },
        )


def workbook_triage_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Workbook triage report fixture smoke"
    scripts = {
        "surface": project_root / "tools" / "create_workbook_surface_fixture.py",
        "external_fixture": project_root / "tools" / "create_external_dependency_fixture.py",
        "cube_fixture": project_root / "tools" / "create_cube_formula_fixture.py",
        "inspect": project_root / "tools" / "inspect_excel_bi_workbook.py",
        "formula": project_root / "tools" / "build_formula_quality_report.py",
        "controls": project_root / "tools" / "build_workbook_controls_report.py",
        "external": project_root / "tools" / "build_external_dependency_report.py",
        "cube": project_root / "tools" / "build_cube_dependency_report.py",
        "triage": project_root / "tools" / "build_workbook_triage_report.py",
    }
    missing = [str(path) for path in scripts.values() if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_workbook_triage_") as tmp:
        tmp_dir = Path(tmp)
        stdout: list[str] = []
        stderr: list[str] = []

        def run_step(command: list[str], label: str, ok_codes: set[int] | None = None) -> CheckResult:
            result = run_command(command, project_root, f"{name}: {label}", ok_codes=ok_codes)
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            return result

        surface_workbook = tmp_dir / "surface.xlsx"
        surface_openxml = tmp_dir / "surface_openxml.json"
        surface_formula = tmp_dir / "surface_formula.json"
        surface_controls = tmp_dir / "surface_controls.json"
        surface_triage = tmp_dir / "surface_triage.json"
        surface_triage_md = tmp_dir / "surface_triage.md"
        surface_steps = [
            [sys.executable, str(scripts["surface"]), "--workbook", str(surface_workbook)],
            [sys.executable, str(scripts["inspect"]), str(surface_workbook), "--out-json", str(surface_openxml)],
            [sys.executable, str(scripts["formula"]), "--openxml-json", str(surface_openxml), "--out-json", str(surface_formula)],
            [sys.executable, str(scripts["controls"]), "--openxml-json", str(surface_openxml), "--out-json", str(surface_controls)],
            [
                sys.executable,
                str(scripts["triage"]),
                "--inspection-json",
                str(surface_openxml),
                "--formula-report-json",
                str(surface_formula),
                "--controls-report-json",
                str(surface_controls),
                "--out-json",
                str(surface_triage),
                "--out-md",
                str(surface_triage_md),
                "--require-pass",
            ],
        ]
        for index, command in enumerate(surface_steps, start=1):
            result = run_step(command, f"surface step {index}")
            if result.status != PASS:
                result.name = name
                return result

        external_workbook = tmp_dir / "external.xlsx"
        external_openxml = tmp_dir / "external_openxml.json"
        external_formula = tmp_dir / "external_formula.json"
        external_controls = tmp_dir / "external_controls.json"
        external_report = tmp_dir / "external_report.json"
        external_triage = tmp_dir / "external_triage.json"
        external_triage_md = tmp_dir / "external_triage.md"
        external_steps = [
            [sys.executable, str(scripts["external_fixture"]), "--workbook", str(external_workbook)],
            [sys.executable, str(scripts["inspect"]), str(external_workbook), "--out-json", str(external_openxml)],
            [sys.executable, str(scripts["formula"]), "--openxml-json", str(external_openxml), "--out-json", str(external_formula)],
            [sys.executable, str(scripts["controls"]), "--openxml-json", str(external_openxml), "--out-json", str(external_controls)],
            [sys.executable, str(scripts["external"]), "--openxml-json", str(external_openxml), "--out-json", str(external_report)],
            [
                sys.executable,
                str(scripts["triage"]),
                "--inspection-json",
                str(external_openxml),
                "--formula-report-json",
                str(external_formula),
                "--controls-report-json",
                str(external_controls),
                "--external-report-json",
                str(external_report),
                "--out-json",
                str(external_triage),
                "--out-md",
                str(external_triage_md),
            ],
        ]
        for index, command in enumerate(external_steps, start=1):
            result = run_step(command, f"external step {index}")
            if result.status != PASS:
                result.name = name
                return result

        strict_external = run_step(
            [
                sys.executable,
                str(scripts["triage"]),
                "--inspection-json",
                str(external_openxml),
                "--formula-report-json",
                str(external_formula),
                "--controls-report-json",
                str(external_controls),
                "--external-report-json",
                str(external_report),
                "--require-pass",
            ],
            "external require-pass",
            ok_codes={1},
        )
        if strict_external.status != PASS:
            strict_external.name = name
            return strict_external

        cube_workbook = tmp_dir / "cube.xlsx"
        cube_model = tmp_dir / "cube_model.json"
        cube_openxml = tmp_dir / "cube_openxml.json"
        cube_report = tmp_dir / "cube_report.json"
        cube_triage = tmp_dir / "cube_triage.json"
        cube_triage_md = tmp_dir / "cube_triage.md"
        cube_steps = [
            [sys.executable, str(scripts["cube_fixture"]), "--workbook", str(cube_workbook), "--model-json", str(cube_model)],
            [sys.executable, str(scripts["inspect"]), str(cube_workbook), "--out-json", str(cube_openxml)],
            [sys.executable, str(scripts["cube"]), "--openxml-json", str(cube_openxml), "--model-json", str(cube_model), "--out-json", str(cube_report)],
            [
                sys.executable,
                str(scripts["triage"]),
                "--inspection-json",
                str(cube_openxml),
                "--cube-report-json",
                str(cube_report),
                "--out-json",
                str(cube_triage),
                "--out-md",
                str(cube_triage_md),
            ],
        ]
        for index, command in enumerate(cube_steps, start=1):
            result = run_step(command, f"cube step {index}")
            if result.status != PASS:
                result.name = name
                return result

        try:
            surface = json.loads(surface_triage.read_text(encoding="utf-8"))
            external = json.loads(external_triage.read_text(encoding="utf-8"))
            cube = json.loads(cube_triage.read_text(encoding="utf-8"))
            external_md = external_triage_md.read_text(encoding="utf-8")
            cube_md = cube_triage_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read triage reports: {exc}")

        failures: list[str] = []
        if surface.get("status") != "pass":
            failures.append(f"surface status={surface.get('status')}")
        if external.get("status") != "blocked":
            failures.append(f"external status={external.get('status')}")
        if cube.get("status") != "blocked":
            failures.append(f"cube status={cube.get('status')}")
        external_gaps = {str(item.get("kind", "")) for item in external.get("coverageGaps", []) if isinstance(item, dict)}
        if "powerQueryLineage" not in external_gaps:
            failures.append(f"external missing powerQueryLineage gap, gaps={sorted(external_gaps)}")
        cube_gaps = {str(item.get("kind", "")) for item in cube.get("coverageGaps", []) if isinstance(item, dict)}
        if "model" not in cube_gaps:
            failures.append(f"cube missing model gap, gaps={sorted(cube_gaps)}")
        if "Workbook Triage Report" not in external_md or "## Boundary" not in external_md:
            failures.append("external markdown heading or boundary missing")
        if "Data Model/CUBE" not in cube_md:
            failures.append("cube markdown missing Data Model/CUBE action")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="pass, external-blocked, and cube/model triage paths verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "surfaceStatus": surface.get("status"),
                "externalStatus": external.get("status"),
                "externalGaps": sorted(external_gaps),
                "cubeStatus": cube.get("status"),
                "cubeGaps": sorted(cube_gaps),
                "strictFailurePathVerified": True,
            },
        )


def pure_deliverable_cleanup_plan_fixture_check(project_root: Path) -> CheckResult:
    name = "Pure deliverable cleanup plan fixture smoke"
    create_script = project_root / "tools" / "create_external_dependency_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    readiness_script = project_root / "tools" / "build_external_dependency_report.py"
    cleanup_script = project_root / "tools" / "build_pure_deliverable_cleanup_plan.py"
    for script in [create_script, inspect_script, readiness_script, cleanup_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_pure_cleanup_plan_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "external_dependency_fixture.xlsx"
        openxml_json = tmp_dir / "openxml.json"
        readiness_json = tmp_dir / "readiness.json"
        plan_json = tmp_dir / "cleanup_plan.json"
        plan_md = tmp_dir / "cleanup_plan.md"

        steps = [
            [sys.executable, str(create_script), "--workbook", str(workbook)],
            [sys.executable, str(inspect_script), str(workbook), "--out-json", str(openxml_json)],
            [
                sys.executable,
                str(readiness_script),
                "--openxml-json",
                str(openxml_json),
                "--out-json",
                str(readiness_json),
            ],
            [
                sys.executable,
                str(cleanup_script),
                "--readiness-json",
                str(readiness_json),
                "--target",
                "pure-xlsx",
                "--out-json",
                str(plan_json),
                "--out-md",
                str(plan_md),
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            plan = json.loads(plan_json.read_text(encoding="utf-8"))
            markdown = plan_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read cleanup plan: {exc}")

        actions = {str(item.get("action", "")) for item in plan.get("steps", []) if isinstance(item, dict)}
        expected_actions = {
            "create-working-copy",
            "refresh-and-freeze-values",
            "replace-external-formulas",
            "remove-external-defined-names",
            "remove-external-links",
            "remove-workbook-connections",
            "remove-power-query-mashup-parts",
            "post-clean-openxml-audit",
        }
        assertion_checks = {
            str(item.get("check", ""))
            for item in plan.get("postCleanupAssertions", [])
            if isinstance(item, dict)
        }
        expected_assertions = {
            "connectionCount",
            "externalLinkPartCount",
            "externalFormulaCount",
            "externalDefinedNameCount",
            "hasMashupLikeParts",
            "hasVbaProject",
        }
        failures: list[str] = []
        if plan.get("status") != "cleanup-required":
            failures.append(f"status={plan.get('status')}")
        missing_actions = expected_actions - actions
        if missing_actions:
            failures.append(f"missing actions={sorted(missing_actions)}")
        missing_assertions = expected_assertions - assertion_checks
        if missing_assertions:
            failures.append(f"missing assertions={sorted(missing_assertions)}")
        if "Pure Deliverable Cleanup Plan" not in markdown:
            failures.append("Markdown report title missing")
        if "cleanup-required" not in markdown:
            failures.append("Markdown status missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="pure-xlsx cleanup actions and post-clean assertions verified" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "status": plan.get("status"),
                "stepCount": plan.get("stepCount"),
                "actions": sorted(actions),
                "assertions": sorted(assertion_checks),
            },
        )


def pure_deliverable_verification_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Pure deliverable verification report fixture smoke"
    create_dirty_script = project_root / "tools" / "create_external_dependency_fixture.py"
    create_clean_script = project_root / "tools" / "create_pure_deliverable_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    readiness_script = project_root / "tools" / "build_external_dependency_report.py"
    cleanup_script = project_root / "tools" / "build_pure_deliverable_cleanup_plan.py"
    verify_script = project_root / "tools" / "build_pure_deliverable_verification_report.py"
    for script in [create_dirty_script, create_clean_script, inspect_script, readiness_script, cleanup_script, verify_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_pure_verify_") as tmp:
        tmp_dir = Path(tmp)
        dirty_workbook = tmp_dir / "external_dependency_fixture.xlsx"
        dirty_openxml_json = tmp_dir / "dirty_openxml.json"
        dirty_readiness_json = tmp_dir / "dirty_readiness.json"
        plan_json = tmp_dir / "cleanup_plan.json"
        clean_workbook = tmp_dir / "pure_deliverable_fixture.xlsx"
        clean_openxml_json = tmp_dir / "clean_openxml.json"
        clean_readiness_json = tmp_dir / "clean_readiness.json"
        verify_json = tmp_dir / "verification.json"
        verify_md = tmp_dir / "verification.md"

        steps = [
            [sys.executable, str(create_dirty_script), "--workbook", str(dirty_workbook)],
            [sys.executable, str(inspect_script), str(dirty_workbook), "--out-json", str(dirty_openxml_json)],
            [
                sys.executable,
                str(readiness_script),
                "--openxml-json",
                str(dirty_openxml_json),
                "--out-json",
                str(dirty_readiness_json),
            ],
            [
                sys.executable,
                str(cleanup_script),
                "--readiness-json",
                str(dirty_readiness_json),
                "--target",
                "pure-xlsx",
                "--out-json",
                str(plan_json),
            ],
            [sys.executable, str(create_clean_script), "--workbook", str(clean_workbook)],
            [sys.executable, str(inspect_script), str(clean_workbook), "--out-json", str(clean_openxml_json)],
            [
                sys.executable,
                str(readiness_script),
                "--openxml-json",
                str(clean_openxml_json),
                "--out-json",
                str(clean_readiness_json),
            ],
            [
                sys.executable,
                str(verify_script),
                "--cleanup-plan-json",
                str(plan_json),
                "--post-readiness-json",
                str(clean_readiness_json),
                "--out-json",
                str(verify_json),
                "--out-md",
                str(verify_md),
                "--fail-on-fail",
            ],
        ]
        stdout: list[str] = []
        stderr: list[str] = []
        for index, command in enumerate(steps, start=1):
            result = run_command(command, project_root, f"{name}: step {index}")
            stdout.append(result.stdout)
            stderr.append(result.stderr)
            if result.status != PASS:
                return CheckResult(
                    name=name,
                    status=FAIL,
                    detail=f"step {index} failed: {result.detail}",
                    command=command,
                    stdout="\n".join(part for part in stdout if part),
                    stderr="\n".join(part for part in stderr if part),
                )

        try:
            report = json.loads(verify_json.read_text(encoding="utf-8"))
            markdown = verify_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read verification report: {exc}")

        assertion_statuses = {
            str(item.get("check", "")): str(item.get("status", ""))
            for item in report.get("assertions", [])
            if isinstance(item, dict)
        }
        expected_passed_assertions = {
            "connectionCount",
            "externalLinkPartCount",
            "externalFormulaCount",
            "externalDefinedNameCount",
            "hasMashupLikeParts",
            "hasPowerPivotLikeParts",
            "cubeFormulaCount",
            "hasVbaProject",
        }
        failures: list[str] = []
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        for check in expected_passed_assertions:
            if assertion_statuses.get(check) != PASS:
                failures.append(f"{check}={assertion_statuses.get(check)}")
        if report.get("failedCount") != 0:
            failures.append(f"failedCount={report.get('failedCount')}")
        if report.get("manualReviewCount") != 0:
            failures.append(f"manualReviewCount={report.get('manualReviewCount')}")
        if "Pure Deliverable Verification Report" not in markdown:
            failures.append("Markdown report title missing")
        if "status: **pass**" not in markdown:
            failures.append("Markdown status missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="post-clean pure-xlsx assertions verified against clean readiness report" if not failures else "; ".join(failures),
            stdout="\n".join(part for part in stdout if part),
            stderr="\n".join(part for part in stderr if part),
            metadata={
                "status": report.get("status"),
                "assertionCount": report.get("assertionCount"),
                "passedCount": report.get("passedCount"),
                "assertions": assertion_statuses,
            },
        )


def goal_coverage_report_check(project_root: Path) -> CheckResult:
    name = "Goal coverage report"
    coverage_script = project_root / "tools" / "build_goal_coverage_report.py"
    if not coverage_script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {coverage_script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_goal_coverage_") as tmp:
        tmp_dir = Path(tmp)
        report_json = tmp_dir / "goal_coverage.json"
        report_md = tmp_dir / "goal_coverage.md"
        command = [
            sys.executable,
            str(coverage_script),
            "--project-root",
            str(project_root),
            "--out-json",
            str(report_json),
            "--out-md",
            str(report_md),
            "--require-pass",
        ]
        result = run_command(command, project_root, name)
        if result.status != PASS:
            return result

        try:
            report = json.loads(report_json.read_text(encoding="utf-8"))
            markdown = report_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read coverage report: {exc}", command=command)

        failures: list[str] = []
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        if int(report.get("areaCount") or 0) < 12:
            failures.append(f"areaCount={report.get('areaCount')}")
        if report.get("failedAreaCount") != 0:
            failures.append(f"failedAreaCount={report.get('failedAreaCount')}")
        if "Goal Coverage Report" not in markdown:
            failures.append("Markdown report title missing")
        if "status: **pass**" not in markdown:
            failures.append("Markdown status missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="public goal areas covered by shipped files and public documentation" if not failures else "; ".join(failures),
            command=command,
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "status": report.get("status"),
                "areaCount": report.get("areaCount"),
                "passedAreaCount": report.get("passedAreaCount"),
                "failedAreaCount": report.get("failedAreaCount"),
            },
        )


def measure_rename_impact_fixture_check(project_root: Path) -> CheckResult:
    name = "Measure rename impact fixture smoke"
    create_script = project_root / "tools" / "create_cube_formula_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    impact_script = project_root / "tools" / "analyze_measure_rename_impact.py"
    for script in [create_script, inspect_script, impact_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_measure_impact_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "cube_formula_fixture.xlsx"
        model_json = tmp_dir / "model.json"
        openxml_json = tmp_dir / "openxml.json"
        impact_json = tmp_dir / "impact.json"

        create_result = run_command([sys.executable, str(create_script), "--workbook", str(workbook), "--model-json", str(model_json)], project_root, f"{name}: create fixture")
        if create_result.status != PASS:
            create_result.name = name
            return create_result
        inspect_result = run_command([sys.executable, str(inspect_script), str(workbook), "--out-json", str(openxml_json)], project_root, f"{name}: inspect fixture")
        if inspect_result.status != PASS:
            inspect_result.name = name
            return inspect_result

        try:
            model = json.loads(model_json.read_text(encoding="utf-8"))
            measures = list(model.get("measures", []))
            measures.append(
                {
                    "name": "Profit",
                    "associatedTable": {"name": "Fact"},
                    "formula": "[Revenue] - SUM(Fact[Cost])",
                }
            )
            model["measures"] = measures
            model["measureCount"] = len(measures)
            model_json.write_text(json.dumps(model, indent=2), encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not prepare model fixture: {exc}")

        impact_result = run_command(
            [
                sys.executable,
                str(impact_script),
                "--model-json",
                str(model_json),
                "--openxml-json",
                str(openxml_json),
                "--rename",
                "Revenue=Net Revenue",
                "--out-json",
                str(impact_json),
                "--fail-on-high-risk",
            ],
            project_root,
            f"{name}: analyze rename",
            ok_codes={1},
        )
        if impact_result.status != PASS:
            impact_result.name = name
            impact_result.status = FAIL
            impact_result.detail = "expected Revenue rename fixture to be high-risk"
            return impact_result

        try:
            report = json.loads(impact_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read impact report: {exc}")

        changes = report.get("changes", [])
        if len(changes) != 1:
            return CheckResult(name=name, status=FAIL, detail=f"expected 1 change, got {len(changes)}")
        change = changes[0]
        failures: list[str] = []
        if change.get("riskLevel") != "high":
            failures.append(f"riskLevel={change.get('riskLevel')}")
        if change.get("daxFormulaHitCount") != 1:
            failures.append(f"daxFormulaHitCount={change.get('daxFormulaHitCount')}")
        if change.get("cubeFormulaHitCount") != 3:
            failures.append(f"cubeFormulaHitCount={change.get('cubeFormulaHitCount')}")
        if "Report" not in change.get("affectedSheets", []):
            failures.append("Report sheet not in affectedSheets")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))
        return CheckResult(
            name=name,
            status=PASS,
            detail="Revenue rename impact detected across DAX and CUBE formulas",
            metadata={
                "riskLevel": change.get("riskLevel"),
                "daxFormulaHitCount": change.get("daxFormulaHitCount"),
                "cubeFormulaHitCount": change.get("cubeFormulaHitCount"),
                "affectedSheets": change.get("affectedSheets"),
            },
        )


def measure_rename_rewrite_plan_fixture_check(project_root: Path) -> CheckResult:
    name = "Measure rename rewrite plan fixture smoke"
    create_script = project_root / "tools" / "create_cube_formula_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    plan_script = project_root / "tools" / "build_measure_rename_rewrite_plan.py"
    for script in [create_script, inspect_script, plan_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_measure_rewrite_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "cube_formula_fixture.xlsx"
        model_json = tmp_dir / "model.json"
        openxml_json = tmp_dir / "openxml.json"
        plan_json = tmp_dir / "rewrite_plan.json"

        create_result = run_command([sys.executable, str(create_script), "--workbook", str(workbook), "--model-json", str(model_json)], project_root, f"{name}: create fixture")
        if create_result.status != PASS:
            create_result.name = name
            return create_result
        inspect_result = run_command([sys.executable, str(inspect_script), str(workbook), "--out-json", str(openxml_json)], project_root, f"{name}: inspect fixture")
        if inspect_result.status != PASS:
            inspect_result.name = name
            return inspect_result

        try:
            model = json.loads(model_json.read_text(encoding="utf-8"))
            measures = list(model.get("measures", []))
            measures.append(
                {
                    "name": "Profit",
                    "associatedTable": {"name": "Fact"},
                    "formula": "[Revenue] - SUM(Fact[Revenue])",
                }
            )
            model["measures"] = measures
            model["measureCount"] = len(measures)
            model_json.write_text(json.dumps(model, indent=2), encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not prepare model fixture: {exc}")

        try:
            openxml = json.loads(openxml_json.read_text(encoding="utf-8"))
            formulas = list(openxml.get("cubeFormulas", []))
            formulas.extend(
                [
                    {
                        "sheet": "Report",
                        "cell": "H2",
                        "formula": 'CUBEMEMBER("ThisWorkbookDataModel","[Measures].[Revenue]","Revenue helper")',
                        "formulaType": "bi",
                        "cachedValue": "Revenue",
                    },
                    {
                        "sheet": "Report",
                        "cell": "I2",
                        "formula": 'CUBEVALUE("ThisWorkbookDataModel",$H$2,$A$2)',
                        "formulaType": "bi",
                        "cachedValue": "120",
                    },
                ]
            )
            openxml["cubeFormulas"] = formulas
            openxml["cubeFormulaCount"] = len(formulas)
            openxml_json.write_text(json.dumps(openxml, indent=2), encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not prepare OpenXML edge fixture: {exc}")

        plan_result = run_command(
            [
                sys.executable,
                str(plan_script),
                "--model-json",
                str(model_json),
                "--openxml-json",
                str(openxml_json),
                "--rename",
                "Revenue=Net Revenue",
                "--out-json",
                str(plan_json),
            ],
            project_root,
            f"{name}: build plan",
        )
        if plan_result.status != PASS:
            plan_result.name = name
            return plan_result

        try:
            report = json.loads(plan_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read rewrite plan: {exc}")

        changes = report.get("changes", [])
        failures: list[str] = []
        if len(changes) != 1:
            failures.append(f"expected 1 change, got {len(changes)}")
        change = changes[0] if changes else {}
        if report.get("daxRewriteCount") != 1:
            failures.append(f"daxRewriteCount={report.get('daxRewriteCount')}")
        if report.get("cubeRewriteCount") != 4:
            failures.append(f"cubeRewriteCount={report.get('cubeRewriteCount')}")
        if report.get("downstreamFormulaImpactCount") != 1:
            failures.append(f"downstreamFormulaImpactCount={report.get('downstreamFormulaImpactCount')}")
        if report.get("manualReviewCount") != 0:
            failures.append(f"manualReviewCount={report.get('manualReviewCount')}")
        dax_rewrites = change.get("daxRewrites", [])
        if not dax_rewrites:
            failures.append("missing DAX rewrite")
        else:
            new_formula = dax_rewrites[0].get("newFormula", "")
            if "[Net Revenue]" not in new_formula:
                failures.append("DAX rewrite did not include [Net Revenue]")
            if "Fact[Revenue]" not in new_formula:
                failures.append("DAX table-column reference was changed unexpectedly")
        cube_rewrites = change.get("cubeRewrites", [])
        if not cube_rewrites or not all("[Measures].[Net Revenue]" in item.get("newFormula", "") for item in cube_rewrites):
            failures.append("CUBE rewrites did not all include [Measures].[Net Revenue]")
        dynamic_rewrites = [item for item in cube_rewrites if item.get("cell") == "G2"]
        if not dynamic_rewrites or "$A$5" not in dynamic_rewrites[0].get("helperCellReferences", []):
            failures.append("dynamic period helper $A$5 was not retained on G2 rewrite")
        helper_rewrites = [item for item in cube_rewrites if item.get("cell") == "H2"]
        if not helper_rewrites:
            failures.append("measure helper cell H2 was not rewritten")
        downstream_impacts = change.get("downstreamFormulaImpacts", [])
        if not downstream_impacts:
            failures.append("missing downstream helper-cell impact")
        elif downstream_impacts[0].get("cell") != "I2" or "report!H2" not in downstream_impacts[0].get("dependsOnAffectedCells", []):
            failures.append("downstream impact did not point from I2 to helper cell H2")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="reviewable DAX, CUBE, helper-cell, and dynamic-period rewrite plan verified" if not failures else "; ".join(failures),
            metadata={
                "daxRewriteCount": report.get("daxRewriteCount"),
                "cubeRewriteCount": report.get("cubeRewriteCount"),
                "downstreamFormulaImpactCount": report.get("downstreamFormulaImpactCount"),
                "manualReviewCount": report.get("manualReviewCount"),
            },
        )


def measure_delete_rewrite_plan_fixture_check(project_root: Path) -> CheckResult:
    name = "Measure delete rewrite plan fixture smoke"
    create_script = project_root / "tools" / "create_cube_formula_fixture.py"
    inspect_script = project_root / "tools" / "inspect_excel_bi_workbook.py"
    plan_script = project_root / "tools" / "build_measure_rename_rewrite_plan.py"
    for script in [create_script, inspect_script, plan_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_measure_delete_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "cube_formula_fixture.xlsx"
        model_json = tmp_dir / "model.json"
        openxml_json = tmp_dir / "openxml.json"
        plan_json = tmp_dir / "delete_plan.json"

        create_result = run_command([sys.executable, str(create_script), "--workbook", str(workbook), "--model-json", str(model_json)], project_root, f"{name}: create fixture")
        if create_result.status != PASS:
            create_result.name = name
            return create_result
        inspect_result = run_command([sys.executable, str(inspect_script), str(workbook), "--out-json", str(openxml_json)], project_root, f"{name}: inspect fixture")
        if inspect_result.status != PASS:
            inspect_result.name = name
            return inspect_result

        try:
            model = json.loads(model_json.read_text(encoding="utf-8"))
            measures = list(model.get("measures", []))
            measures.append(
                {
                    "name": "Profit",
                    "associatedTable": {"name": "Fact"},
                    "formula": "[Revenue] - SUM(Fact[Revenue])",
                }
            )
            model["measures"] = measures
            model["measureCount"] = len(measures)
            model_json.write_text(json.dumps(model, indent=2), encoding="utf-8")

            openxml = json.loads(openxml_json.read_text(encoding="utf-8"))
            formulas = list(openxml.get("cubeFormulas", []))
            formulas.extend(
                [
                    {
                        "sheet": "Report",
                        "cell": "H2",
                        "formula": 'CUBEMEMBER("ThisWorkbookDataModel","[Measures].[Revenue]","Revenue helper")',
                        "formulaType": "bi",
                        "cachedValue": "Revenue",
                    },
                    {
                        "sheet": "Report",
                        "cell": "I2",
                        "formula": 'CUBEVALUE("ThisWorkbookDataModel",$H$2,$A$2)',
                        "formulaType": "bi",
                        "cachedValue": "120",
                    },
                ]
            )
            openxml["cubeFormulas"] = formulas
            openxml["cubeFormulaCount"] = len(formulas)
            openxml_json.write_text(json.dumps(openxml, indent=2), encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not prepare delete fixture: {exc}")

        plan_result = run_command(
            [
                sys.executable,
                str(plan_script),
                "--model-json",
                str(model_json),
                "--openxml-json",
                str(openxml_json),
                "--delete",
                "Revenue",
                "--out-json",
                str(plan_json),
                "--fail-on-manual-review",
            ],
            project_root,
            f"{name}: build plan",
            ok_codes={1},
        )
        if plan_result.status != PASS:
            plan_result.name = name
            plan_result.status = FAIL
            plan_result.detail = "expected delete plan to fail on manual-review items"
            return plan_result

        try:
            report = json.loads(plan_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read delete plan: {exc}")

        changes = report.get("changes", [])
        failures: list[str] = []
        if len(changes) != 1:
            failures.append(f"expected 1 change, got {len(changes)}")
        change = changes[0] if changes else {}
        if report.get("daxRewriteCount") != 0:
            failures.append(f"daxRewriteCount={report.get('daxRewriteCount')}")
        if report.get("cubeRewriteCount") != 0:
            failures.append(f"cubeRewriteCount={report.get('cubeRewriteCount')}")
        if report.get("downstreamFormulaImpactCount") != 1:
            failures.append(f"downstreamFormulaImpactCount={report.get('downstreamFormulaImpactCount')}")
        if report.get("manualReviewCount") != 6:
            failures.append(f"manualReviewCount={report.get('manualReviewCount')}")
        manual = change.get("manualReview", [])
        kinds = {item.get("kind") for item in manual}
        expected_kinds = {"dax", "cube", "cube-downstream"}
        missing_kinds = expected_kinds - kinds
        if missing_kinds:
            failures.append(f"missing manual-review kinds={sorted(missing_kinds)}")
        downstream_manual = [item for item in manual if item.get("kind") == "cube-downstream"]
        if not downstream_manual:
            failures.append("missing downstream manual-review item")
        elif downstream_manual[0].get("cell") != "I2" or "report!H2" not in downstream_manual[0].get("dependsOnAffectedCells", []):
            failures.append("downstream manual-review item did not point from I2 to H2")
        if change.get("daxRewrites") or change.get("cubeRewrites"):
            failures.append("delete plan should not propose formula rewrites")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="delete plan blocks direct and downstream measure references for manual review" if not failures else "; ".join(failures),
            metadata={
                "manualReviewCount": report.get("manualReviewCount"),
                "downstreamFormulaImpactCount": report.get("downstreamFormulaImpactCount"),
                "manualReviewKinds": sorted(kinds),
            },
        )


def escaped_mdx_measure_reference_fixture_check(project_root: Path) -> CheckResult:
    name = "Escaped MDX measure reference fixture smoke"
    impact_script = project_root / "tools" / "analyze_measure_rename_impact.py"
    plan_script = project_root / "tools" / "build_measure_rename_rewrite_plan.py"
    cube_report_script = project_root / "tools" / "build_cube_dependency_report.py"
    model_report_script = project_root / "tools" / "build_excel_bi_model_report.py"
    for script in [impact_script, plan_script, cube_report_script, model_report_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_escaped_mdx_") as tmp:
        tmp_dir = Path(tmp)
        model_json = tmp_dir / "model.json"
        openxml_json = tmp_dir / "openxml.json"
        impact_json = tmp_dir / "impact.json"
        plan_json = tmp_dir / "rewrite_plan.json"
        cube_report_json = tmp_dir / "cube_report.json"
        model_report_json = tmp_dir / "model_report.json"

        model = {
            "workbookPath": "escaped_mdx_fixture.xlsx",
            "modelAvailable": True,
            "tableCount": 1,
            "relationshipCount": 0,
            "measureCount": 2,
            "measures": [
                {
                    "name": "Revenue ] Special",
                    "associatedTable": {"name": "Fact"},
                    "formula": "SUM(Fact[Amount])",
                },
                {
                    "name": "Awareness",
                    "associatedTable": {"name": "Fact"},
                    "formula": "AVERAGE(Fact[AwarenessScore])",
                },
            ],
            "tables": [{"name": "Fact", "sourceName": "FixtureFact", "recordCount": 2, "columnCount": 2}],
            "relationships": [],
        }
        openxml = {
            "workbookPath": "escaped_mdx_fixture.xlsx",
            "cubeFormulaCount": 4,
            "cubeFormulas": [
                {
                    "sheet": "Report",
                    "cell": "A2",
                    "formula": 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue ]] Special]","[Calendar].[Quarter].[All].[2026Q1]")',
                    "formulaType": "bi",
                    "cachedValue": "100",
                },
                {
                    "sheet": "Report",
                    "cell": "B2",
                    "formula": 'CUBEMEMBER("ThisWorkbookDataModel","[Measures].[Revenue ]] Special]","Escaped helper")',
                    "formulaType": "bi",
                    "cachedValue": "Revenue ] Special",
                },
                {
                    "sheet": "Report",
                    "cell": "C2",
                    "formula": 'CUBEVALUE("ThisWorkbookDataModel",$B$2)',
                    "formulaType": "bi",
                    "cachedValue": "100",
                },
                {
                    "sheet": "Report",
                    "cell": "D2",
                    "formula": 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Awareness]")',
                    "formulaType": "bi",
                    "cachedValue": "42",
                },
            ],
        }
        model_json.write_text(json.dumps(model, indent=2), encoding="utf-8")
        openxml_json.write_text(json.dumps(openxml, indent=2), encoding="utf-8")

        rename_arg = "Revenue ] Special=Net ] Revenue"
        commands = [
            (
                "impact",
                [
                    sys.executable,
                    str(impact_script),
                    "--model-json",
                    str(model_json),
                    "--openxml-json",
                    str(openxml_json),
                    "--rename",
                    rename_arg,
                    "--out-json",
                    str(impact_json),
                ],
            ),
            (
                "plan",
                [
                    sys.executable,
                    str(plan_script),
                    "--model-json",
                    str(model_json),
                    "--openxml-json",
                    str(openxml_json),
                    "--rename",
                    rename_arg,
                    "--out-json",
                    str(plan_json),
                ],
            ),
            (
                "cube report",
                [
                    sys.executable,
                    str(cube_report_script),
                    "--openxml-json",
                    str(openxml_json),
                    "--model-json",
                    str(model_json),
                    "--out-json",
                    str(cube_report_json),
                ],
            ),
            (
                "model report",
                [
                    sys.executable,
                    str(model_report_script),
                    "--model-json",
                    str(model_json),
                    "--openxml-json",
                    str(openxml_json),
                    "--out-json",
                    str(model_report_json),
                ],
            ),
        ]
        for label, command in commands:
            result = run_command(command, project_root, f"{name}: {label}")
            if result.status != PASS:
                result.name = name
                return result

        try:
            impact = json.loads(impact_json.read_text(encoding="utf-8"))
            plan = json.loads(plan_json.read_text(encoding="utf-8"))
            cube_report = json.loads(cube_report_json.read_text(encoding="utf-8"))
            model_report = json.loads(model_report_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read escaped MDX reports: {exc}")

        failures: list[str] = []
        impact_change = (impact.get("changes") or [{}])[0]
        plan_change = (plan.get("changes") or [{}])[0]
        if impact_change.get("cubeFormulaHitCount") != 2:
            failures.append(f"cubeFormulaHitCount={impact_change.get('cubeFormulaHitCount')}")
        if plan.get("cubeRewriteCount") != 2:
            failures.append(f"cubeRewriteCount={plan.get('cubeRewriteCount')}")
        if plan.get("downstreamFormulaImpactCount") != 1:
            failures.append(f"downstreamFormulaImpactCount={plan.get('downstreamFormulaImpactCount')}")
        if plan.get("manualReviewCount") != 0:
            failures.append(f"manualReviewCount={plan.get('manualReviewCount')}")
        new_formulas = [item.get("newFormula", "") for item in plan_change.get("cubeRewrites", [])]
        if not new_formulas or not all("[Measures].[Net ]] Revenue]" in formula for formula in new_formulas):
            failures.append("escaped replacement [Measures].[Net ]] Revenue] not found in all rewrites")
        downstream = plan_change.get("downstreamFormulaImpacts", [])
        if not downstream or "report!B2" not in downstream[0].get("dependsOnAffectedCells", []):
            failures.append("downstream helper dependency from C2 to B2 not detected")
        if cube_report.get("byMeasure", {}).get("Revenue ] Special") != 2:
            failures.append(f"cube report byMeasure={cube_report.get('byMeasure')}")
        if "Revenue ] Special" not in model_report.get("measuresReferencedByCubeFormulas", []):
            failures.append("model report missing escaped measure reference")
        if model_report.get("cubeFormulaReferencesMissingModelMeasure"):
            failures.append(f"model report false missing measures={model_report.get('cubeFormulaReferencesMissingModelMeasure')}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="escaped MDX measure names are recognized across impact, reports, and rewrite planning" if not failures else "; ".join(failures),
            metadata={
                "impactCubeFormulaHitCount": impact_change.get("cubeFormulaHitCount"),
                "cubeRewriteCount": plan.get("cubeRewriteCount"),
                "downstreamFormulaImpactCount": plan.get("downstreamFormulaImpactCount"),
            },
        )


def vba_source_lint_fixture_check(project_root: Path) -> CheckResult:
    name = "VBA source lint fixture smoke"
    script = project_root / ".agents" / "skills" / "excel-vba-workbook-engineering" / "scripts" / "lint_vba_source.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")
    with tempfile.TemporaryDirectory(prefix="excel_bi_vba_lint_") as tmp:
        tmp_dir = Path(tmp)
        source_dir = tmp_dir / "vba"
        source_dir.mkdir()
        (source_dir / "modMain.bas").write_text(
            "\n".join(
                [
                    'Attribute VB_Name = "modMain"',
                    "Option Explicit",
                    "",
                    "Public Sub RunReport()",
                    "    Dim value As Long",
                    "    value = HelperValue()",
                    "End Sub",
                    "",
                    "Private Function HelperValue() As Long",
                    "    HelperValue = 1",
                    "End Function",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        (source_dir / "ThisWorkbook.cls").write_text(
            "\n".join(
                [
                    'Attribute VB_Name = "ThisWorkbook"',
                    "Option Explicit",
                    "",
                    "Private Sub Workbook_Open()",
                    "End Sub",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        out_json = tmp_dir / "vba_lint.json"
        result = run_command([sys.executable, str(script), str(source_dir), "--strict-option-explicit", "--out-json", str(out_json)], project_root, name)
        if result.status != PASS:
            return result
        try:
            report = json.loads(out_json.read_text(encoding="utf-8"))
        except Exception as exc:
            result.status = FAIL
            result.detail = f"could not read lint report: {exc}"
            return result
        failures: list[str] = []
        if report.get("moduleCount") != 2:
            failures.append(f"expected 2 modules, got {report.get('moduleCount')}")
        if report.get("procedureCount") != 3:
            failures.append(f"expected 3 procedures, got {report.get('procedureCount')}")
        if report.get("publicEntryCount") != 1:
            failures.append(f"expected 1 public entry, got {report.get('publicEntryCount')}")
        if report.get("errors"):
            failures.append(f"expected no lint errors, got {report.get('errors')}")
        if failures:
            result.status = FAIL
            result.detail = "; ".join(failures)
        else:
            result.detail = "2 modules, 3 procedures, 1 public entry verified"
            result.metadata = {
                "moduleCount": report.get("moduleCount"),
                "procedureCount": report.get("procedureCount"),
                "publicEntryCount": report.get("publicEntryCount"),
            }
        return result


def vba_button_binding_report_fixture_check(project_root: Path) -> CheckResult:
    name = "VBA button binding report fixture smoke"
    fixture_script = project_root / "tools" / "create_vba_button_binding_fixture.py"
    report_script = project_root / "tools" / "build_vba_button_binding_report.py"
    missing = [str(path) for path in [fixture_script, report_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_vba_button_binding_") as tmp:
        tmp_dir = Path(tmp)
        fixture_dir = tmp_dir / "fixture"
        manifest_json = tmp_dir / "fixture_manifest.json"
        report_json = tmp_dir / "vba_button_bindings.json"
        report_md = tmp_dir / "vba_button_bindings.md"
        strict_report_json = tmp_dir / "vba_button_bindings_strict.json"

        fixture_result = run_command(
            [
                sys.executable,
                str(fixture_script),
                "--out-dir",
                str(fixture_dir),
                "--out-json",
                str(manifest_json),
            ],
            project_root,
            name,
        )
        if fixture_result.status != PASS:
            return fixture_result

        try:
            manifest = json.loads(manifest_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read fixture manifest: {exc}")

        workbook_inventory_json = Path(str(manifest.get("workbookInventoryJson", "")))
        vba_lint_json = Path(str(manifest.get("vbaLintJson", "")))
        expected = manifest.get("expected", {})
        if not isinstance(expected, dict):
            return CheckResult(name=name, status=FAIL, detail="fixture manifest expected section is missing")

        report_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--workbook-inventory-json",
                str(workbook_inventory_json),
                "--vba-lint-json",
                str(vba_lint_json),
                "--out-json",
                str(report_json),
                "--out-md",
                str(report_md),
            ],
            project_root,
            name,
        )
        if report_result.status != PASS:
            return report_result

        strict_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--workbook-inventory-json",
                str(workbook_inventory_json),
                "--vba-lint-json",
                str(vba_lint_json),
                "--out-json",
                str(strict_report_json),
                "--fail-on-unresolved",
            ],
            project_root,
            name,
            ok_codes={1},
        )
        if strict_result.status != PASS:
            return strict_result

        try:
            report = json.loads(report_json.read_text(encoding="utf-8"))
            report_text = report_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read VBA button binding reports: {exc}")

        summary = report.get("summary", {})
        findings = report.get("findings", [])
        failures: list[str] = []
        if report.get("status") != FAIL:
            failures.append(f"expected report status fail, got {report.get('status')}")
        for field in ["shapeActionCount", "resolvedCount", "missingMacroCount"]:
            if summary.get(field) != expected.get(field):
                failures.append(f"expected {field}={expected.get(field)}, got {summary.get(field)}")
        expected_missing = expected.get("missingMacro")
        if not any(
            isinstance(item, dict)
            and item.get("code") == "missing-onaction-macro"
            and item.get("normalizedMacro") == expected_missing
            for item in findings
        ):
            failures.append(f"missing expected unresolved macro finding for {expected_missing}")
        if "VBA Button Binding Report" not in report_text:
            failures.append("markdown report heading missing")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))

        return CheckResult(
            name=name,
            status=PASS,
            detail="2 resolved bindings and 1 missing OnAction macro verified",
            metadata={
                "shapeActionCount": summary.get("shapeActionCount"),
                "resolvedCount": summary.get("resolvedCount"),
                "missingMacroCount": summary.get("missingMacroCount"),
                "strictFailurePathVerified": True,
            },
        )


def vba_import_export_run_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "VBA import/export/run fixture smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")

    skill_scripts = project_root / ".agents" / "skills" / "excel-vba-workbook-engineering" / "scripts"
    import_script = skill_scripts / "import_vba.ps1"
    export_script = skill_scripts / "export_vba.ps1"
    missing = [str(path) for path in [import_script, export_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    unavailable_markers = [
        "cannot create activex component",
        "class not registered",
        "excel.application",
        "programmatic access to visual basic project is not trusted",
        "access to visual basic project is not trusted",
        "trust access to the vba project object model",
        "vbproject",
        "retrieving the com class factory",
    ]

    with tempfile.TemporaryDirectory(prefix="excel_bi_vba_roundtrip_") as tmp:
        tmp_dir = Path(tmp)
        base_workbook = tmp_dir / "base.xlsx"
        output_workbook = tmp_dir / "probe.xlsm"
        source_dir = tmp_dir / "src"
        exported_dir = tmp_dir / "exported"
        run_json = tmp_dir / "run_smoke.json"
        create_script = tmp_dir / "create_base_workbook.ps1"
        run_script = tmp_dir / "run_smoke_macro.ps1"
        source_dir.mkdir()

        (source_dir / "modSmoke.bas").write_text(
            "\r\n".join(
                [
                    'Attribute VB_Name = "modSmoke"',
                    "Option Explicit",
                    "",
                    "Public Sub RunSmoke()",
                    '    With ThisWorkbook.Worksheets("Data")',
                    '        .Range("B2").Value = "PASS"',
                    "        .Range(\"B3\").Value = 12345",
                    "    End With",
                    "End Sub",
                    "",
                ]
            ),
            encoding="ascii",
        )

        create_script.write_text(
            r'''
param([Parameter(Mandatory = $true)][string]$WorkbookPath)
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "Data"
    $sheet.Range("A1").Value2 = "Status"
    $sheet.Range("B1").Value2 = "Value"
    $workbook.SaveAs($WorkbookPath, 51)
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''.lstrip(),
            encoding="utf-8",
        )

        run_script.write_text(
            r'''
param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$OutJson
)
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($WorkbookPath, $null, $false)
    $macroName = "'" + $workbook.Name + "'!RunSmoke"
    [void]$excel.Run($macroName)
    $sheet = $workbook.Worksheets.Item("Data")
    $status = [string]$sheet.Range("B2").Value2
    $value = [double]$sheet.Range("B3").Value2
    $hasVBProject = [bool]$workbook.HasVBProject
    $workbook.Save()
    [ordered]@{
        workbookPath = $WorkbookPath
        macroName = $macroName
        status = $status
        value = $value
        hasVBProject = $hasVBProject
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $OutJson -Encoding UTF8
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''.lstrip(),
            encoding="utf-8",
        )

        create_result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(create_script),
                "-WorkbookPath",
                str(base_workbook),
            ],
            project_root,
            name,
        )
        if create_result.status != PASS:
            combined = f"{create_result.stdout}\n{create_result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                create_result.status = SKIP
                create_result.detail = "Excel COM runtime unavailable while creating base workbook"
            return create_result

        import_result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(import_script),
                "-WorkbookPath",
                str(base_workbook),
                "-SourceDir",
                str(source_dir),
                "-OutputWorkbookPath",
                str(output_workbook),
            ],
            project_root,
            name,
        )
        if import_result.status != PASS:
            combined = f"{import_result.stdout}\n{import_result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                import_result.status = SKIP
                import_result.detail = "Excel VBA project import unavailable; Trust access to the VBA project object model may be disabled"
            return import_result

        try:
            import_report = json.loads(import_result.stdout)
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not parse import_vba output: {exc}")

        run_result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(run_script),
                "-WorkbookPath",
                str(output_workbook),
                "-OutJson",
                str(run_json),
            ],
            project_root,
            name,
        )
        if run_result.status != PASS:
            combined = f"{run_result.stdout}\n{run_result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                run_result.status = SKIP
                run_result.detail = "Excel macro runtime unavailable while running imported fixture macro"
            return run_result

        export_result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(export_script),
                "-WorkbookPath",
                str(output_workbook),
                "-OutDir",
                str(exported_dir),
            ],
            project_root,
            name,
        )
        if export_result.status != PASS:
            combined = f"{export_result.stdout}\n{export_result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                export_result.status = SKIP
                export_result.detail = "Excel VBA project export unavailable; Trust access to the VBA project object model may be disabled"
            return export_result

        try:
            run_report = json.loads(run_json.read_text(encoding="utf-8-sig"))
            export_report = json.loads(export_result.stdout)
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read VBA fixture reports: {exc}")

        imported = [Path(str(item)).name for item in import_report.get("imported", [])]
        exported = [Path(str(item)).name for item in export_report.get("exported", [])]
        exported_module = exported_dir / "modSmoke.bas"
        failures: list[str] = []
        if "modSmoke.bas" not in imported:
            failures.append(f"imported={imported}")
        if str(import_report.get("outputWorkbook", "")).lower() != str(output_workbook).lower():
            failures.append("outputWorkbook did not match fixture .xlsm")
        if run_report.get("status") != "PASS":
            failures.append(f"macro status={run_report.get('status')}")
        if float(run_report.get("value", 0) or 0) != 12345:
            failures.append(f"macro value={run_report.get('value')}")
        if run_report.get("hasVBProject") is not True:
            failures.append(f"hasVBProject={run_report.get('hasVBProject')}")
        if "modSmoke.bas" not in exported:
            failures.append(f"exported={exported}")
        if not exported_module.is_file():
            failures.append("exported modSmoke.bas missing")
        elif "Public Sub RunSmoke" not in exported_module.read_text(encoding="utf-8", errors="ignore"):
            failures.append("exported modSmoke.bas missing RunSmoke procedure")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="VBA module import, macro execution, and export round-trip verified" if not failures else "; ".join(failures),
            metadata={
                "imported": imported,
                "exported": exported,
                "macroStatus": run_report.get("status"),
                "macroValue": run_report.get("value"),
                "hasVBProject": run_report.get("hasVBProject"),
            },
        )


def dax_compat_lint_fixture_check(project_root: Path) -> CheckResult:
    name = "DAX compatibility lint fixture smoke"
    script = project_root / ".agents" / "skills" / "power-pivot-dax-modeling" / "scripts" / "lint_dax_compat.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")
    with tempfile.TemporaryDirectory(prefix="excel_bi_dax_lint_") as tmp:
        tmp_dir = Path(tmp)
        good = tmp_dir / "good.dax"
        bad = tmp_dir / "bad.dax"
        version_sensitive = tmp_dir / "version_sensitive.dax"
        division = tmp_dir / "division.dax"
        good_json = tmp_dir / "good.json"
        bad_json = tmp_dir / "bad.json"
        version_sensitive_json = tmp_dir / "version_sensitive.json"
        division_json = tmp_dir / "division.json"
        good.write_text(
            "Share := DIVIDE([Sales], CALCULATE([Sales], ALL('DimProduct')))\n",
            encoding="utf-8",
        )
        bad.write_text(
            "Share := DIVIDE([Sales], CALCULATE([Sales], REMOVEFILTERS('DimProduct')))\n",
            encoding="utf-8",
        )
        version_sensitive.write_text(
            'Selected Period := SELECTEDVALUE(\'Period\'[Period], "All")\n',
            encoding="utf-8",
        )
        division.write_text(
            "Unsafe Ratio := [Sales] / [Customers]\n",
            encoding="utf-8",
        )

        good_result = run_command([sys.executable, str(script), str(good), "--out-json", str(good_json)], project_root, f"{name}: compatible formula")
        if good_result.status != PASS:
            good_result.name = name
            return good_result

        bad_result = run_command([sys.executable, str(script), str(bad), "--out-json", str(bad_json)], project_root, f"{name}: incompatible formula", ok_codes={1})
        if bad_result.status != PASS:
            bad_result.name = name
            bad_result.status = FAIL
            bad_result.detail = "expected REMOVEFILTERS fixture to fail"
            return bad_result

        version_sensitive_result = run_command(
            [sys.executable, str(script), str(version_sensitive), "--out-json", str(version_sensitive_json)],
            project_root,
            f"{name}: version-sensitive formula",
        )
        if version_sensitive_result.status != PASS:
            version_sensitive_result.name = name
            return version_sensitive_result

        division_result = run_command(
            [sys.executable, str(script), str(division), "--warn-division", "--out-json", str(division_json)],
            project_root,
            f"{name}: division warning formula",
        )
        if division_result.status != PASS:
            division_result.name = name
            return division_result

        try:
            good_report = json.loads(good_json.read_text(encoding="utf-8"))
            bad_report = json.loads(bad_json.read_text(encoding="utf-8"))
            version_sensitive_report = json.loads(version_sensitive_json.read_text(encoding="utf-8"))
            division_report = json.loads(division_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read lint reports: {exc}")

        failures: list[str] = []
        if good_report.get("errorCount") != 0:
            failures.append(f"compatible fixture errorCount={good_report.get('errorCount')}")
        if bad_report.get("errorCount") != 1:
            failures.append(f"incompatible fixture errorCount={bad_report.get('errorCount')}")
        if "REMOVEFILTERS" not in bad_report.get("functionCounts", {}):
            failures.append("incompatible fixture did not report REMOVEFILTERS")
        if version_sensitive_report.get("errorCount") != 0 or version_sensitive_report.get("warningCount") != 1:
            failures.append(
                "version-sensitive fixture expected 0 errors and 1 warning, "
                f"got errors={version_sensitive_report.get('errorCount')} warnings={version_sensitive_report.get('warningCount')}"
            )
        version_sensitive_codes = {issue.get("code") for issue in version_sensitive_report.get("issues", [])}
        if "excel-version-sensitive-function" not in version_sensitive_codes:
            failures.append(f"version-sensitive fixture codes={sorted(version_sensitive_codes)}")
        if "SELECTEDVALUE" not in version_sensitive_report.get("functionCounts", {}):
            failures.append("version-sensitive fixture did not report SELECTEDVALUE")
        if division_report.get("errorCount") != 0 or division_report.get("warningCount") != 1:
            failures.append(
                "division fixture expected 0 errors and 1 warning, "
                f"got errors={division_report.get('errorCount')} warnings={division_report.get('warningCount')}"
            )
        division_codes = {issue.get("code") for issue in division_report.get("issues", [])}
        if "operator-division" not in division_codes:
            failures.append(f"division fixture codes={sorted(division_codes)}")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))
        return CheckResult(
            name=name,
            status=PASS,
            detail="compatible formula passed, REMOVEFILTERS failed, and SELECTEDVALUE/division warnings were reported as expected",
            metadata={
                "compatibleExpressionCount": good_report.get("expressionCount"),
                "incompatibleErrorCount": bad_report.get("errorCount"),
                "incompatibleFunctions": bad_report.get("functionCounts"),
                "versionSensitiveWarningCount": version_sensitive_report.get("warningCount"),
                "divisionWarningCount": division_report.get("warningCount"),
            },
        )


def dax_dependency_fixture_check(project_root: Path) -> CheckResult:
    name = "DAX dependency analysis fixture smoke"
    script = project_root / ".agents" / "skills" / "power-pivot-dax-modeling" / "scripts" / "analyze_dax_dependencies.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")
    with tempfile.TemporaryDirectory(prefix="excel_bi_dax_dependency_") as tmp:
        tmp_dir = Path(tmp)
        good = tmp_dir / "good_model.json"
        bad = tmp_dir / "bad_model.json"
        good_json = tmp_dir / "good_report.json"
        bad_json = tmp_dir / "bad_report.json"
        good.write_text(
            json.dumps(
                {
                    "measures": [
                        {"name": "Sales", "formula": "SUM(Fact[Amount])"},
                        {"name": "Margin", "formula": "[Sales] - SUM(Fact[Cost])"},
                        {"name": "Margin %", "formula": "DIVIDE([Margin], [Sales])"},
                    ]
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        bad.write_text(
            json.dumps(
                {
                    "measures": [
                        {"name": "Sales", "formula": "SUM(Fact[Amount]) + [Missing Measure]"},
                        {"name": "Self", "formula": "[Self] + 1"},
                        {"name": "A", "formula": "[B] + 1"},
                        {"name": "B", "formula": "[A] + 1"},
                    ]
                },
                indent=2,
            ),
            encoding="utf-8",
        )

        good_result = run_command([sys.executable, str(script), str(good), "--out-json", str(good_json)], project_root, f"{name}: valid dependencies")
        if good_result.status != PASS:
            good_result.name = name
            return good_result

        bad_result = run_command([sys.executable, str(script), str(bad), "--out-json", str(bad_json)], project_root, f"{name}: invalid dependencies", ok_codes={1})
        if bad_result.status != PASS:
            bad_result.name = name
            bad_result.status = FAIL
            bad_result.detail = "expected invalid dependency fixture to fail"
            return bad_result

        try:
            good_report = json.loads(good_json.read_text(encoding="utf-8"))
            bad_report = json.loads(bad_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read dependency reports: {exc}")

        bad_codes = {issue.get("code") for issue in bad_report.get("issues", [])}
        expected = {"missing-measure-reference", "self-reference", "dependency-cycle"}
        failures: list[str] = []
        if good_report.get("errorCount") != 0:
            failures.append(f"valid fixture errorCount={good_report.get('errorCount')}")
        missing = expected - bad_codes
        if missing:
            failures.append(f"invalid fixture missing codes={sorted(missing)}")
        if bad_report.get("errorCount", 0) < 3:
            failures.append(f"invalid fixture errorCount={bad_report.get('errorCount')}")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))
        return CheckResult(
            name=name,
            status=PASS,
            detail="valid dependency graph passed and invalid graph failed as expected",
            metadata={
                "validMeasureCount": good_report.get("measureCount"),
                "invalidErrorCount": bad_report.get("errorCount"),
                "invalidCodes": sorted(bad_codes),
            },
        )


def power_query_m_lint_fixture_check(project_root: Path) -> CheckResult:
    name = "Power Query M lint fixture smoke"
    script = project_root / ".agents" / "skills" / "power-query-m-engineering" / "scripts" / "lint_power_query_m.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")
    with tempfile.TemporaryDirectory(prefix="excel_bi_pq_m_lint_") as tmp:
        tmp_dir = Path(tmp)
        good = tmp_dir / "good.m"
        risky = tmp_dir / "risky.m"
        good_json = tmp_dir / "good.json"
        risky_json = tmp_dir / "risky.json"
        good.write_text(
            "\n".join(
                [
                    "let",
                    '    Source = Folder.Files("C:\\Data"),',
                    '    Filtered = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$")),',
                    '    Kept = Table.SelectColumns(Filtered, {"Name", "Content"}, MissingField.UseNull),',
                    '    Sorted = Table.Sort(Kept, {{"Name", Order.Ascending}})',
                    "in",
                    "    Sorted",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        risky.write_text(
            "\n".join(
                [
                    "let",
                    '    Source = Folder.Files("C:\\Data"),',
                    '    Joined = Table.NestedJoin(Source, {"Name"}, Source, {"Name"}, "Hit", JoinKind.LeftOuter),',
                    '    Expanded = Table.ExpandTableColumn(Joined, "Hit", {"Content"})',
                    "in",
                    "    Expanded",
                    "",
                ]
            ),
            encoding="utf-8",
        )

        good_result = run_command([sys.executable, str(script), str(good), "--out-json", str(good_json)], project_root, f"{name}: safe query")
        if good_result.status != PASS:
            good_result.name = name
            return good_result

        risky_result = run_command(
            [sys.executable, str(script), str(risky), "--warnings-as-errors", "--out-json", str(risky_json)],
            project_root,
            f"{name}: risky query",
            ok_codes={1},
        )
        if risky_result.status != PASS:
            risky_result.name = name
            risky_result.status = FAIL
            risky_result.detail = "expected risky M fixture to fail with warnings as errors"
            return risky_result

        try:
            good_report = json.loads(good_json.read_text(encoding="utf-8"))
            risky_report = json.loads(risky_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read lint reports: {exc}")

        failures: list[str] = []
        if good_report.get("errorCount") != 0 or good_report.get("warningCount") != 0:
            failures.append(
                f"safe fixture errors={good_report.get('errorCount')} warnings={good_report.get('warningCount')}"
            )
        risky_codes = {issue.get("code") for issue in risky_report.get("issues", [])}
        expected_codes = {"folder-files-temp-filter", "folder-files-hidden-filter", "join-cardinality", "hard-coded-expand-columns"}
        missing_codes = expected_codes - risky_codes
        if missing_codes:
            failures.append(f"risky fixture missing codes={sorted(missing_codes)}")
        if risky_report.get("errorCount", 0) < len(expected_codes):
            failures.append(f"risky fixture errorCount={risky_report.get('errorCount')}")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))
        return CheckResult(
            name=name,
            status=PASS,
            detail="safe query passed and risky Folder.Files/join/expand query failed as expected",
            metadata={
                "safeQueryCount": good_report.get("queryCount"),
                "riskyErrorCount": risky_report.get("errorCount"),
                "riskyCodes": sorted(risky_codes),
            },
        )


def power_query_lineage_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Power Query lineage/source-risk report fixture smoke"
    fixture_script = project_root / "tools" / "create_power_query_lineage_fixture.py"
    report_script = project_root / "tools" / "build_power_query_lineage_report.py"
    missing = [str(path) for path in [fixture_script, report_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_pq_lineage_") as tmp:
        tmp_dir = Path(tmp)
        fixture_dir = tmp_dir / "fixture"
        manifest_json = tmp_dir / "fixture_manifest.json"
        safe_json = tmp_dir / "safe_report.json"
        safe_md = tmp_dir / "safe_report.md"
        risky_json = tmp_dir / "risky_report.json"
        risky_md = tmp_dir / "risky_report.md"

        fixture_result = run_command(
            [
                sys.executable,
                str(fixture_script),
                "--out-dir",
                str(fixture_dir),
                "--out-json",
                str(manifest_json),
            ],
            project_root,
            f"{name}: create fixture",
        )
        if fixture_result.status != PASS:
            fixture_result.name = name
            return fixture_result

        try:
            manifest = json.loads(manifest_json.read_text(encoding="utf-8"))
            safe = manifest["safe"]
            risky = manifest["risky"]
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read lineage fixture manifest: {exc}")

        safe_result = run_command(
            [
                sys.executable,
                str(report_script),
                safe["queryDir"],
                "--out-json",
                str(safe_json),
                "--out-md",
                str(safe_md),
                "--fail-on-high-risk",
            ],
            project_root,
            f"{name}: safe report",
        )
        if safe_result.status != PASS:
            safe_result.name = name
            return safe_result

        risky_result = run_command(
            [
                sys.executable,
                str(report_script),
                risky["queryDir"],
                "--out-json",
                str(risky_json),
                "--out-md",
                str(risky_md),
            ],
            project_root,
            f"{name}: risky report",
        )
        if risky_result.status != PASS:
            risky_result.name = name
            return risky_result

        strict_result = run_command(
            [
                sys.executable,
                str(report_script),
                risky["queryDir"],
                "--fail-on-high-risk",
            ],
            project_root,
            f"{name}: risky strict report",
            ok_codes={1},
        )
        if strict_result.status != PASS:
            strict_result.name = name
            strict_result.status = FAIL
            strict_result.detail = "expected risky lineage fixture to fail with --fail-on-high-risk"
            return strict_result

        try:
            safe_report = json.loads(safe_json.read_text(encoding="utf-8"))
            risky_report = json.loads(risky_json.read_text(encoding="utf-8"))
            risky_markdown = risky_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read lineage reports: {exc}")

        safe_expected = safe.get("expected", {})
        risky_expected = risky.get("expected", {})
        safe_summary = safe_report.get("summary", {})
        risky_summary = risky_report.get("summary", {})
        risky_codes = {finding.get("code") for finding in risky_report.get("findings", [])}
        expected_codes = set(risky_expected.get("requiredCodes", []))
        failures: list[str] = []
        if safe_summary.get("readiness") != safe_expected.get("readiness"):
            failures.append(f"safe readiness={safe_summary.get('readiness')}")
        if safe_summary.get("queryCount") != safe_expected.get("queryCount"):
            failures.append(f"safe queryCount={safe_summary.get('queryCount')}")
        if safe_summary.get("dependencyCount") != safe_expected.get("dependencyCount"):
            failures.append(f"safe dependencyCount={safe_summary.get('dependencyCount')}")
        if safe_summary.get("findingCount") != safe_expected.get("findingCount"):
            failures.append(f"safe findingCount={safe_summary.get('findingCount')}")
        for kind, expected_count in safe_expected.get("sourceKindCounts", {}).items():
            actual_count = safe_summary.get("sourceKindCounts", {}).get(kind)
            if actual_count != expected_count:
                failures.append(f"safe sourceKindCounts[{kind}]={actual_count}")
        if risky_summary.get("readiness") != risky_expected.get("readiness"):
            failures.append(f"risky readiness={risky_summary.get('readiness')}")
        if risky_summary.get("queryCount") != risky_expected.get("queryCount"):
            failures.append(f"risky queryCount={risky_summary.get('queryCount')}")
        if risky_summary.get("dependencyCount") != risky_expected.get("dependencyCount"):
            failures.append(f"risky dependencyCount={risky_summary.get('dependencyCount')}")
        if risky_summary.get("highFindingCount") != risky_expected.get("highFindingCount"):
            failures.append(f"risky highFindingCount={risky_summary.get('highFindingCount')}")
        if risky_summary.get("mediumFindingCount") != risky_expected.get("mediumFindingCount"):
            failures.append(f"risky mediumFindingCount={risky_summary.get('mediumFindingCount')}")
        for kind, expected_count in risky_expected.get("sourceKindCounts", {}).items():
            actual_count = risky_summary.get("sourceKindCounts", {}).get(kind)
            if actual_count != expected_count:
                failures.append(f"risky sourceKindCounts[{kind}]={actual_count}")
        missing_codes = expected_codes - risky_codes
        if missing_codes:
            failures.append(f"risky fixture missing codes={sorted(missing_codes)}")
        if "# Power Query Lineage And Source-Risk Report" not in risky_markdown:
            failures.append("risky markdown heading missing")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))

        return CheckResult(
            name=name,
            status=PASS,
            detail="safe parameterized lineage passed and risky local/web/database/cloud-service/native-query/credential-like/mixed-source/cycle fixture failed strict mode as expected",
            metadata={
                "safeQueryCount": safe_summary.get("queryCount"),
                "safeSourceKindCounts": safe_summary.get("sourceKindCounts"),
                "riskyHighFindingCount": risky_summary.get("highFindingCount"),
                "riskyMediumFindingCount": risky_summary.get("mediumFindingCount"),
                "riskySourceKindCounts": risky_summary.get("sourceKindCounts"),
                "riskyCodes": sorted(risky_codes),
            },
        )


def power_query_refresh_error_classifier_fixture_check(project_root: Path) -> CheckResult:
    name = "Power Query refresh error classifier fixture smoke"
    script = project_root / ".agents" / "skills" / "power-query-m-engineering" / "scripts" / "classify_power_query_refresh_errors.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")
    with tempfile.TemporaryDirectory(prefix="excel_bi_pq_refresh_classify_") as tmp:
        tmp_dir = Path(tmp)
        failed = tmp_dir / "refresh_failed.json"
        clean = tmp_dir / "refresh_clean.json"
        failed_json = tmp_dir / "failed_report.json"
        clean_json = tmp_dir / "clean_report.json"
        failed.write_text(
            json.dumps(
                {
                    "workbookPath": "C:/Temp/book.xlsx",
                    "queryName": "SalesQuery",
                    "failedAt": "2026-06-18T00:00:00Z",
                    "error": "Expression.Error: The column 'Amount' of the table wasn't found.",
                    "errors": [
                        {
                            "phase": "CalculateUntilAsyncQueriesDone",
                            "message": "Formula.Firewall: Query 'SalesQuery' references other queries or steps, so it may not directly access a data source.",
                        },
                        {"phase": "Refresh", "message": "DataFormat.Error: We couldn't convert to Number."},
                    ],
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        clean.write_text(
            json.dumps(
                {
                    "workbookPath": "C:/Temp/book.xlsx",
                    "queryName": "",
                    "startedAt": "2026-06-18T00:00:00Z",
                    "completedAt": "2026-06-18T00:00:02Z",
                    "elapsedSeconds": 2.0,
                    "errors": [],
                },
                indent=2,
            ),
            encoding="utf-8",
        )

        failed_result = run_command([sys.executable, str(script), str(failed), "--out-json", str(failed_json)], project_root, f"{name}: failed refresh")
        if failed_result.status != PASS:
            failed_result.name = name
            return failed_result
        clean_result = run_command([sys.executable, str(script), str(clean), "--out-json", str(clean_json)], project_root, f"{name}: clean refresh")
        if clean_result.status != PASS:
            clean_result.name = name
            return clean_result

        try:
            failed_report = json.loads(failed_json.read_text(encoding="utf-8"))
            clean_report = json.loads(clean_json.read_text(encoding="utf-8"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read classifier reports: {exc}")

        failed_codes = {finding.get("code") for finding in failed_report.get("findings", [])}
        expected = {"privacy-firewall", "missing-column", "type-conversion"}
        failures: list[str] = []
        missing = expected - failed_codes
        if missing:
            failures.append(f"failed fixture missing codes={sorted(missing)}")
        if "timeout-or-background-refresh" in failed_codes:
            failures.append("failed fixture incorrectly classified CalculateUntilAsyncQueriesDone phase as timeout")
        if clean_report.get("findingCount") != 0 or clean_report.get("status") != "no-known-errors":
            failures.append(
                f"clean fixture status={clean_report.get('status')} findingCount={clean_report.get('findingCount')}"
            )
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))
        return CheckResult(
            name=name,
            status=PASS,
            detail="refresh error categories and clean refresh status verified",
            metadata={
                "failedCodes": sorted(failed_codes),
                "cleanStatus": clean_report.get("status"),
                "cleanFindingCount": clean_report.get("findingCount"),
            },
        )


def power_query_refresh_performance_report_fixture_check(project_root: Path) -> CheckResult:
    name = "Power Query refresh performance report fixture smoke"
    script = project_root / ".agents" / "skills" / "power-query-m-engineering" / "scripts" / "build_power_query_refresh_report.py"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_pq_refresh_report_") as tmp:
        tmp_dir = Path(tmp)
        clean = tmp_dir / "clean_refresh.json"
        slow = tmp_dir / "slow_refresh.json"
        failed = tmp_dir / "failed_refresh.json"
        clean_report_json = tmp_dir / "clean_report.json"
        clean_report_md = tmp_dir / "clean_report.md"
        slow_report_json = tmp_dir / "slow_report.json"
        failed_report_json = tmp_dir / "failed_report.json"
        failed_report_md = tmp_dir / "failed_report.md"

        clean.write_text(
            json.dumps(
                {
                    "workbookPath": "C:/Temp/book.xlsx",
                    "queryName": "SalesQuery",
                    "startedAt": "2026-06-19T00:00:00Z",
                    "completedAt": "2026-06-19T00:00:02Z",
                    "elapsedSeconds": 2.0,
                    "disableBackgroundRefresh": True,
                    "calculateFull": False,
                    "backgroundChanges": [
                        {
                            "sheet": "Output",
                            "listObject": "SalesTable",
                            "oldBackgroundQuery": True,
                            "newBackgroundQuery": False,
                        }
                    ],
                    "targetLoadRefreshes": [
                        {
                            "type": "ListObject.QueryTable",
                            "sheet": "Output",
                            "listObject": "SalesTable",
                            "connection": "Query - SalesQuery",
                        }
                    ],
                    "beforeConnections": [{"name": "Query - SalesQuery", "refreshing": False}],
                    "afterConnections": [{"name": "Query - SalesQuery", "refreshing": False}],
                    "errors": [],
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        slow.write_text(
            json.dumps(
                {
                    "workbookPath": "C:/Temp/book.xlsx",
                    "queryName": "SlowQuery",
                    "startedAt": "2026-06-19T00:00:00Z",
                    "completedAt": "2026-06-19T00:00:20Z",
                    "elapsedSeconds": 20.0,
                    "disableBackgroundRefresh": True,
                    "afterConnections": [{"name": "Query - SlowQuery", "refreshing": False}],
                    "errors": [],
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        failed.write_text(
            json.dumps(
                {
                    "workbookPath": "C:/Temp/book.xlsx",
                    "queryName": "BadQuery",
                    "startedAt": "2026-06-19T00:00:00Z",
                    "failedAt": "2026-06-19T00:00:03Z",
                    "elapsedSeconds": 3.0,
                    "disableBackgroundRefresh": False,
                    "error": "Expression.Error: The column 'Amount' of the table was not found.",
                    "afterConnections": [{"name": "Query - BadQuery", "refreshing": True}],
                    "errors": [{"phase": "Refresh", "message": "DataFormat.Error: We could not convert to Number."}],
                },
                indent=2,
            ),
            encoding="utf-8",
        )

        clean_result = run_command(
            [
                sys.executable,
                str(script),
                str(clean),
                "--require-completed",
                "--max-elapsed-seconds",
                "5",
                "--out-json",
                str(clean_report_json),
                "--out-md",
                str(clean_report_md),
                "--fail-on-warning",
            ],
            project_root,
            f"{name}: clean refresh",
        )
        if clean_result.status != PASS:
            clean_result.name = name
            return clean_result

        slow_result = run_command(
            [
                sys.executable,
                str(script),
                str(slow),
                "--require-completed",
                "--max-elapsed-seconds",
                "5",
                "--out-json",
                str(slow_report_json),
                "--fail-on-slow",
            ],
            project_root,
            f"{name}: slow refresh",
            ok_codes={1},
        )
        if slow_result.status != PASS:
            slow_result.name = name
            return slow_result

        failed_result = run_command(
            [
                sys.executable,
                str(script),
                str(failed),
                "--require-completed",
                "--out-json",
                str(failed_report_json),
                "--out-md",
                str(failed_report_md),
                "--fail-on-error",
            ],
            project_root,
            f"{name}: failed refresh",
            ok_codes={1},
        )
        if failed_result.status != PASS:
            failed_result.name = name
            return failed_result

        try:
            clean_report = json.loads(clean_report_json.read_text(encoding="utf-8"))
            slow_report = json.loads(slow_report_json.read_text(encoding="utf-8"))
            failed_report = json.loads(failed_report_json.read_text(encoding="utf-8"))
            clean_markdown = clean_report_md.read_text(encoding="utf-8")
            failed_markdown = failed_report_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read refresh reports: {exc}")

        failures: list[str] = []
        if clean_report.get("status") != PASS or clean_report.get("refreshStatus") != "completed":
            failures.append(f"clean status={clean_report.get('status')} refresh={clean_report.get('refreshStatus')}")
        if clean_report.get("warningFindingCount") != 0 or clean_report.get("errorFindingCount") != 0:
            failures.append(
                f"clean findings warning={clean_report.get('warningFindingCount')} error={clean_report.get('errorFindingCount')}"
            )
        if clean_report.get("targetLoadRefreshCount") != 1:
            failures.append(f"clean targetLoadRefreshCount={clean_report.get('targetLoadRefreshCount')}")
        if slow_report.get("status") != WARN or slow_report.get("elapsedStatus") != "slow":
            failures.append(f"slow status={slow_report.get('status')} elapsedStatus={slow_report.get('elapsedStatus')}")
        if not any(item.get("code") == "slow-refresh" for item in slow_report.get("findings", [])):
            failures.append("slow report missing slow-refresh finding")
        failed_codes = {item.get("code") for item in failed_report.get("findings", [])}
        required_failed_codes = {"refresh-failed", "refresh-not-completed", "connections-still-refreshing"}
        missing_failed_codes = required_failed_codes - failed_codes
        if failed_report.get("status") != FAIL or missing_failed_codes:
            failures.append(f"failed status={failed_report.get('status')} missing={sorted(missing_failed_codes)}")
        if "# Power Query Refresh Performance Report" not in clean_markdown:
            failures.append("clean markdown heading missing")
        if "Follow-Up Diagnostic Command" not in failed_markdown:
            failures.append("failed markdown diagnostic command missing")
        if failures:
            return CheckResult(name=name, status=FAIL, detail="; ".join(failures))

        return CheckResult(
            name=name,
            status=PASS,
            detail="clean, slow, and failed refresh timing/status reports verified",
            metadata={
                "cleanStatus": clean_report.get("status"),
                "slowElapsedStatus": slow_report.get("elapsedStatus"),
                "failedCodes": sorted(str(code) for code in failed_codes),
            },
        )


def power_query_live_refresh_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "Power Query live refresh fixture smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")

    script_dir = project_root / ".agents" / "skills" / "power-query-m-engineering" / "scripts"
    create_script = script_dir / "create_power_query_fixture_excel_com.ps1"
    manage_script = script_dir / "manage_power_queries_excel_com.ps1"
    refresh_script = script_dir / "refresh_power_queries_excel_com.ps1"
    export_script = script_dir / "export_power_queries_excel_com.ps1"
    missing = [str(path) for path in [create_script, manage_script, refresh_script, export_script] if not path.is_file()]
    if missing:
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {missing}")

    unavailable_markers = [
        "cannot create activex component",
        "class not registered",
        "excel.application",
        "microsoft.mashup.oledb",
        "provider cannot be found",
        "provider is not registered",
        "retrieving the com class factory",
    ]

    def single_or_list(value: object) -> list[object]:
        if value is None:
            return []
        if isinstance(value, list):
            return value
        return [value]

    with tempfile.TemporaryDirectory(prefix="excel_bi_pq_live_refresh_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "pq_live.xlsx"
        updated_workbook = tmp_dir / "pq_live_updated.xlsx"
        refreshed_workbook = tmp_dir / "pq_live_refreshed.xlsx"
        create_json = tmp_dir / "create.json"
        list_json = tmp_dir / "list.json"
        update_json = tmp_dir / "update.json"
        refresh_json = tmp_dir / "refresh.json"
        table_json = tmp_dir / "table.json"
        export_dir = tmp_dir / "exported"
        formula_path = tmp_dir / "SmokeQuery.updated.m"
        inspect_script = tmp_dir / "inspect_loaded_table.ps1"

        formula_path.write_text(
            'let Source = #table({"A","B"}, {{1,"x"},{2,"y"},{3,"z"}}) in Source\n',
            encoding="utf-8",
        )
        inspect_script.write_text(
            r'''
param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$OutJson
)
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($WorkbookPath, $null, $true)
    $sheet = $workbook.Worksheets.Item("PQ_Load")
    $table = $sheet.ListObjects.Item("SmokeQueryTable")
    $rows = 0
    $columns = [int]$table.Range.Columns.Count
    $values = @()
    if ($null -ne $table.DataBodyRange) {
        $rows = [int]$table.DataBodyRange.Rows.Count
        for ($r = 1; $r -le $rows; $r++) {
            $values += ,@(
                $table.DataBodyRange.Cells.Item($r, 1).Value2,
                $table.DataBodyRange.Cells.Item($r, 2).Value2
            )
        }
    }
    [ordered]@{
        workbookPath = $WorkbookPath
        sheetName = [string]$sheet.Name
        tableName = [string]$table.Name
        rows = $rows
        columns = $columns
        values = $values
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutJson -Encoding UTF8
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''.lstrip(),
            encoding="utf-8",
        )

        commands = [
            (
                "create",
                [
                    ps_exe,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(create_script),
                    "-OutputWorkbookPath",
                    str(workbook),
                    "-QueryName",
                    "SmokeQuery",
                    "-TableName",
                    "SmokeQueryTable",
                    "-OutJson",
                    str(create_json),
                ],
            ),
            (
                "list",
                [
                    ps_exe,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(manage_script),
                    "-WorkbookPath",
                    str(workbook),
                    "-Action",
                    "List",
                    "-OutJson",
                    str(list_json),
                ],
            ),
            (
                "update",
                [
                    ps_exe,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(manage_script),
                    "-WorkbookPath",
                    str(workbook),
                    "-Action",
                    "Update",
                    "-QueryName",
                    "SmokeQuery",
                    "-FormulaPath",
                    str(formula_path),
                    "-OutputWorkbookPath",
                    str(updated_workbook),
                    "-OutJson",
                    str(update_json),
                ],
            ),
            (
                "refresh",
                [
                    ps_exe,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(refresh_script),
                    "-WorkbookPath",
                    str(updated_workbook),
                    "-QueryName",
                    "SmokeQuery",
                    "-OutputWorkbookPath",
                    str(refreshed_workbook),
                    "-DisableBackgroundRefresh",
                    "-CalculateFull",
                    "-OutJson",
                    str(refresh_json),
                ],
            ),
            (
                "export",
                [
                    ps_exe,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(export_script),
                    "-WorkbookPath",
                    str(refreshed_workbook),
                    "-OutDir",
                    str(export_dir),
                ],
            ),
            (
                "inspect",
                [
                    ps_exe,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(inspect_script),
                    "-WorkbookPath",
                    str(refreshed_workbook),
                    "-OutJson",
                    str(table_json),
                ],
            ),
        ]

        command_results: dict[str, CheckResult] = {}
        for phase, command in commands:
            result = run_command(command, project_root, f"{name}: {phase}")
            command_results[phase] = result
            if result.status != PASS:
                combined = f"{result.stdout}\n{result.stderr}".lower()
                result.name = name
                if any(marker in combined for marker in unavailable_markers):
                    result.status = SKIP
                    result.detail = f"Power Query live refresh runtime unavailable during {phase}"
                return result

        try:
            create_report = json.loads(create_json.read_text(encoding="utf-8-sig"))
            list_report = json.loads(list_json.read_text(encoding="utf-8-sig"))
            update_report = json.loads(update_json.read_text(encoding="utf-8-sig"))
            refresh_report = json.loads(refresh_json.read_text(encoding="utf-8-sig"))
            export_report = json.loads(command_results["export"].stdout)
            table_report = json.loads(table_json.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read Power Query live fixture reports: {exc}")

        failures: list[str] = []
        if create_report.get("queryName") != "SmokeQuery":
            failures.append(f"create queryName={create_report.get('queryName')}")
        if create_report.get("rows") != 2 or create_report.get("columns") != 2:
            failures.append(f"create rows/columns={create_report.get('rows')}/{create_report.get('columns')}")

        listed = [item for item in single_or_list(list_report.get("after")) if isinstance(item, dict)]
        if not any(item.get("name") == "SmokeQuery" for item in listed):
            failures.append(f"list after={listed}")

        updated = [item for item in single_or_list(update_report.get("after")) if isinstance(item, dict)]
        updated_formula = ""
        for item in updated:
            if item.get("name") == "SmokeQuery":
                updated_formula = str(item.get("formula", ""))
                break
        if '{{1,"x"},{2,"y"},{3,"z"}}' not in updated_formula.replace(" ", ""):
            failures.append("updated formula does not contain expected 3-row table")
        if str(update_report.get("outputWorkbookPath", "")).lower() != str(updated_workbook).lower():
            failures.append("update output workbook did not match fixture")

        if refresh_report.get("queryName") != "SmokeQuery":
            failures.append(f"refresh queryName={refresh_report.get('queryName')}")
        if str(refresh_report.get("outputWorkbookPath", "")).lower() != str(refreshed_workbook).lower():
            failures.append("refresh output workbook did not match fixture")
        if refresh_report.get("errors"):
            failures.append(f"refresh errors={refresh_report.get('errors')}")
        if refresh_report.get("disableBackgroundRefresh") is not True:
            failures.append(f"disableBackgroundRefresh={refresh_report.get('disableBackgroundRefresh')}")
        if refresh_report.get("calculateFull") is not True:
            failures.append(f"calculateFull={refresh_report.get('calculateFull')}")
        after_connections = [item for item in single_or_list(refresh_report.get("afterConnections")) if isinstance(item, dict)]
        if not after_connections:
            failures.append("refresh afterConnections missing")
        if any(item.get("refreshing") is True for item in after_connections):
            failures.append(f"still refreshing={after_connections}")

        exported_queries = [item for item in single_or_list(export_report.get("queries")) if isinstance(item, dict)]
        if export_report.get("queryCount") != 1:
            failures.append(f"export queryCount={export_report.get('queryCount')}")
        if not any(item.get("name") == "SmokeQuery" and "3,\"z\"" in str(item.get("formula", "")) for item in exported_queries):
            failures.append(f"exported queries={exported_queries}")

        values = table_report.get("values", [])
        if table_report.get("sheetName") != "PQ_Load" or table_report.get("tableName") != "SmokeQueryTable":
            failures.append(f"loaded table={table_report.get('sheetName')}/{table_report.get('tableName')}")
        if table_report.get("rows") != 3 or table_report.get("columns") != 2:
            failures.append(f"loaded rows/columns={table_report.get('rows')}/{table_report.get('columns')}")
        if [3, "z"] not in values and [3.0, "z"] not in values:
            failures.append(f"loaded values={values}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="Power Query create, list, update, refresh wait, export, and loaded-table result verified"
            if not failures
            else "; ".join(failures),
            metadata={
                "createdRows": create_report.get("rows"),
                "loadedRows": table_report.get("rows"),
                "refreshElapsedSeconds": refresh_report.get("elapsedSeconds"),
                "queryCount": export_report.get("queryCount"),
                "connectionCount": len(after_connections),
            },
        )


def excel_workbook_com_inventory_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "Excel workbook COM inventory fixture smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")

    inspector = project_root / ".agents" / "skills" / "excel-vba-workbook-engineering" / "scripts" / "inspect_workbook.ps1"
    if not inspector.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {inspector}")

    unavailable_markers = [
        "cannot create activex component",
        "class not registered",
        "excel.application",
        "retrieving the com class factory",
    ]

    with tempfile.TemporaryDirectory(prefix="excel_bi_workbook_inventory_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "inspect_probe.xlsx"
        out_json = tmp_dir / "inspect_probe.json"
        create_script = tmp_dir / "create_inspect_probe.ps1"
        create_script.write_text(
            r'''
param([Parameter(Mandatory = $true)][string]$WorkbookPath)
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "Data"
    $sheet.Range("A1").Value2 = "Channel"
    $sheet.Range("B1").Value2 = "Budget"
    $sheet.Range("C1").Value2 = "Score"
    $sheet.Range("A2").Value2 = "A"
    $sheet.Range("B2").Value2 = 10
    $sheet.Range("C2").Formula = "=B2*2"
    $sheet.Range("A3").Value2 = "B"
    $sheet.Range("B3").Value2 = 20
    $sheet.Range("C3").Formula = "=B3*2"
    $sheet.Range("D1").Value2 = "Total"
    $sheet.Range("D2").Formula = "=SUM(C2:C3)"
    [void]$workbook.Names.Add("TotalAmount", "=Data!`$D`$2")
    $shape = $sheet.Shapes.AddShape(1, 250, 40, 120, 30)
    $shape.Name = "btnRunProbe"
    $shape.TextFrame2.TextRange.Text = "Run Probe"
    $shape.OnAction = "RunProbe"
    $workbook.SaveAs($WorkbookPath, 51)
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''.lstrip(),
            encoding="utf-8",
        )

        create_result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(create_script),
                "-WorkbookPath",
                str(workbook),
            ],
            project_root,
            name,
        )
        if create_result.status != PASS:
            combined = f"{create_result.stdout}\n{create_result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                create_result.status = SKIP
                create_result.detail = "Excel COM runtime unavailable while creating fixture workbook"
            return create_result

        inspect_result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(inspector),
                "-WorkbookPath",
                str(workbook),
                "-OutJson",
                str(out_json),
            ],
            project_root,
            name,
        )
        if inspect_result.status != PASS:
            combined = f"{inspect_result.stdout}\n{inspect_result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                inspect_result.status = SKIP
                inspect_result.detail = "Excel COM runtime unavailable while inspecting fixture workbook"
            return inspect_result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read workbook inventory report: {exc}")

        failures: list[str] = []
        if not str(report.get("workbookPath", "")).lower().endswith("inspect_probe.xlsx"):
            failures.append("workbookPath did not point to fixture workbook")
        if report.get("name") != "inspect_probe.xlsx":
            failures.append(f"name={report.get('name')}")
        if report.get("fileFormat") != 51:
            failures.append(f"fileFormat={report.get('fileFormat')}")
        if report.get("hasVBProject") is not False:
            failures.append(f"hasVBProject={report.get('hasVBProject')}")

        worksheets = report.get("worksheets", [])
        if not isinstance(worksheets, list):
            failures.append("worksheets is not a list")
            worksheets = []
        sheet_map = {item.get("name"): item for item in worksheets if isinstance(item, dict)}
        data_sheet = sheet_map.get("Data")
        if not isinstance(data_sheet, dict):
            failures.append("Data sheet missing")
            data_sheet = {}
        if isinstance(data_sheet, dict):
            if data_sheet.get("visible") != "Visible":
                failures.append(f"Data.visible={data_sheet.get('visible')}")
            if int(data_sheet.get("rows", 0) or 0) < 3:
                failures.append(f"Data.rows={data_sheet.get('rows')}")
            if int(data_sheet.get("columns", 0) or 0) < 4:
                failures.append(f"Data.columns={data_sheet.get('columns')}")
            if int(data_sheet.get("formulaCount", 0) or 0) < 1:
                failures.append(f"Data.formulaCount={data_sheet.get('formulaCount')}")
            shapes = data_sheet.get("shapes", [])
            if not isinstance(shapes, list):
                failures.append("Data.shapes is not a list")
                shapes = []
            shape_map = {shape.get("name"): shape for shape in shapes if isinstance(shape, dict)}
            button = shape_map.get("btnRunProbe")
            if not isinstance(button, dict):
                failures.append("btnRunProbe shape missing")
            else:
                if button.get("onAction") != "RunProbe":
                    failures.append(f"btnRunProbe.onAction={button.get('onAction')}")
                if button.get("text") != "Run Probe":
                    failures.append(f"btnRunProbe.text={button.get('text')}")

        names = report.get("names", [])
        if not isinstance(names, list):
            failures.append("names is not a list")
            names = []
        name_map = {item.get("name"): item for item in names if isinstance(item, dict)}
        total_name = name_map.get("TotalAmount")
        if not isinstance(total_name, dict):
            failures.append("TotalAmount name missing")
        else:
            refers_to = str(total_name.get("refersTo", "")).replace("'", "")
            if refers_to != "=Data!$D$2":
                failures.append(f"TotalAmount.refersTo={total_name.get('refersTo')}")
            if total_name.get("visible") is not True:
                failures.append(f"TotalAmount.visible={total_name.get('visible')}")

        for key in ["links", "connections", "queries"]:
            if not isinstance(report.get(key), list):
                failures.append(f"{key} is not a list")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="Excel COM inspector returned expected workbook, sheet, formula, name, and button inventory"
            if not failures
            else "; ".join(failures),
            metadata={
                "workbookName": report.get("name"),
                "fileFormat": report.get("fileFormat"),
                "sheetNames": sorted(str(item.get("name")) for item in worksheets if isinstance(item, dict)),
                "formulaCount": data_sheet.get("formulaCount") if isinstance(data_sheet, dict) else None,
                "shapeNames": sorted(str(item.get("name")) for item in data_sheet.get("shapes", []) if isinstance(item, dict))
                if isinstance(data_sheet, dict)
                else [],
                "nameCount": len(names),
                "queryAccessError": report.get("queryAccessError"),
                "vbaAccessError": report.get("vbaAccessError"),
            },
        )


def provider_probe_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "Provider probe fixture smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")
    script = project_root / "tools" / "probe_excel_bi_providers.ps1"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_provider_probe_") as tmp:
        tmp_dir = Path(tmp)
        out_json = tmp_dir / "provider_probe.json"
        smoke_workbook = tmp_dir / "provider probe ado smoke.xlsx"
        result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-RunExcelComSmoke",
                "-RunAdoWorkbookSmoke",
                "-SmokeWorkbookPath",
                str(smoke_workbook),
                "-OutJson",
                str(out_json),
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            unavailable_markers = [
                "cannot create activex component",
                "class not registered",
                "provider cannot be found",
                "provider is not registered",
                "microsoft.ace.oledb",
                "excel.application",
            ]
            combined = f"{result.stdout}\n{result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                result.status = SKIP
                result.detail = "Excel COM, ACE OLEDB, or provider runtime unavailable"
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read provider probe report: {exc}")

        failures: list[str] = []
        providers = report.get("providers", [])
        com_prog_ids = report.get("comProgIds", [])
        if not isinstance(providers, list):
            failures.append("providers is not a list")
            providers = []
        if not isinstance(com_prog_ids, list):
            failures.append("comProgIds is not a list")
            com_prog_ids = []
        provider_map = {item.get("provider"): item for item in providers if isinstance(item, dict)}
        com_map = {item.get("progId"): item for item in com_prog_ids if isinstance(item, dict)}

        expected_providers = ["Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0", "MSOLAP"]
        missing_provider_rows = [item for item in expected_providers if item not in provider_map]
        if missing_provider_rows:
            failures.append(f"missing provider rows={missing_provider_rows}")

        expected_com = ["Excel.Application", "ADODB.Connection", "ADODB.Recordset", "ADOMD.Catalog", "ADOMD.Cellset"]
        missing_com_rows = [item for item in expected_com if item not in com_map]
        if missing_com_rows:
            failures.append(f"missing COM rows={missing_com_rows}")

        unavailable: list[str] = []
        for prog_id in ["Excel.Application", "ADODB.Connection"]:
            item = com_map.get(prog_id, {})
            if isinstance(item, dict) and item.get("creatable") is not True:
                unavailable.append(f"{prog_id}: {item.get('error', '')}")

        machine = report.get("machine", {})
        if not isinstance(machine, dict):
            failures.append("machine is not an object")
            machine = {}
        for key in ["is64BitOperatingSystem", "is64BitProcess", "powershellVersion"]:
            if key not in machine:
                failures.append(f"machine.{key} missing")

        excel_smoke = report.get("excelComSmoke")
        ado_smoke = report.get("adoWorkbookSmoke")
        if not isinstance(excel_smoke, dict):
            failures.append("excelComSmoke missing")
            excel_smoke = {}
        if not isinstance(ado_smoke, dict):
            failures.append("adoWorkbookSmoke missing")
            ado_smoke = {}
        if isinstance(excel_smoke, dict) and excel_smoke.get("succeeded") is not True:
            unavailable.append(f"Excel COM smoke: {excel_smoke.get('error', '')}")
        if isinstance(ado_smoke, dict) and ado_smoke.get("succeeded") is not True:
            unavailable.append(f"ADO workbook smoke: {ado_smoke.get('error', '')}")

        if isinstance(ado_smoke, dict) and ado_smoke.get("succeeded") is True:
            if ado_smoke.get("rowCount") != 2:
                failures.append(f"adoWorkbookSmoke.rowCount={ado_smoke.get('rowCount')}")
            if ado_smoke.get("fields") != ["Category", "TotalAmount"]:
                failures.append(f"adoWorkbookSmoke.fields={ado_smoke.get('fields')}")
            if int(ado_smoke.get("schemaTableCount", 0) or 0) < 1:
                failures.append(f"adoWorkbookSmoke.schemaTableCount={ado_smoke.get('schemaTableCount')}")

        interpretation = [str(item) for item in report.get("interpretation", [])]
        if not any("registryDetected means" in item for item in interpretation):
            failures.append("missing registryDetected interpretation")
        if not any("ADOMD/MSOLAP availability" in item for item in interpretation):
            failures.append("missing ADOMD/MSOLAP boundary interpretation")

        metadata = {
            "machine": machine,
            "providerRegistryDetected": {
                name: provider_map.get(name, {}).get("registryDetected") for name in expected_providers
            },
            "comCreatable": {name: com_map.get(name, {}).get("creatable") for name in expected_com},
            "excelComSmoke": {
                "succeeded": excel_smoke.get("succeeded"),
                "version": excel_smoke.get("version"),
                "build": excel_smoke.get("build"),
            },
            "adoWorkbookSmoke": {
                "succeeded": ado_smoke.get("succeeded"),
                "provider": ado_smoke.get("provider"),
                "rowCount": ado_smoke.get("rowCount"),
                "fields": ado_smoke.get("fields"),
            },
        }
        if unavailable:
            return CheckResult(
                name=name,
                status=SKIP,
                detail="Excel COM, ACE OLEDB, or provider runtime unavailable: " + "; ".join(unavailable),
                metadata=metadata,
            )

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="provider registry, COM activation, Excel COM smoke, and ACE ADO workbook smoke verified" if not failures else "; ".join(failures),
            metadata=metadata,
        )


def provider_environment_report_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "Provider environment report"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")
    report_script = project_root / "tools" / "build_provider_environment_report.py"
    if not report_script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {report_script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_provider_env_report_") as tmp:
        tmp_dir = Path(tmp)
        out_json = tmp_dir / "provider_environment.json"
        out_md = tmp_dir / "provider_environment.md"
        result = run_command(
            [
                sys.executable,
                str(report_script),
                "--project-root",
                str(project_root),
                "--run-probe",
                "--powershell",
                ps_exe,
                "--excel-com",
                "--ado-workbook-smoke",
                "--out-json",
                str(out_json),
                "--out-md",
                str(out_md),
                "--require-pass",
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            unavailable_markers = [
                "cannot create activex component",
                "class not registered",
                "provider cannot be found",
                "provider is not registered",
                "microsoft.ace.oledb",
                "excel.application",
            ]
            combined = f"{result.stdout}\n{result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                result.status = SKIP
                result.detail = "Excel COM, ACE OLEDB, or provider runtime unavailable"
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8-sig"))
            markdown = out_md.read_text(encoding="utf-8")
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read provider environment report: {exc}")

        failures: list[str] = []
        if report.get("status") != PASS:
            failures.append(f"status={report.get('status')}")
        summary = report.get("summary", {})
        if not isinstance(summary, dict):
            failures.append("summary missing")
            summary = {}
        readiness = summary.get("readiness", {})
        if not isinstance(readiness, dict):
            failures.append("readiness missing")
            readiness = {}
        for key in ["excelAutomationReady", "adodbReady", "workbookSqlReady", "aceProviderRegistered"]:
            if readiness.get(key) is not True:
                failures.append(f"{key}={readiness.get(key)}")
        provider_status = summary.get("providerStatus", {})
        com_status = summary.get("comStatus", {})
        if not isinstance(provider_status, dict):
            failures.append("providerStatus missing")
            provider_status = {}
        if not isinstance(com_status, dict):
            failures.append("comStatus missing")
            com_status = {}
        for key in ["Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0", "MSOLAP"]:
            if key not in provider_status:
                failures.append(f"providerStatus.{key} missing")
        for key in ["Excel.Application", "ADODB.Connection", "ADOMD.Catalog", "ADOMD.Cellset"]:
            if key not in com_status:
                failures.append(f"comStatus.{key} missing")
        if "Provider Environment Report" not in markdown:
            failures.append("Markdown title missing")
        if "status: **pass**" not in markdown:
            failures.append("Markdown status missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="Excel COM, ACE workbook SQL, MSOLAP, and ADOMD capability matrix reported" if not failures else "; ".join(failures),
            stdout=result.stdout,
            stderr=result.stderr,
            metadata={
                "readiness": readiness,
                "excelComSmoke": summary.get("excelComSmoke"),
                "adoWorkbookSmoke": summary.get("adoWorkbookSmoke"),
            },
        )


def provider_environment_baseline_fixture_check(project_root: Path) -> CheckResult:
    name = "Provider environment baseline drift fixture smoke"
    fixture_script = project_root / "tools" / "create_provider_environment_fixture.py"
    report_script = project_root / "tools" / "build_provider_environment_report.py"
    for script in [fixture_script, report_script]:
        if not script.is_file():
            return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_provider_baseline_") as tmp:
        tmp_dir = Path(tmp)
        fixture_dir = tmp_dir / "fixture"
        manifest_json = tmp_dir / "fixture_manifest.json"
        matching_report_json = tmp_dir / "matching_report.json"
        matching_report_md = tmp_dir / "matching_report.md"
        drift_report_json = tmp_dir / "drift_report.json"
        drift_report_md = tmp_dir / "drift_report.md"

        create_result = run_command(
            [
                sys.executable,
                str(fixture_script),
                "--out-dir",
                str(fixture_dir),
                "--out-json",
                str(manifest_json),
            ],
            project_root,
            name,
        )
        if create_result.status != PASS:
            return create_result

        try:
            manifest = json.loads(manifest_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read fixture manifest: {exc}")

        current_probe = manifest.get("currentProbe", "")
        matching_baseline = manifest.get("matchingBaseline", "")
        drifting_baseline = manifest.get("driftingBaseline", "")
        matching_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--project-root",
                str(project_root),
                "--probe-json",
                str(current_probe),
                "--baseline-json",
                str(matching_baseline),
                "--fail-on-drift",
                "--out-json",
                str(matching_report_json),
                "--out-md",
                str(matching_report_md),
                "--require-pass",
            ],
            project_root,
            f"{name}: matching baseline",
        )
        if matching_result.status != PASS:
            return CheckResult(
                name=name,
                status=FAIL,
                detail=f"matching baseline failed: {matching_result.detail}",
                command=matching_result.command,
                stdout=matching_result.stdout,
                stderr=matching_result.stderr,
            )

        drift_result = run_command(
            [
                sys.executable,
                str(report_script),
                "--project-root",
                str(project_root),
                "--probe-json",
                str(current_probe),
                "--baseline-json",
                str(drifting_baseline),
                "--fail-on-drift",
                "--out-json",
                str(drift_report_json),
                "--out-md",
                str(drift_report_md),
                "--require-pass",
            ],
            project_root,
            f"{name}: drifting baseline",
            ok_codes={1},
        )
        if drift_result.status != PASS:
            return CheckResult(
                name=name,
                status=FAIL,
                detail=f"drifting baseline did not produce expected failure: {drift_result.detail}",
                command=drift_result.command,
                stdout=drift_result.stdout,
                stderr=drift_result.stderr,
            )

        try:
            matching_report = json.loads(matching_report_json.read_text(encoding="utf-8"))
            drift_report = json.loads(drift_report_json.read_text(encoding="utf-8"))
            drift_markdown = drift_report_md.read_text(encoding="utf-8")
        except (OSError, json.JSONDecodeError) as exc:
            return CheckResult(name=name, status=FAIL, detail=f"cannot read provider baseline reports: {exc}")

        failures: list[str] = []
        matching_comparison = matching_report.get("comparison", {})
        if matching_report.get("status") != PASS:
            failures.append(f"matching status={matching_report.get('status')}")
        if matching_comparison.get("changedCount") != 0:
            failures.append(f"matching changedCount={matching_comparison.get('changedCount')}")

        drift_comparison = drift_report.get("comparison", {})
        drift_changes = drift_comparison.get("changes", [])
        drift_paths = {item.get("path") for item in drift_changes if isinstance(item, dict)}
        expected = manifest.get("expected", {}) if isinstance(manifest.get("expected", {}), dict) else {}
        minimum_changed = int(expected.get("driftingMinimumChangedCount", 1))
        if drift_report.get("status") != FAIL:
            failures.append(f"drift status={drift_report.get('status')}")
        if drift_comparison.get("changedCount", 0) < minimum_changed:
            failures.append(f"drift changedCount={drift_comparison.get('changedCount')}")
        required_paths = {
            "providerStatus.MSOLAP",
            "comStatus.ADOMD.Cellset",
            "readiness.workbookSqlReady",
            "readiness.msolapRegistered",
            "readiness.adomdComReady",
            "adoWorkbookSmoke.succeeded",
        }
        missing_paths = sorted(required_paths - drift_paths)
        if missing_paths:
            failures.append(f"missing drift paths={missing_paths}")
        if not any("provider environment drift detected" in str(error) for error in drift_report.get("errors", [])):
            failures.append("missing fail-on-drift error")
        if "Changed fields" not in drift_markdown:
            failures.append("Markdown baseline comparison missing")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="matching baseline passes and drifting baseline fails with expected change paths"
            if not failures
            else "; ".join(failures),
            stdout="\n".join(part for part in [create_result.stdout, matching_result.stdout, drift_result.stdout] if part),
            stderr="\n".join(part for part in [create_result.stderr, matching_result.stderr, drift_result.stderr] if part),
            metadata={
                "matchingChangedCount": matching_comparison.get("changedCount"),
                "driftChangedCount": drift_comparison.get("changedCount"),
                "driftPaths": sorted(str(item) for item in drift_paths),
            },
        )


def ado_workbook_sql_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "ADO workbook SQL fixture smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")
    script = project_root / "tools" / "test_excel_ado_sql_access.ps1"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_ado_sql_") as tmp:
        tmp_dir = Path(tmp)
        workbook = tmp_dir / "ado sql fixture.xlsx"
        out_json = tmp_dir / "ado_sql_report.json"
        result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-WorkbookPath",
                str(workbook),
                "-CreateFixture",
                "-SqlText",
                "SELECT * FROM [Data$]",
                "-IncludeSchema",
                "-OutJson",
                str(out_json),
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            provider_missing_markers = [
                "provider cannot be found",
                "provider is not registered",
                "class not registered",
                "microsoft.ace.oledb",
            ]
            combined = f"{result.stdout}\n{result.stderr}".lower()
            if any(marker in combined for marker in provider_missing_markers):
                result.status = SKIP
                result.detail = "ACE OLEDB provider or Excel COM runtime unavailable"
            return result
        try:
            report = json.loads(out_json.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read ADO report: {exc}")

        failures: list[str] = []
        if report.get("succeeded") is not True:
            failures.append(f"succeeded={report.get('succeeded')}")
        if report.get("rowCount") != 4:
            failures.append(f"rowCount={report.get('rowCount')}")
        field_names = [field.get("name") for field in report.get("fields", [])]
        if field_names != ["Region", "Category", "Amount", "Period"]:
            failures.append(f"fields={field_names}")
        schema_names = {item.get("tableName") for item in report.get("schemaTables", []) if isinstance(item, dict)}
        if "Data$" not in schema_names:
            failures.append(f"schemaTables={sorted(str(item) for item in schema_names)}")
        rows = report.get("rows", [])
        total_amount = sum(float(row.get("Amount", 0)) for row in rows if isinstance(row, dict))
        if total_amount != 500:
            failures.append(f"amountTotal={total_amount}")
        if not str(report.get("workbookPath", "")).lower().endswith(".xlsx"):
            failures.append("workbookPath did not point to .xlsx fixture")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="ACE OLEDB workbook SQL returned expected rows, fields, schema, and numeric total" if not failures else "; ".join(failures),
            metadata={
                "provider": report.get("provider"),
                "rowCount": report.get("rowCount"),
                "fields": field_names,
                "amountTotal": total_amount,
            },
        )


def adomd_com_probe_fixture_check(project_root: Path, ps_exe: str | None) -> CheckResult:
    name = "ADOMD COM probe fixture smoke"
    if platform.system().lower() != "windows":
        return CheckResult(name=name, status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name=name, status=SKIP, detail="PowerShell not found")
    script = project_root / "tools" / "test_excel_adomd_query.ps1"
    if not script.is_file():
        return CheckResult(name=name, status=FAIL, detail=f"script not found: {script}")

    with tempfile.TemporaryDirectory(prefix="excel_bi_adomd_probe_") as tmp:
        out_json = Path(tmp) / "adomd_probe.json"
        result = run_command(
            [
                ps_exe,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-ProbeOnly",
                "-OutJson",
                str(out_json),
            ],
            project_root,
            name,
        )
        if result.status != PASS:
            unavailable_markers = [
                "cannot create activex component",
                "class not registered",
                "adomd",
                "msolap",
                "provider is not registered",
            ]
            combined = f"{result.stdout}\n{result.stderr}".lower()
            if any(marker in combined for marker in unavailable_markers):
                result.status = SKIP
                result.detail = "ADOMD/MSOLAP runtime unavailable"
            return result

        try:
            report = json.loads(out_json.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            return CheckResult(name=name, status=FAIL, detail=f"could not read ADOMD probe report: {exc}")

        probe = report.get("probe", {})
        if not isinstance(probe, dict):
            return CheckResult(name=name, status=FAIL, detail="probe section missing or not an object")

        expected = {
            "adodbConnection": "ADODB.Connection",
            "adomdCatalog": "ADOMD.Catalog",
            "adomdCellset": "ADOMD.Cellset",
        }
        failures: list[str] = []
        unavailable: list[str] = []
        for key, prog_id in expected.items():
            item = probe.get(key, {})
            if not isinstance(item, dict):
                failures.append(f"{key} missing")
                continue
            if item.get("progId") != prog_id:
                failures.append(f"{key}.progId={item.get('progId')}")
            if item.get("creatable") is not True:
                unavailable.append(f"{prog_id}: {item.get('error', '')}")

        limitations = [str(item) for item in report.get("limitations", [])]
        if report.get("mode") != "ProbeOnly":
            failures.append(f"mode={report.get('mode')}")
        if not any("COM activation only" in item for item in limitations):
            failures.append("missing COM activation limitation")
        if not any("connection string and MDX query" in item for item in limitations):
            failures.append("missing real-query limitation")

        metadata = {
            "mode": report.get("mode"),
            "succeeded": report.get("succeeded"),
            "probe": probe,
            "limitations": limitations,
        }
        if unavailable:
            return CheckResult(
                name=name,
                status=SKIP,
                detail="ADOMD/MSOLAP runtime unavailable: " + "; ".join(unavailable),
                metadata=metadata,
            )
        if report.get("succeeded") is not True:
            failures.append(f"succeeded={report.get('succeeded')}")

        return CheckResult(
            name=name,
            status=PASS if not failures else FAIL,
            detail="ADODB and ADOMD COM activation verified; endpoint query remains explicit-user-input only" if not failures else "; ".join(failures),
            metadata=metadata,
        )


def excel_process_check(strict: bool, ps_exe: str | None, project_root: Path) -> CheckResult:
    if platform.system().lower() != "windows":
        return CheckResult(name="Excel process check", status=SKIP, detail="not Windows")
    if not ps_exe:
        return CheckResult(name="Excel process check", status=SKIP, detail="PowerShell not found")
    command = [ps_exe, "-NoProfile", "-Command", "Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id,ProcessName,MainWindowTitle | ConvertTo-Json -Depth 3"]
    result = run_command(command, project_root, "Excel process check", ok_codes={0})
    has_output = bool(result.stdout.strip())
    if has_output:
        result.status = FAIL if strict else WARN
        result.detail = "EXCEL process detected"
    else:
        result.status = PASS
        result.detail = "no EXCEL process detected"
    return result


def write_json(path: Path, data: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def markdown_report(data: dict[str, object]) -> str:
    checks = data["checks"]
    assert isinstance(checks, list)
    lines = [
        "# Release Gate Report",
        "",
        f"- Project: `{data['projectRoot']}`",
        f"- Version: `{data.get('version', '')}`",
        f"- Profile: `{data.get('profile', '')}`",
        f"- Overall: **{data['overallStatus']}**",
        f"- Generated: `{data['generatedAt']}`",
        "",
        "| Check | Status | Detail |",
        "|---|---:|---|",
    ]
    for check in checks:
        assert isinstance(check, dict)
        detail = str(check.get("detail", "")).replace("|", "\\|")
        lines.append(f"| {check.get('name', '')} | {check.get('status', '')} | {detail} |")
    lines.append("")
    lines.append("Statuses: `pass` and `skip` are non-failing; `warn` indicates manual review; `fail` blocks release.")
    return "\n".join(lines) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", help="Plugin project root")
    parser.add_argument("--plugin-validator", default="", help="Path to plugin-creator validate_plugin.py")
    parser.add_argument("--local-plugin", default="", help="Optional local plugin copy to validate")
    parser.add_argument("--cache-plugin", default="", help="Optional installed cache plugin copy to validate")
    parser.add_argument("--out-json", default="", help="Write JSON report")
    parser.add_argument("--out-md", default="", help="Write Markdown report")
    parser.add_argument(
        "--profile",
        choices=["full", "structural"],
        default="full",
        help="Use 'structural' for cross-platform package/OpenXML validation without Excel runtime or installed-plugin checks",
    )
    parser.add_argument("--strict-excel-process", action="store_true", help="Fail if any EXCEL.EXE process is running")
    parser.add_argument("--no-default-sensitive-markers", action="store_true", help="Disable built-in legacy customer marker scan")
    parser.add_argument("--sensitive-marker", action="append", default=[], help="Additional marker that must not appear in the package")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).expanduser().resolve()
    plugin = read_plugin_json(project_root)
    plugin_name = str(plugin.get("name", project_root.name))
    version = str(plugin.get("version", ""))

    validator = Path(args.plugin_validator).expanduser().resolve() if args.plugin_validator else default_plugin_validator()
    local_plugin = Path(args.local_plugin).expanduser().resolve() if args.local_plugin else Path.home() / "plugins" / plugin_name
    cache_plugin = (
        Path(args.cache_plugin).expanduser().resolve()
        if args.cache_plugin
        else Path.home() / ".codex" / "plugins" / "cache" / "personal" / plugin_name / version
    )

    checks: list[CheckResult] = []
    checks.append(run_command([sys.executable, str(project_root / "tools" / "validate-skills.py"), str(project_root)], project_root, "source skill validation"))
    if validator.is_file():
        checks.append(run_command([sys.executable, str(validator), str(project_root)], project_root, "source plugin validation"))
    else:
        checks.append(CheckResult(name="source plugin validation", status=SKIP, detail=f"validator not found: {validator}"))

    checks.append(python_compile_check(project_root))
    ps_exe = find_powershell()
    bash_exe = find_bash()
    checks.append(powershell_parse_check(project_root, ps_exe))
    checks.append(bash_syntax_check(project_root, bash_exe))
    checks.append(portable_structural_wrapper_fixture_check(project_root, bash_exe))
    checks.append(run_command([sys.executable, str(project_root / "tools" / "validate_project_docs.py"), "--project-root", str(project_root)], project_root, "project documentation consistency validation"))
    checks.append(run_command([sys.executable, str(project_root / "tools" / "validate_github_community_health.py"), "--project-root", str(project_root)], project_root, "GitHub community health validation"))
    checks.append(run_command([sys.executable, str(project_root / "tools" / "validate_task_recipes.py"), "--project-root", str(project_root)], project_root, "task recipe documentation validation"))
    checks.append(run_command([sys.executable, str(project_root / "tools" / "validate_official_docs_index.py"), "--project-root", str(project_root)], project_root, "official documentation index validation"))
    checks.append(official_docs_drift_report_check(project_root))
    checks.append(run_command([sys.executable, str(project_root / "tools" / "sync-skills.py"), "--project-root", str(project_root), "--all-project-mirrors", "--check-drift"], project_root, "cross-agent mirror drift"))
    checks.append(cross_agent_forward_test_pack_check(project_root))
    checks.append(cross_agent_forward_test_result_scorer_check(project_root))
    checks.append(cross_agent_forward_test_runbook_check(project_root))
    checks.append(cross_agent_forward_test_handoff_bundle_check(project_root))
    checks.append(cross_agent_response_collection_report_check(project_root))
    checks.append(excel_bi_router_fixture_check(project_root))
    checks.append(capability_catalog_fixture_check(project_root))
    checks.append(agent_bootstrap_bundle_fixture_check(project_root))
    checks.append(goal_coverage_report_check(project_root))
    checks.append(completion_readiness_audit_check(project_root))
    checks.append(release_evidence_bundle_check(project_root))
    checks.append(real_sanitized_case_regression_check(project_root))
    checks.append(sanitized_fixture_bundle_check(project_root))
    checks.append(cube_formula_fixture_check(project_root))
    checks.append(power_pivot_model_report_fixture_check(project_root))
    checks.append(external_dependency_fixture_check(project_root))
    checks.append(workbook_surface_fixture_check(project_root))
    checks.append(visual_qa_report_fixture_check(project_root))
    checks.append(formula_quality_report_fixture_check(project_root))
    checks.append(workbook_controls_report_fixture_check(project_root))
    checks.append(external_dependency_report_fixture_check(project_root))
    checks.append(workbook_triage_report_fixture_check(project_root))
    checks.append(pure_deliverable_cleanup_plan_fixture_check(project_root))
    checks.append(pure_deliverable_verification_report_fixture_check(project_root))
    checks.append(measure_rename_impact_fixture_check(project_root))
    checks.append(measure_rename_rewrite_plan_fixture_check(project_root))
    checks.append(measure_delete_rewrite_plan_fixture_check(project_root))
    checks.append(escaped_mdx_measure_reference_fixture_check(project_root))
    checks.append(vba_source_lint_fixture_check(project_root))
    checks.append(vba_button_binding_report_fixture_check(project_root))
    checks.append(dax_compat_lint_fixture_check(project_root))
    checks.append(dax_dependency_fixture_check(project_root))
    checks.append(power_query_m_lint_fixture_check(project_root))
    checks.append(power_query_lineage_report_fixture_check(project_root))
    checks.append(power_query_refresh_error_classifier_fixture_check(project_root))
    checks.append(power_query_refresh_performance_report_fixture_check(project_root))
    checks.append(provider_environment_baseline_fixture_check(project_root))
    checks.append(scan_regex(project_root, PLACEHOLDER_RE, "placeholder marker scan"))

    markers = [] if args.no_default_sensitive_markers else list(DEFAULT_SENSITIVE_MARKERS)
    markers.extend(args.sensitive_marker)
    checks.append(scan_sensitive_markers(project_root, markers))
    checks.append(artifact_hygiene_report_check(project_root))
    checks.append(pycache_check(project_root))

    if args.profile == "full":
        checks.append(power_query_live_refresh_fixture_check(project_root, ps_exe))
        checks.append(excel_workbook_com_inventory_fixture_check(project_root, ps_exe))
        checks.append(visual_qa_render_evidence_fixture_check(project_root, ps_exe))
        checks.append(vba_import_export_run_fixture_check(project_root, ps_exe))
        checks.append(provider_probe_fixture_check(project_root, ps_exe))
        checks.append(provider_environment_report_check(project_root, ps_exe))
        checks.append(ado_workbook_sql_fixture_check(project_root, ps_exe))
        checks.append(adomd_com_probe_fixture_check(project_root, ps_exe))
        checks.append(excel_process_check(args.strict_excel_process, ps_exe, project_root))

        if validator.is_file() and local_plugin.exists():
            checks.append(run_command([sys.executable, str(validator), str(local_plugin)], project_root, "local plugin copy validation"))
        else:
            checks.append(CheckResult(name="local plugin copy validation", status=SKIP, detail=f"not found or validator missing: {local_plugin}"))

        if validator.is_file() and cache_plugin.exists():
            checks.append(run_command([sys.executable, str(validator), str(cache_plugin)], project_root, "installed cache validation"))
        else:
            checks.append(CheckResult(name="installed cache validation", status=SKIP, detail=f"not found or validator missing: {cache_plugin}"))

        codex = find_codex()
        if codex:
            result = run_command([codex, "plugin", "list"], project_root, "codex plugin list")
            if result.status == PASS:
                if plugin_name in result.stdout and (not version or version in result.stdout):
                    result.detail = "installed plugin entry found"
                else:
                    result.status = WARN
                    result.detail = "plugin/version not found in codex plugin list output"
            checks.append(result)
        else:
            checks.append(CheckResult(name="codex plugin list", status=SKIP, detail="codex command not found"))
    else:
        checks.append(CheckResult(name="Excel process check", status=SKIP, detail="skipped by structural profile"))
        checks.append(CheckResult(name="local plugin copy validation", status=SKIP, detail="skipped by structural profile"))
        checks.append(CheckResult(name="installed cache validation", status=SKIP, detail="skipped by structural profile"))
        checks.append(CheckResult(name="codex plugin list", status=SKIP, detail="skipped by structural profile"))

    failing = [check for check in checks if check.status == FAIL]
    warning = [check for check in checks if check.status == WARN]
    overall = FAIL if failing else (WARN if warning else PASS)

    report = {
        "generatedAt": now_iso(),
        "projectRoot": str(project_root),
        "pluginName": plugin_name,
        "version": version,
        "profile": args.profile,
        "overallStatus": overall,
        "checks": [check.__dict__ for check in checks],
    }

    if args.out_json:
        write_json(Path(args.out_json).expanduser().resolve(), report)
    if args.out_md:
        out_md = Path(args.out_md).expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_report(report), encoding="utf-8")

    print(f"Release gate profile: {args.profile}")
    print(f"Release gate overall: {overall}")
    for check in checks:
        print(f"- {check.status}: {check.name} ({check.detail})")
    return 1 if overall == FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
