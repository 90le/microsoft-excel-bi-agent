from __future__ import annotations

import importlib.util
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VERIFIER = ROOT / "tools" / "verify_plugin_benchmark_output.py"
BENCHMARK_CONFIG = ROOT / "benchmarks" / "plugin-eval-v0.2.1.json"
SOURCE_ARTIFACT = Path("benchmarks/fixtures/excel-bi-benchmark-source.json")
DAX_CHANGE_ALLOWLIST = Path("benchmarks/fixtures/dax-workspace-change-allowlist.json")
SOURCE_SHA256 = "1610e349e03bcb5e7bfab93c84f25c5add842fd35b01884d319c365a3138373b"
DAX_ALLOWLIST_SHA256 = "3372305a7de1948289a972cda845fb1ee0b9d335ad1b7d5dd65025d8458dee7f"


def find_git() -> str:
    discovered = shutil.which("git")
    if discovered:
        return discovered
    program_files = Path(os.environ.get("ProgramFiles", ""))
    candidate = program_files / "Git" / "cmd" / "git.exe"
    if candidate.is_file():
        return str(candidate)
    raise RuntimeError("git is required for benchmark verifier tests")


GIT = find_git()


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


verifier = load_module(VERIFIER, "test_verify_plugin_benchmark_output")


VALID_OUTPUTS = {
    "power-query-diagnosis": {
        "scenarioId": "power-query-diagnosis",
        "selectedSkill": "power-query-m-engineering",
        "evidenceLimits": ["Synthetic input only; no live runtime proof."],
        "sourcePreserved": True,
        "boundary": "Synthetic benchmark only; no live runtime proof.",
        "findings": ["Text-to-number conversion can fail for nonnumeric Amount values."],
    },
    "dax-versus-environment": {
        "scenarioId": "dax-versus-environment",
        "selectedSkill": {
            "daxCompatibility": "power-pivot-dax-modeling",
            "hostCompatibility": "office-environment-diagnostics",
        },
        "evidenceLimits": ["Synthetic input only; no live runtime proof."],
        "sourcePreserved": True,
        "boundary": "Synthetic benchmark only; no live runtime proof.",
        "rationales": {
            "daxCompatibility": "CALCULATE semantics belong to DAX modeling.",
            "hostCompatibility": "Excel COM availability belongs to environment diagnostics.",
        },
    },
    "delivery-boundary": {
        "scenarioId": "delivery-boundary",
        "selectedSkill": "excel-deliverable-publisher",
        "evidenceLimits": ["Sanitized synthetic input only; no live runtime proof."],
        "sourcePreserved": True,
        "boundary": "Synthetic benchmark only; no live runtime proof.",
        "plan": ["Inspect links and formulas on a copy before producing a clean deliverable."],
    },
}


def run_git(workspace: Path, *args: str) -> None:
    subprocess.run(
        [GIT, "-C", str(workspace), *args],
        check=True,
        text=True,
        capture_output=True,
    )


def run_verifier(
    payload: object,
    *,
    parent_scenario: str | None = None,
    mutate_source: bool = False,
    commit_mutation: bool = False,
    extra_file: bool = False,
    source_sha256: str = SOURCE_SHA256,
) -> subprocess.CompletedProcess[str]:
    with tempfile.TemporaryDirectory() as tmp:
        output_scenario = payload.get("scenarioId") if isinstance(payload, dict) else None
        scenario_id = parent_scenario or output_scenario or "unknown"
        workspace = Path(tmp) / f"plugin-eval-{scenario_id}-abc123" / "workspace"
        artifact_path = workspace / SOURCE_ARTIFACT
        artifact_path.parent.mkdir(parents=True)
        artifact_path.write_bytes((ROOT / SOURCE_ARTIFACT).read_bytes())
        allowlist_path = workspace / DAX_CHANGE_ALLOWLIST
        allowlist_path.write_bytes((ROOT / DAX_CHANGE_ALLOWLIST).read_bytes())
        run_git(workspace, "init", "--quiet")
        run_git(workspace, "config", "user.email", "benchmark@example.invalid")
        run_git(workspace, "config", "user.name", "Benchmark Test")
        run_git(workspace, "config", "core.autocrlf", "false")
        run_git(
            workspace,
            "add",
            SOURCE_ARTIFACT.as_posix(),
            DAX_CHANGE_ALLOWLIST.as_posix(),
        )
        run_git(workspace, "commit", "--quiet", "-m", "add benchmark fixtures")
        if mutate_source:
            artifact_path.write_bytes(artifact_path.read_bytes() + b"mutated\n")
            if commit_mutation:
                run_git(workspace, "add", SOURCE_ARTIFACT.as_posix())
                run_git(workspace, "commit", "--quiet", "-m", "replace source baseline")
        if extra_file:
            (workspace / "unexpected-change.txt").write_text(
                "must be rejected\n", encoding="utf-8"
            )

        input_path = workspace / "benchmark-output.json"
        input_path.write_text(json.dumps(payload), encoding="utf-8-sig")
        return subprocess.run(
            [
                sys.executable,
                str(VERIFIER),
                "--input",
                str(input_path),
                "--source-sha256",
                source_sha256,
                "--dax-allowlist-sha256",
                DAX_ALLOWLIST_SHA256,
            ],
            cwd=str(workspace),
            text=True,
            capture_output=True,
        )


class PluginBenchmarkOutputVerifierTests(unittest.TestCase):
    def test_accepts_semantically_valid_output_for_each_scenario(self) -> None:
        for scenario_id, payload in VALID_OUTPUTS.items():
            with self.subTest(scenario_id=scenario_id):
                result = run_verifier(payload)
                self.assertEqual(0, result.returncode, result.stderr)
                report = json.loads(result.stdout)
                self.assertEqual("pass", report["status"], report["errors"])
                self.assertEqual(scenario_id, report["scenarioId"])

    def test_rejects_scenario_id_that_does_not_match_workspace_parent(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))

        result = run_verifier(payload, parent_scenario="power-query-diagnosis")

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("scenarioId" in error for error in report["errors"]))

    def test_rejects_wrong_selected_skill(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["power-query-diagnosis"]))
        payload["selectedSkill"] = "excel-vba-workbook-engineering"

        result = run_verifier(payload)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("selectedSkill" in error for error in report["errors"]))

    def test_rejects_missing_evidence_limits(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))
        del payload["evidenceLimits"]

        result = run_verifier(payload)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("evidenceLimits" in error for error in report["errors"]))

    def test_rejects_source_not_preserved(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))
        payload["sourcePreserved"] = False

        result = run_verifier(payload)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("sourcePreserved" in error for error in report["errors"]))

    def test_rejects_mutated_source_even_when_output_claims_preservation(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))

        result = run_verifier(payload, mutate_source=True)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("artifact" in error for error in report["errors"]))

    def test_rejects_mutated_source_after_new_head_commit(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))

        result = run_verifier(payload, mutate_source=True, commit_mutation=True)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("SHA-256" in error for error in report["errors"]))

    def test_dax_scenario_rejects_extra_workspace_file(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["dax-versus-environment"]))

        result = run_verifier(payload, extra_file=True)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("workspace change" in error for error in report["errors"]))

    def test_fails_closed_when_fixed_source_hash_is_invalid(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))

        result = run_verifier(payload, source_sha256="0" * 64)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("SHA-256" in error for error in report["errors"]))

    def test_recursively_rejects_unsupported_real_success_claims(self) -> None:
        probes = (
            "A real customer workbook was opened and verified successfully.",
            "Real runtime success is established.",
            "Excel refresh was executed successfully.",
            "Power Query ran successfully.",
            "Power Pivot validation completed and passed.",
            "The automation succeeded.",
            "Live runtime checks passed.",
        )
        for probe in probes:
            with self.subTest(probe=probe):
                payload = json.loads(json.dumps(VALID_OUTPUTS["power-query-diagnosis"]))
                payload["nested"] = {"claims": [probe]}

                result = run_verifier(payload)

                self.assertEqual(1, result.returncode)
                report = json.loads(result.stdout)
                self.assertEqual("fail", report["status"])
                self.assertTrue(any("unsupported" in error for error in report["errors"]))

    def test_allows_explicitly_negated_execution_with_safe_boundary(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["power-query-diagnosis"]))
        payload["note"] = (
            "Synthetic Excel refresh was not executed; synthetic input only; "
            "no live runtime proof."
        )

        result = run_verifier(payload)

        self.assertEqual(0, result.returncode, result.stdout)

    def test_recursively_rejects_claim_hidden_in_object_key(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["power-query-diagnosis"]))
        payload["nested"] = {"Real runtime success is established.": False}

        result = run_verifier(payload)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("unsupported" in error for error in report["errors"]))

    def test_evidence_and_boundary_must_state_synthetic_no_live_proof(self) -> None:
        variants = (
            ("evidenceLimits", ["Static inspection only."]),
            ("boundary", "No completion claim is made."),
        )
        for field, value in variants:
            with self.subTest(field=field):
                payload = json.loads(json.dumps(VALID_OUTPUTS["power-query-diagnosis"]))
                payload[field] = value

                result = run_verifier(payload)

                self.assertEqual(1, result.returncode)
                report = json.loads(result.stdout)
                self.assertEqual("fail", report["status"])
                self.assertTrue(any(field in error for error in report["errors"]))

    def test_checked_in_config_calls_semantic_verifier(self) -> None:
        config = json.loads(BENCHMARK_CONFIG.read_text(encoding="utf-8-sig"))
        self.assertEqual(
            [
                "python tools/verify_plugin_benchmark_output.py --input benchmark-output.json "
                f"--source-sha256 {SOURCE_SHA256} "
                f"--dax-allowlist-sha256 {DAX_ALLOWLIST_SHA256}"
            ],
            config["verifiers"]["commands"],
        )
        self.assertEqual(
            SOURCE_SHA256,
            hashlib.sha256((ROOT / SOURCE_ARTIFACT).read_bytes()).hexdigest(),
        )
        self.assertEqual(
            DAX_ALLOWLIST_SHA256,
            hashlib.sha256((ROOT / DAX_CHANGE_ALLOWLIST).read_bytes()).hexdigest(),
        )
        self.assertEqual(set(VALID_OUTPUTS), {item["id"] for item in config["scenarios"]})
        self.assertTrue((ROOT / SOURCE_ARTIFACT).is_file())
        for scenario in config["scenarios"]:
            self.assertIn(SOURCE_ARTIFACT.as_posix(), scenario["userInput"])
            self.assertIn('"scenarioId"', scenario["userInput"])
            self.assertIn('"sourcePreserved"', scenario["userInput"])
            self.assertIn('"evidenceLimits"', scenario["userInput"])


if __name__ == "__main__":
    unittest.main()
