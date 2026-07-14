from __future__ import annotations

import importlib.util
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
    commit_source: bool = True,
) -> subprocess.CompletedProcess[str]:
    with tempfile.TemporaryDirectory() as tmp:
        output_scenario = payload.get("scenarioId") if isinstance(payload, dict) else None
        scenario_id = parent_scenario or output_scenario or "unknown"
        workspace = Path(tmp) / f"plugin-eval-{scenario_id}-abc123" / "workspace"
        artifact_path = workspace / SOURCE_ARTIFACT
        artifact_path.parent.mkdir(parents=True)
        artifact_path.write_text(
            json.dumps(
                {
                    "synthetic": True,
                    "label": "sanitized Excel BI benchmark source",
                    "amountValues": ["10", "not-a-number"],
                },
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
        run_git(workspace, "init", "--quiet")
        run_git(workspace, "config", "user.email", "benchmark@example.invalid")
        run_git(workspace, "config", "user.name", "Benchmark Test")
        run_git(workspace, "config", "core.autocrlf", "false")
        if commit_source:
            run_git(workspace, "add", SOURCE_ARTIFACT.as_posix())
            run_git(workspace, "commit", "--quiet", "-m", "add sanitized source")
        if mutate_source:
            artifact_path.write_bytes(artifact_path.read_bytes() + b"mutated\n")

        input_path = workspace / "benchmark-output.json"
        input_path.write_text(json.dumps(payload), encoding="utf-8-sig")
        return subprocess.run(
            [sys.executable, str(VERIFIER), "--input", str(input_path)],
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

    def test_fails_closed_when_git_baseline_is_unavailable(self) -> None:
        payload = json.loads(json.dumps(VALID_OUTPUTS["delivery-boundary"]))

        result = run_verifier(payload, commit_source=False)

        self.assertEqual(1, result.returncode)
        report = json.loads(result.stdout)
        self.assertEqual("fail", report["status"])
        self.assertTrue(any("baseline" in error for error in report["errors"]))

    def test_recursively_rejects_unsupported_real_success_claims(self) -> None:
        probes = (
            "A real customer workbook was opened and verified successfully.",
            "Real runtime success is established.",
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
            ["python tools/verify_plugin_benchmark_output.py --input benchmark-output.json"],
            config["verifiers"]["commands"],
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
