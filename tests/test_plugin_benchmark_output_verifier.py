from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VERIFIER = ROOT / "tools" / "verify_plugin_benchmark_output.py"
BENCHMARK_CONFIG = ROOT / "benchmarks" / "plugin-eval-v0.2.1.json"


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
        "evidenceLimits": ["The M query was inspected statically; Excel refresh was not executed."],
        "sourcePreserved": True,
        "boundary": "The report does not claim Excel runtime behavior.",
        "findings": ["Text-to-number conversion can fail for nonnumeric Amount values."],
    },
    "dax-versus-environment": {
        "scenarioId": "dax-versus-environment",
        "selectedSkill": {
            "daxCompatibility": "power-pivot-dax-modeling",
            "hostCompatibility": "office-environment-diagnostics",
        },
        "evidenceLimits": ["No DAX model or Linux Excel runtime was executed."],
        "sourcePreserved": True,
        "boundary": "Formula compatibility and host compatibility are separate decisions.",
        "rationales": {
            "daxCompatibility": "CALCULATE semantics belong to DAX modeling.",
            "hostCompatibility": "Excel COM availability belongs to environment diagnostics.",
        },
    },
    "delivery-boundary": {
        "scenarioId": "delivery-boundary",
        "selectedSkill": "excel-deliverable-publisher",
        "evidenceLimits": ["The sanitized case was planned from metadata; Excel was not opened."],
        "sourcePreserved": True,
        "boundary": "The output is an audit plan, not proof of a completed workbook delivery.",
        "plan": ["Inspect links and formulas on a copy before producing a clean deliverable."],
    },
}


def run_verifier(payload: object) -> subprocess.CompletedProcess[str]:
    with tempfile.TemporaryDirectory() as tmp:
        input_path = Path(tmp) / "benchmark-output.json"
        input_path.write_text(json.dumps(payload), encoding="utf-8-sig")
        return subprocess.run(
            [sys.executable, str(VERIFIER), "--input", str(input_path)],
            cwd=str(ROOT),
            text=True,
            capture_output=True,
        )


class PluginBenchmarkOutputVerifierTests(unittest.TestCase):
    def test_accepts_semantically_valid_output_for_each_scenario(self) -> None:
        for scenario_id, payload in VALID_OUTPUTS.items():
            with self.subTest(scenario_id=scenario_id):
                report = verifier.verify_benchmark_output(payload)
                self.assertEqual("pass", report["status"], report["errors"])
                self.assertEqual(scenario_id, report["scenarioId"])

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

    def test_checked_in_config_calls_semantic_verifier(self) -> None:
        config = json.loads(BENCHMARK_CONFIG.read_text(encoding="utf-8-sig"))
        self.assertEqual(
            ["python tools/verify_plugin_benchmark_output.py --input benchmark-output.json"],
            config["verifiers"]["commands"],
        )
        self.assertEqual(set(VALID_OUTPUTS), {item["id"] for item in config["scenarios"]})
        for scenario in config["scenarios"]:
            self.assertIn('"scenarioId"', scenario["userInput"])
            self.assertIn('"sourcePreserved"', scenario["userInput"])
            self.assertIn('"evidenceLimits"', scenario["userInput"])


if __name__ == "__main__":
    unittest.main()
