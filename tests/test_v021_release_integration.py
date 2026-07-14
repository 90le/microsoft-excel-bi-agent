from __future__ import annotations

import importlib.util
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RELEASE_GATE = PROJECT_ROOT / "tools" / "run_release_gate.py"
CATALOG_BUILDER = PROJECT_ROOT / "tools" / "build_capability_catalog.py"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


release_gate = load_module(RELEASE_GATE, "v021_run_release_gate")
catalog_builder = load_module(CATALOG_BUILDER, "v021_build_capability_catalog")


def copy_gate_fixture(destination: Path) -> Path:
    project = destination / "plugin"
    for relative in [
        Path(".agents/skills"),
        Path(".codex-plugin"),
        Path("fixtures"),
        Path("benchmarks"),
    ]:
        shutil.copytree(PROJECT_ROOT / relative, project / relative)
    (project / "tools").mkdir(parents=True)
    shutil.copy2(
        PROJECT_ROOT / "tools" / "validate_skill_trigger_benchmark.py",
        project / "tools" / "validate_skill_trigger_benchmark.py",
    )
    return project


class V021ReleaseIntegrationTests(unittest.TestCase):
    def test_release_gate_runs_required_validator_and_reports_checked_in_contract(self) -> None:
        result = release_gate.skill_trigger_efficiency_benchmark_check(PROJECT_ROOT)

        self.assertEqual("Skill trigger efficiency benchmark", result.name)
        self.assertEqual("pass", result.status, result.detail)
        self.assertIsNotNone(result.command)
        self.assertEqual("validate_skill_trigger_benchmark.py", Path(result.command[1]).name)
        self.assertIn("--require-pass", result.command or [])
        self.assertIn(Path(result.command[0]).name.lower(), {"python", "python.exe"})
        self.assertEqual(12, result.metadata["skillCount"])
        self.assertEqual(36, result.metadata["caseCount"])
        self.assertEqual(3, result.metadata["realBenchmarkScenarioCount"])
        self.assertEqual(3, result.metadata["promptCount"])
        self.assertLessEqual(result.metadata["maxPromptChars"], 110)
        self.assertEqual(
            ["dax-versus-environment", "delivery-boundary", "power-query-diagnosis"],
            sorted(result.metadata["scenarioIds"]),
        )
        catalog_check = release_gate.capability_catalog_fixture_check(PROJECT_ROOT)
        self.assertEqual("pass", catalog_check.status, catalog_check.detail)

    def test_release_gate_rejects_unsafe_workspace_sources(self) -> None:
        unsafe_sources = [
            "file:///tmp/plugin",
            "https://example.test/plugin",
            "/tmp/plugin",
            "\\\\server\\share\\plugin",
            "C:\\customer\\plugin",
            "C:relative-but-drive-prefixed",
            "../outside",
            "workspace/../../outside",
        ]
        for source_path in unsafe_sources:
            with self.subTest(source_path=source_path), tempfile.TemporaryDirectory() as tmp:
                project = copy_gate_fixture(Path(tmp))
                path = project / "benchmarks" / "plugin-eval-v0.2.1.json"
                config = json.loads(path.read_text(encoding="utf-8-sig"))
                config["workspace"]["sourcePath"] = source_path
                path.write_text(json.dumps(config, indent=2), encoding="utf-8")

                result = release_gate.skill_trigger_efficiency_benchmark_check(project)

                self.assertEqual("fail", result.status)

    def test_release_gate_rejects_valid_scenario_id_with_generic_content(self) -> None:
        for scenario_index in range(3):
            with self.subTest(scenario_index=scenario_index), tempfile.TemporaryDirectory() as tmp:
                project = copy_gate_fixture(Path(tmp))
                path = project / "benchmarks" / "plugin-eval-v0.2.1.json"
                config = json.loads(path.read_text(encoding="utf-8-sig"))
                scenario_id = config["scenarios"][scenario_index]["id"]
                config["scenarios"][scenario_index] = {
                    "id": scenario_id,
                    "title": "Scenario",
                    "purpose": "Run a generic task.",
                    "userInput": (
                        "Write benchmark-output.json and repeat: synthetic input only; "
                        "no live runtime proof."
                    ),
                    "successChecklist": ["The generic task is done."],
                }
                path.write_text(json.dumps(config, indent=2), encoding="utf-8")

                result = release_gate.skill_trigger_efficiency_benchmark_check(project)

                self.assertEqual("fail", result.status)

    def test_release_gate_rejects_wrong_starter_prompt_count(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = copy_gate_fixture(Path(tmp))
            path = project / ".codex-plugin" / "plugin.json"
            manifest = json.loads(path.read_text(encoding="utf-8-sig"))
            manifest["interface"]["defaultPrompt"].append("A fourth prompt.")
            path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

            result = release_gate.skill_trigger_efficiency_benchmark_check(project)

            self.assertEqual("fail", result.status)

    def test_release_gate_rejects_oversized_starter_prompt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = copy_gate_fixture(Path(tmp))
            path = project / ".codex-plugin" / "plugin.json"
            manifest = json.loads(path.read_text(encoding="utf-8-sig"))
            manifest["interface"]["defaultPrompt"][0] = "x" * 111
            path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

            result = release_gate.skill_trigger_efficiency_benchmark_check(project)

            self.assertEqual("fail", result.status)

    def test_release_checks_fail_closed_for_non_object_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = copy_gate_fixture(Path(tmp))
            path = project / ".codex-plugin" / "plugin.json"
            path.write_text("[]\n", encoding="utf-8")

            manifest_result = release_gate.plugin_manifest_release_check(project)
            benchmark_result = release_gate.skill_trigger_efficiency_benchmark_check(project)

            self.assertEqual("fail", manifest_result.status)
            self.assertEqual("fail", benchmark_result.status)

    def test_release_gate_fails_closed_when_validator_is_nonzero(self) -> None:
        validator_result = release_gate.CheckResult(
            name="validator",
            status="fail",
            detail="exit_code=1",
            stdout=json.dumps(
                {
                    "kind": "excel-skill-trigger-benchmark-report",
                    "schemaVersion": "1.0",
                    "status": "fail",
                    "summary": {},
                }
            ),
        )
        with mock.patch.object(release_gate, "run_command", return_value=validator_result):
            result = release_gate.skill_trigger_efficiency_benchmark_check(PROJECT_ROOT)

        self.assertEqual("fail", result.status)
        self.assertIn("exit_code=1", result.detail)

    def test_release_gate_fails_closed_for_malformed_validator_stdout(self) -> None:
        validator_result = release_gate.CheckResult(
            name="validator", status="pass", detail="exit_code=0", stdout="not JSON"
        )
        with mock.patch.object(release_gate, "run_command", return_value=validator_result):
            result = release_gate.skill_trigger_efficiency_benchmark_check(PROJECT_ROOT)

        self.assertEqual("fail", result.status)
        self.assertIn("cannot parse trigger validator JSON", result.detail)

    def test_release_gate_fails_closed_for_non_object_validator_report(self) -> None:
        for payload in [[], None, "report"]:
            with self.subTest(payload=payload):
                validator_result = release_gate.CheckResult(
                    name="validator",
                    status="pass",
                    detail="exit_code=0",
                    stdout=json.dumps(payload),
                )
                with mock.patch.object(release_gate, "run_command", return_value=validator_result):
                    result = release_gate.skill_trigger_efficiency_benchmark_check(PROJECT_ROOT)

                self.assertEqual("fail", result.status)

    def test_release_gate_fails_closed_for_non_object_benchmark_config(self) -> None:
        for payload in [[], None, "config"]:
            with self.subTest(payload=payload), tempfile.TemporaryDirectory() as tmp:
                project = copy_gate_fixture(Path(tmp))
                path = project / "benchmarks" / "plugin-eval-v0.2.1.json"
                path.write_text(json.dumps(payload), encoding="utf-8")

                result = release_gate.skill_trigger_efficiency_benchmark_check(project)

                self.assertEqual("fail", result.status)

    def test_capability_catalog_exposes_trigger_validator_and_workflow(self) -> None:
        catalog = catalog_builder.build_catalog(PROJECT_ROOT)
        tools = {item["name"]: item for item in catalog["tools"]}
        workflows = {item["id"]: item for item in catalog["workflows"]}

        self.assertEqual("pass", catalog["status"], catalog["findings"])
        self.assertTrue(tools["validate_skill_trigger_benchmark.py"]["required"])
        workflow = workflows["skill-trigger-benchmark"]
        self.assertIn("validate_skill_trigger_benchmark.py", workflow["tools"])
        self.assertIn("synthetic", workflow["boundary"].lower())
        self.assertIn("real task", workflow["boundary"].lower())


if __name__ == "__main__":
    unittest.main()
