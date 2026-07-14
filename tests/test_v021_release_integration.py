from __future__ import annotations

import importlib.util
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path


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
        self.assertIn("validate_skill_trigger_benchmark.py", " ".join(result.command or []))
        self.assertIn("--require-pass", result.command or [])
        self.assertNotIn("codex", " ".join(result.command or []).lower())
        self.assertNotIn("excel", Path(result.command[0]).name.lower())
        self.assertEqual(12, result.metadata["skillCount"])
        self.assertEqual(36, result.metadata["caseCount"])
        self.assertEqual(3, result.metadata["realBenchmarkScenarioCount"])
        self.assertEqual(3, result.metadata["promptCount"])
        self.assertLessEqual(result.metadata["maxPromptChars"], 110)
        self.assertEqual(
            ["dax-versus-environment", "delivery-boundary", "power-query-diagnosis"],
            sorted(result.metadata["scenarioIds"]),
        )
        gate_source = RELEASE_GATE.read_text(encoding="utf-8")
        self.assertIn(
            "checks.append(skill_trigger_efficiency_benchmark_check(project_root))",
            gate_source,
        )

    def test_release_gate_rejects_absolute_workspace_and_generic_scenario(self) -> None:
        mutations = {
            "absolute workspace": lambda config: config["workspace"].__setitem__(
                "sourcePath", "C:/customer/workbook"
            ),
            "generic scenario": lambda config: config["scenarios"].__setitem__(
                0,
                {
                    "id": "scenario-1",
                    "title": "Scenario",
                    "purpose": "Run a generic task.",
                    "userInput": "Do the task.",
                    "successChecklist": ["The task is done."],
                },
            ),
        }
        for label, mutate in mutations.items():
            with self.subTest(label=label), tempfile.TemporaryDirectory() as tmp:
                project = copy_gate_fixture(Path(tmp))
                path = project / "benchmarks" / "plugin-eval-v0.2.1.json"
                config = json.loads(path.read_text(encoding="utf-8-sig"))
                mutate(config)
                path.write_text(json.dumps(config, indent=2), encoding="utf-8")

                result = release_gate.skill_trigger_efficiency_benchmark_check(project)

                self.assertEqual("fail", result.status)

    def test_release_gate_rejects_a_fourth_or_oversized_starter_prompt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = copy_gate_fixture(Path(tmp))
            path = project / ".codex-plugin" / "plugin.json"
            manifest = json.loads(path.read_text(encoding="utf-8-sig"))
            manifest["interface"]["defaultPrompt"] = ["x" * 111] * 4
            path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

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
