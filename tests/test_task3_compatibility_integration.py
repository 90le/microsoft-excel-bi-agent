from __future__ import annotations

import argparse
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ROUTER_SCRIPT = PROJECT_ROOT / ".agents" / "skills" / "excel-bi-router" / "scripts" / "route_excel_bi_task.py"
PROFILE_SCRIPT = PROJECT_ROOT / "tools" / "run_task_profile.py"
CATALOG_SCRIPT = PROJECT_ROOT / "tools" / "build_capability_catalog.py"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


router = load_module(ROUTER_SCRIPT, "task3_route_excel_bi_task")
task_profile = load_module(PROFILE_SCRIPT, "task3_run_task_profile")
capability_catalog = load_module(CATALOG_SCRIPT, "task3_build_capability_catalog")


def profile_args(out_dir: Path, *, probe_json: str = "", required: list[str] | None = None) -> argparse.Namespace:
    return argparse.Namespace(
        profile="env-diagnostics",
        project_root=str(PROJECT_ROOT),
        workbook="",
        model_json="",
        query_dir="",
        out_dir=str(out_dir),
        out_json="",
        out_md="",
        execute=False,
        probe_json=probe_json,
        require_capability=required or [],
    )


class Task3CompatibilityIntegrationTests(unittest.TestCase):
    def test_platform_compatibility_routes_to_environment_diagnostics(self) -> None:
        report = router.build_report(
            "Can this Excel COM automation run on Linux or macOS, and what compatibility evidence is available?"
        )
        self.assertEqual(report["skill"], "office-environment-diagnostics")
        self.assertEqual(report["layer"], "Office environment")

    def test_dax_compatibility_remains_with_dax_specialist(self) -> None:
        report = router.build_report(
            "Review DAX compatibility of this Power Pivot measure using REMOVEFILTERS in Excel."
        )
        self.assertEqual(report["skill"], "power-pivot-dax-modeling")
        self.assertEqual(report["layer"], "Power Pivot DAX")

    def test_captured_probe_builds_compatibility_report_without_live_probe(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            captured = Path(tmp) / "captured capabilities.json"
            commands = task_profile.profile_commands(
                profile_args(
                    Path(tmp) / "out",
                    probe_json=str(captured),
                    required=["excel.com.activation", "excel.vba.project-access"],
                )
            )

        self.assertEqual([item["name"] for item in commands], ["Build Excel compatibility report"])
        command = commands[0]["command"]
        self.assertIn(str(captured.resolve()), command)
        self.assertEqual(command.count("--require-capability"), 2)
        self.assertIn("--require-pass", command)
        self.assertFalse(any(part.lower().startswith("powershell") for part in command))

    def test_live_environment_profile_keeps_provider_detail_and_adds_compatibility(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            commands = task_profile.profile_commands(profile_args(Path(tmp)))

        names = [item["name"] for item in commands]
        self.assertEqual(
            names,
            [
                "Probe Office and BI providers",
                "Build provider environment report",
                "Probe Excel compatibility capabilities",
                "Build Excel compatibility report",
            ],
        )
        self.assertIn("probe_excel_capabilities.ps1", " ".join(commands[2]["command"]))
        self.assertIn("build_excel_compatibility_report.py", " ".join(commands[3]["command"]))

    def test_catalog_registers_compatibility_tools_and_workflow(self) -> None:
        catalog = capability_catalog.build_catalog(PROJECT_ROOT)
        tools = {item["name"] for item in catalog["tools"]}
        workflows = {item["id"]: item for item in catalog["workflows"]}
        self.assertTrue(
            {
                "probe_excel_capabilities.ps1",
                "build_excel_compatibility_report.py",
                "create_excel_capability_fixture.py",
            }.issubset(tools)
        )
        self.assertIn("excel-compatibility", workflows)
        self.assertEqual(workflows["excel-compatibility"]["skills"], ["office-environment-diagnostics"])

    def test_skills_define_three_evidence_levels_and_separate_target_environment(self) -> None:
        router_skill = (PROJECT_ROOT / ".agents" / "skills" / "excel-bi-router" / "SKILL.md").read_text(
            encoding="utf-8"
        )
        environment_skill = (
            PROJECT_ROOT / ".agents" / "skills" / "office-environment-diagnostics" / "SKILL.md"
        ).read_text(encoding="utf-8")
        combined = router_skill + "\n" + environment_skill
        for phrase in ["Structural evidence", "Runtime capability evidence", "Workbook behavior evidence"]:
            self.assertIn(phrase, combined)
        self.assertIn("execution environment", combined.lower())
        self.assertIn("target environment", combined.lower())


if __name__ == "__main__":
    unittest.main()
