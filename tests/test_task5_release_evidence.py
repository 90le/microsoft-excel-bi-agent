from __future__ import annotations

import importlib.util
import json
import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RELEASE_GATE = PROJECT_ROOT / "tools" / "run_release_gate.py"
CASE_RUNNER = PROJECT_ROOT / "tools" / "run_case_regression.py"
VERSION = "0.2.0+codex.20260714"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


release_gate = load_module(RELEASE_GATE, "task5_run_release_gate")
case_runner = load_module(CASE_RUNNER, "task5_run_case_regression")


class Task5ReleaseEvidenceTests(unittest.TestCase):
    def test_structural_release_checks_cover_compatibility_runtime_and_manifest(self) -> None:
        compatibility = release_gate.excel_compatibility_fixture_report_check(PROJECT_ROOT)
        runtime = release_gate.runtime_package_fixture_check(PROJECT_ROOT)
        manifest = release_gate.plugin_manifest_release_check(PROJECT_ROOT)

        self.assertEqual(compatibility.status, "pass", compatibility.detail)
        self.assertEqual(compatibility.metadata["fixtureCount"], 4)
        self.assertEqual(compatibility.metadata["reportCount"], 4)
        self.assertEqual(runtime.status, "pass", runtime.detail)
        self.assertIn(
            "fixtures/real-sanitized-cases/cases/excel-capability-routing.json",
            runtime.metadata["paths"],
        )
        self.assertEqual(manifest.status, "pass", manifest.detail)
        self.assertEqual(manifest.metadata["version"], VERSION)
        self.assertEqual(manifest.metadata["promptCount"], 3)

    def test_live_probe_is_optional_without_powershell(self) -> None:
        result = release_gate.excel_capability_live_probe_check(PROJECT_ROOT, None)
        self.assertEqual(result.status, "skip")

    def test_environment_case_is_manifested_and_valid(self) -> None:
        case_root = PROJECT_ROOT / "fixtures" / "real-sanitized-cases"
        manifest = json.loads((case_root / "manifest.json").read_text(encoding="utf-8"))
        refs = {item["id"]: item["specPath"] for item in manifest["cases"]}
        self.assertEqual(refs["excel-capability-routing"], "cases/excel-capability-routing.json")

        report = case_runner.validate(PROJECT_ROOT, case_root)
        self.assertEqual(report["status"], "pass", report["errors"])
        self.assertIn("environment", report["coveredLayers"])

        case = json.loads((case_root / refs["excel-capability-routing"]).read_text(encoding="utf-8"))
        case_text = json.dumps(case, ensure_ascii=False).lower()
        for phrase in [
            "capability probe",
            "office-environment-diagnostics",
            "power-pivot-dax-modeling",
            "structural evidence",
            "runtime capability evidence",
            "workbook behavior evidence",
        ]:
            self.assertIn(phrase, case_text)

        gate_result = release_gate.real_sanitized_case_regression_check(PROJECT_ROOT)
        self.assertEqual(gate_result.status, "pass", gate_result.detail)
        self.assertEqual(gate_result.metadata["caseCount"], 7)
        self.assertIn("environment", gate_result.metadata["coveredLayers"])
        self.assertIn("7 sanitized", gate_result.detail)

    def test_public_docs_define_release_and_compatibility_contract(self) -> None:
        compatibility = (PROJECT_ROOT / "docs" / "compatibility.md").read_text(encoding="utf-8").lower()
        for phrase in [
            "windows",
            "macos",
            "excel for web",
            "excel 2007",
            "excel 2010",
            "excel 2013",
            "excel 2016",
            "excel 2019",
            "ltsc",
            "microsoft 365",
            "32-bit",
            "64-bit",
            "offline",
            "third-party",
            "structural evidence",
            "runtime capability evidence",
            "workbook behavior evidence",
            "authoring target",
            "automation target",
            "consumer target",
            "recipient target",
            "confidence",
        ]:
            self.assertIn(phrase, compatibility)

        manifest = json.loads((PROJECT_ROOT / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["version"], VERSION)
        self.assertEqual(len(manifest["interface"]["defaultPrompt"]), 3)

        current_status = (PROJECT_ROOT / "docs" / "current-status.md").read_text(encoding="utf-8")
        release_en = (PROJECT_ROOT / "docs" / "release-notes.en-US.md").read_text(encoding="utf-8")
        release_zh = (PROJECT_ROOT / "docs" / "release-notes.zh-CN.md").read_text(encoding="utf-8")
        for text in [current_status, release_en, release_zh]:
            self.assertIn("v0.2.0", text)
            self.assertIn(VERSION, text)

        for path in [PROJECT_ROOT / "README.md", PROJECT_ROOT / "README.zh-CN.md"]:
            readme = path.read_text(encoding="utf-8").lower()
            self.assertIn("v0.2.0", readme)
            self.assertIn("docs/compatibility.md", readme)
            self.assertIn("runtime", readme)


if __name__ == "__main__":
    unittest.main()
