from __future__ import annotations

import importlib.util
import json
import re
import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RELEASE_GATE = PROJECT_ROOT / "tools" / "run_release_gate.py"
CASE_RUNNER = PROJECT_ROOT / "tools" / "run_case_regression.py"
DOC_VALIDATOR = PROJECT_ROOT / "tools" / "validate_project_docs.py"
VERSION = "0.2.1+codex.20260714"
STABLE_RELEASE = "0.2.0"


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
doc_validator = load_module(DOC_VALIDATOR, "task5_validate_project_docs")


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
        self.assertIn("Current Stable Release", current_status)
        self.assertIn(f"v{STABLE_RELEASE}", current_status)
        self.assertIn("Unreleased Release Candidate", current_status)
        self.assertIn(VERSION, current_status)
        self.assertIn("Unreleased Release Candidate", release_en)
        self.assertIn(VERSION, release_en)
        self.assertIn("未发布候选", release_zh)
        self.assertIn(VERSION, release_zh)

        readme_en = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")
        readme_zh = (PROJECT_ROOT / "README.zh-CN.md").read_text(encoding="utf-8")
        self.assertIn(f"Current stable release: **v{STABLE_RELEASE}**", readme_en)
        self.assertIn("Unreleased release candidate: **v0.2.1**", readme_en)
        self.assertIn(f"当前稳定版：**v{STABLE_RELEASE}**", readme_zh)
        self.assertIn("未发布候选：**v0.2.1**", readme_zh)
        for readme in [readme_en.lower(), readme_zh.lower()]:
            self.assertIn("docs/compatibility.md", readme)
            self.assertIn("runtime", readme)

    def test_landing_pages_keep_stable_release_link_and_label_manifest_candidate(self) -> None:
        manifest = json.loads((PROJECT_ROOT / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8"))
        candidate_release = manifest["version"].split("+", 1)[0]
        pages = [
            (PROJECT_ROOT / "docs" / "intro.html", "Unreleased release candidate"),
            (PROJECT_ROOT / "docs" / "intro.zh-CN.html", "未发布候选"),
        ]
        for path, candidate_marker in pages:
            landing = path.read_text(encoding="utf-8")
            release_links = re.findall(r"/releases/tag/v(\d+\.\d+\.\d+)", landing)
            self.assertTrue(release_links, path)
            self.assertEqual(STABLE_RELEASE, release_links[0])
            self.assertNotIn(candidate_release, release_links)
            self.assertIn(f"v{candidate_release}", landing)
            self.assertIn(candidate_marker, landing)
            self.assertIn("./compatibility.md", landing)

        report = doc_validator.validate(PROJECT_ROOT)
        self.assertEqual(report["status"], "pass", report["errors"])

    def test_benchmark_claims_require_synthetic_and_observed_evidence_boundary(self) -> None:
        unsafe = {
            "docs/release-notes.en-US.md": (
                "The synthetic trigger benchmark proves real task success with a 98% success result."
            )
        }
        safe = {
            "docs/release-notes.en-US.md": (
                "The 36-case trigger benchmark uses synthetic inputs. Synthetic benchmark output "
                "validates mechanics only and does not prove real task success; observed usage is "
                "separate evidence."
            )
        }

        self.assertTrue(doc_validator.benchmark_evidence_boundary_errors(unsafe))
        self.assertEqual([], doc_validator.benchmark_evidence_boundary_errors(safe))

    def test_benchmark_output_format_reference_is_not_an_evidence_claim(self) -> None:
        subject_only = {
            "docs/task-recipes.md": "See the benchmark output format in benchmark-output.json."
        }

        self.assertEqual([], doc_validator.benchmark_evidence_boundary_errors(subject_only))

    def test_benchmark_evidence_claim_is_detected_when_claim_precedes_subject(self) -> None:
        unsafe = {
            "docs/release-notes.en-US.md": "Evidence from the trigger benchmark is conclusive."
        }

        self.assertTrue(doc_validator.benchmark_evidence_boundary_errors(unsafe))

    def test_benchmark_boundaries_do_not_leak_across_sections(self) -> None:
        mixed = {
            "docs/release-notes.en-US.md": (
                "## Mechanics\n\nSynthetic benchmark output validates mechanics only and does not prove "
                "real task success; observed usage is separate evidence.\n\n"
                "## Results\n\nThe trigger benchmark proves real task success with a 98% success result."
            )
        }

        self.assertTrue(doc_validator.benchmark_evidence_boundary_errors(mixed))

    def test_benchmark_boundaries_do_not_leak_between_claim_blocks_in_one_section(self) -> None:
        mixed = {
            "docs/release-notes.en-US.md": (
                "## Results\n\nSynthetic benchmark output validates mechanics only and does not prove "
                "real task success; observed usage is separate evidence.\n\n"
                "The trigger benchmark proves real task success with a 98% success result."
            )
        }

        self.assertTrue(doc_validator.benchmark_evidence_boundary_errors(mixed))

    def test_positive_real_success_assertion_is_rejected_even_with_same_block_disclaimer(self) -> None:
        contradictory_claims = [
            (
                "The synthetic trigger benchmark proves real task success. Synthetic benchmark "
                "output does not prove real task success; observed usage is separate evidence."
            ),
            (
                "The trigger benchmark establishes actual task success. Synthetic output does not "
                "prove real task success; observed usage is separate evidence."
            ),
            (
                "The trigger benchmark demonstrates live workbook success. Synthetic output does not "
                "prove real task success; observed usage is separate evidence."
            ),
            (
                "The trigger benchmark confirms real workbook success. Synthetic output does not prove "
                "real task success; observed usage is separate evidence."
            ),
            (
                "The trigger benchmark validates actual workbook success. Synthetic output does not "
                "prove real task success; observed usage is separate evidence."
            ),
            (
                "The trigger benchmark shows real task success. Synthetic output does not prove real "
                "task success; observed usage is separate evidence."
            ),
            (
                "The trigger benchmark achieves actual task success. Synthetic output does not prove "
                "real task success; observed usage is separate evidence."
            ),
        ]
        for text in contradictory_claims:
            with self.subTest(text=text):
                self.assertTrue(
                    doc_validator.benchmark_evidence_boundary_errors(
                        {"docs/release-notes.en-US.md": text}
                    )
                )

    def test_benchmark_schema_and_file_references_are_not_outcome_claims(self) -> None:
        references = [
            "The benchmark results schema includes score and evidence fields.",
            "The plugin-eval result is stored in benchmark-result.json.",
            "Use the benchmark score field when parsing the JSON format.",
        ]
        for text in references:
            with self.subTest(text=text):
                self.assertEqual(
                    [],
                    doc_validator.benchmark_evidence_boundary_errors(
                        {"docs/task-recipes.md": text}
                    ),
                )


if __name__ == "__main__":
    unittest.main()
