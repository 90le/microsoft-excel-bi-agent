from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CASES = ROOT / "fixtures" / "skill-trigger-benchmark.json"
VALIDATOR = ROOT / "tools" / "validate_skill_trigger_benchmark.py"

CANONICAL_SKILLS = (
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
)


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


benchmark = load_module(VALIDATOR, "test_validate_skill_trigger_benchmark")
validate_trigger_cases = benchmark.validate_trigger_cases


def write_synthetic_project(root: Path, *, invalid: bool = False) -> None:
    plugin_dir = root / ".codex-plugin"
    plugin_dir.mkdir(parents=True)
    prompts = [
        "Inspect this synthetic Excel input.",
        "Diagnose this synthetic Excel issue.",
        "Publish this synthetic Excel deliverable.",
    ]
    if invalid:
        prompts[0] = "x" * 111
    (plugin_dir / "plugin.json").write_text(
        json.dumps(
            {
                "name": "synthetic-excel-bi-plugin",
                "version": "0.0.0",
                "skills": "./skills/",
                "interface": {"defaultPrompt": prompts},
            }
        ),
        encoding="utf-8-sig",
    )

    for index, skill_id in enumerate(CANONICAL_SKILLS):
        skill_dir = root / ".agents" / "skills" / skill_id
        skill_dir.mkdir(parents=True)
        description = f"Use when synthetic Excel work must route to {skill_id}."
        if invalid and index == 0:
            description = "Synthetic metadata without the required trigger prefix."
        (skill_dir / "SKILL.md").write_text(
            f"---\nname: {skill_id}\ndescription: {description}\n---\n\n# Synthetic skill\n",
            encoding="utf-8-sig",
        )


def run_validator(project: Path, *, require_pass: bool) -> subprocess.CompletedProcess[str]:
    command = [
        sys.executable,
        str(VALIDATOR),
        "--project-root",
        str(project),
        "--cases-json",
        str(CASES),
    ]
    if require_pass:
        command.append("--require-pass")
    return subprocess.run(command, cwd=str(ROOT), text=True, capture_output=True)


class SkillTriggerBenchmarkTests(unittest.TestCase):
    def test_trigger_corpus_covers_every_skill(self) -> None:
        report = validate_trigger_cases(ROOT, CASES)
        self.assertEqual("pass", report["status"])
        self.assertEqual(
            {
                "skillCount": 12,
                "caseCount": 36,
                "positiveCount": 24,
                "confusableNegativeCount": 12,
            },
            report["summary"],
        )

    def test_valid_utf8_sig_synthetic_metadata_passes_cli(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = Path(tmp) / "synthetic-plugin"
            write_synthetic_project(project)
            result = run_validator(project, require_pass=True)

        self.assertEqual(0, result.returncode, result.stderr)
        report = json.loads(result.stdout)
        self.assertEqual("excel-skill-trigger-benchmark-report", report["kind"])
        self.assertEqual("1.0", report["schemaVersion"])
        self.assertEqual("pass", report["status"])

    def test_require_pass_rejects_invalid_synthetic_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = Path(tmp) / "synthetic-invalid-plugin"
            write_synthetic_project(project, invalid=True)
            result = run_validator(project, require_pass=True)

        self.assertEqual(1, result.returncode)
        self.assertEqual("fail", json.loads(result.stdout)["status"])

    def test_case_schema_rejects_unknown_fields_and_duplicate_ids(self) -> None:
        source = json.loads(CASES.read_text(encoding="utf-8-sig"))
        variants = []

        unknown = json.loads(json.dumps(source))
        unknown["cases"][0]["unexpected"] = True
        variants.append(unknown)

        duplicate = json.loads(json.dumps(source))
        duplicate["cases"][1]["id"] = duplicate["cases"][0]["id"]
        variants.append(duplicate)

        with tempfile.TemporaryDirectory() as tmp:
            for index, variant in enumerate(variants):
                with self.subTest(index=index):
                    path = Path(tmp) / f"invalid-{index}.json"
                    path.write_text(json.dumps(variant), encoding="utf-8")
                    report = validate_trigger_cases(ROOT, path)
                    self.assertEqual("fail", report["status"])
                    self.assertTrue(report["errors"])

    def test_case_text_rejects_private_or_credential_like_artifacts(self) -> None:
        source = json.loads(CASES.read_text(encoding="utf-8-sig"))
        unsafe_texts = (
            r"Open (C:\Users\customer\private\report) and inspect it.",
            r"Open '\\server\share\private\report' and inspect it.",
            "Open (/srv/customer/private/report) and inspect it.",
            r"Open C:\Users\customer\private\report.xlsx and inspect it.",
            "Connect with password=Summer2026! and refresh the query.",
            "Inspect customer-finance.xlsx and publish a clean copy.",
        )

        with tempfile.TemporaryDirectory() as tmp:
            for index, unsafe_text in enumerate(unsafe_texts):
                with self.subTest(text=unsafe_text):
                    variant = json.loads(json.dumps(source))
                    variant["cases"][0]["text"] = unsafe_text
                    path = Path(tmp) / f"unsafe-{index}.json"
                    path.write_text(json.dumps(variant), encoding="utf-8")
                    report = validate_trigger_cases(ROOT, path)
                    self.assertEqual("fail", report["status"])
                    self.assertTrue(report["errors"])

    def test_case_known_fields_reject_wrong_types_without_crashing(self) -> None:
        source = json.loads(CASES.read_text(encoding="utf-8-sig"))
        invalid_values = (None, 42, [], {})

        with tempfile.TemporaryDirectory() as tmp:
            for field in ("kind", "targetSkill", "expectedSkill", "text"):
                for value_index, invalid_value in enumerate(invalid_values):
                    with self.subTest(field=field, value=invalid_value):
                        variant = json.loads(json.dumps(source))
                        variant["cases"][0][field] = invalid_value
                        path = Path(tmp) / f"{field}-{value_index}.json"
                        path.write_text(json.dumps(variant), encoding="utf-8")
                        report = validate_trigger_cases(ROOT, path)
                        self.assertEqual("fail", report["status"])
                        self.assertTrue(any(field in error for error in report["errors"]))

            for value_index, invalid_value in enumerate((None, 42, {}, "not-an-array", ["ok", 42])):
                with self.subTest(field="successChecklist", value=invalid_value):
                    variant = json.loads(json.dumps(source))
                    variant["cases"][0]["successChecklist"] = invalid_value
                    path = Path(tmp) / f"success-checklist-{value_index}.json"
                    path.write_text(json.dumps(variant), encoding="utf-8")
                    report = validate_trigger_cases(ROOT, path)
                    self.assertEqual("fail", report["status"])
                    self.assertTrue(
                        any("successChecklist" in error for error in report["errors"])
                    )

    def test_manifest_known_fields_reject_wrong_types_without_crashing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = Path(tmp) / "synthetic-plugin"
            write_synthetic_project(project)
            manifest_path = project / ".codex-plugin" / "plugin.json"
            valid_manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))

            for value_index, invalid_interface in enumerate((None, 42, [], "not-an-object")):
                with self.subTest(field="interface", value=invalid_interface):
                    manifest = json.loads(json.dumps(valid_manifest))
                    manifest["interface"] = invalid_interface
                    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
                    report = benchmark.validate_benchmark(project, CASES)
                    self.assertEqual("fail", report["status"])
                    self.assertTrue(any("interface" in error for error in report["errors"]))

            for value_index, invalid_prompts in enumerate(
                (None, 42, {}, "not-an-array", ["valid", None, "valid"])
            ):
                with self.subTest(field="defaultPrompt", value=invalid_prompts):
                    manifest = json.loads(json.dumps(valid_manifest))
                    manifest["interface"]["defaultPrompt"] = invalid_prompts
                    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
                    report = benchmark.validate_benchmark(project, CASES)
                    self.assertEqual("fail", report["status"])
                    self.assertTrue(any("defaultPrompt" in error for error in report["errors"]))


if __name__ == "__main__":
    unittest.main()
