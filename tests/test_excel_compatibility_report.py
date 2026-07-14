from __future__ import annotations

import importlib.util
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TOOLS_DIR = PROJECT_ROOT / "tools"
REPORT_SCRIPT = TOOLS_DIR / "build_excel_compatibility_report.py"
FIXTURE_SCRIPT = TOOLS_DIR / "create_excel_capability_fixture.py"
PROBE_SCRIPT = TOOLS_DIR / "probe_excel_capabilities.ps1"

CAPABILITY_IDS = {
    "excel.com.activation",
    "excel.workbook.roundtrip",
    "excel.vba.project-access",
    "excel.power-query.object-model",
    "excel.power-query.async-wait",
    "excel.data-model.object-model",
    "excel.pdf-export",
    "ado.com.activation",
    "ace.workbook-sql",
    "msolap.registration",
    "adomd.com.activation",
}

OPERATION_IDS = {
    "workbook-automation",
    "vba-engineering",
    "power-query-automation",
    "data-model-inspection",
    "ado-workbook-sql",
    "adomd-endpoint-query",
    "rendered-pdf-evidence",
}


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def capability(status: str = "pass", evidence_level: str = "smoke") -> dict[str, str]:
    return {
        "status": status,
        "evidenceLevel": evidence_level,
        "detail": "generic fixture evidence",
        "errorCategory": "",
        "error": "",
    }


def valid_probe() -> dict[str, object]:
    return {
        "schemaVersion": "1.0",
        "kind": "excel-capability-probe",
        "generatedAt": "2026-07-14T00:00:00+00:00",
        "probe": {"profile": "runtime", "platform": "windows", "syntheticFixture": True},
        "environment": {
            "osVersion": "Generic Windows fixture",
            "is64BitOperatingSystem": True,
            "is64BitProcess": True,
            "powershellVersion": "5.1",
            "excelVersion": "16.0",
            "excelBuild": "00000",
        },
        "capabilities": {item: capability() for item in sorted(CAPABILITY_IDS)},
        "boundaries": ["Synthetic fixture; not local Office evidence."],
        "errors": [],
    }


class ExcelCompatibilityReportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.report_module = (
            load_module(REPORT_SCRIPT, "build_excel_compatibility_report") if REPORT_SCRIPT.is_file() else None
        )

    def setUp(self) -> None:
        if self._testMethodName != "test_report_script_exists" and self.report_module is None:
            self.skipTest("report implementation does not exist yet")

    def test_report_script_exists(self) -> None:
        self.assertTrue(REPORT_SCRIPT.is_file(), "report implementation is missing")

    def test_declares_exact_stable_capability_ids(self) -> None:
        self.assertEqual(set(self.report_module.CAPABILITY_IDS), CAPABILITY_IDS)

    def test_supported_probe_is_ready(self) -> None:
        report = self.report_module.build_report(valid_probe(), required_capabilities=[])
        self.assertEqual(report["status"], "pass")
        self.assertEqual(report["summary"]["passCount"], 11)
        self.assertEqual(report["summary"]["failCount"], 0)
        self.assertEqual(report["requirements"]["unmet"], [])
        self.assertEqual(report["operations"]["workbook-automation"]["readiness"], "ready")

    def test_missing_capability_is_evidence_not_contract_crash(self) -> None:
        probe = valid_probe()
        capabilities = probe["capabilities"]
        assert isinstance(capabilities, dict)
        capabilities.pop("excel.pdf-export")

        report = self.report_module.build_report(probe, required_capabilities=["excel.pdf-export"])

        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["summary"]["skipCount"], 1)
        self.assertEqual(report["requirements"]["unmet"], ["excel.pdf-export"])
        self.assertEqual(report["capabilities"]["excel.pdf-export"]["status"], "skip")
        self.assertEqual(report["operations"]["rendered-pdf-evidence"]["readiness"], "unknown")

    def test_malformed_probe_contract_fails_cleanly(self) -> None:
        probe = valid_probe()
        probe.pop("schemaVersion")
        report = self.report_module.build_report(probe, required_capabilities=[])
        self.assertEqual(report["status"], "fail")
        self.assertTrue(any("schemaVersion" in error for error in report["errors"]))

    def test_unknown_required_capability_fails_validation(self) -> None:
        report = self.report_module.build_report(valid_probe(), required_capabilities=["excel.future.magic"])
        self.assertEqual(report["status"], "fail")
        self.assertTrue(any("unknown required capability" in error for error in report["errors"]))

    def test_probe_errors_make_report_incomplete_and_require_pass_exit_one(self) -> None:
        probe = valid_probe()
        cleanup_error = "Owned Excel PID 4321 remained after a successful COM Quit; no process was force-terminated."
        probe["errors"] = [cleanup_error]

        report = self.report_module.build_report(probe, required_capabilities=[])

        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["probeErrors"], [cleanup_error])
        self.assertTrue(any("incomplete probe evidence" in error for error in report["errors"]))

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            probe_json = tmp_dir / "probe-with-cleanup-error.json"
            report_json = tmp_dir / "report.json"
            probe_json.write_text(json.dumps(probe), encoding="utf-8")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(REPORT_SCRIPT),
                    "--probe-json",
                    str(probe_json),
                    "--out-json",
                    str(report_json),
                    "--require-pass",
                ],
                cwd=str(PROJECT_ROOT),
                text=True,
                capture_output=True,
            )
            self.assertEqual(completed.returncode, 1)
            self.assertEqual(json.loads(report_json.read_text(encoding="utf-8"))["status"], "fail")

    def test_invalid_capability_status_fails_contract_validation(self) -> None:
        probe = valid_probe()
        capabilities = probe["capabilities"]
        assert isinstance(capabilities, dict)
        capabilities["excel.com.activation"]["status"] = "maybe"
        report = self.report_module.build_report(probe, required_capabilities=[])
        self.assertEqual(report["status"], "fail")
        self.assertTrue(any("excel.com.activation.status" in error for error in report["errors"]))

    def test_strict_probe_contract_rejects_unknown_and_missing_fields(self) -> None:
        cases = []

        unknown_top = valid_probe()
        unknown_top["future"] = True
        cases.append((unknown_top, "unknown top-level field"))

        unknown_probe = valid_probe()
        unknown_probe["probe"]["future"] = True
        cases.append((unknown_probe, "unknown probe field"))

        unknown_environment = valid_probe()
        unknown_environment["environment"]["future"] = True
        cases.append((unknown_environment, "unknown environment field"))

        unknown_capability = valid_probe()
        unknown_capability["capabilities"]["excel.future.magic"] = capability()
        cases.append((unknown_capability, "unknown capability ID"))

        unknown_capability_field = valid_probe()
        unknown_capability_field["capabilities"]["excel.com.activation"]["future"] = True
        cases.append((unknown_capability_field, "unknown capability field"))

        missing_environment_field = valid_probe()
        missing_environment_field["environment"].pop("excelBuild")
        cases.append((missing_environment_field, "environment.excelBuild is required"))

        invalid_probe_type = valid_probe()
        invalid_probe_type["probe"]["syntheticFixture"] = 1
        cases.append((invalid_probe_type, "probe.syntheticFixture must be a boolean"))

        invalid_environment_type = valid_probe()
        invalid_environment_type["environment"]["powershellVersion"] = 5.1
        cases.append((invalid_environment_type, "environment.powershellVersion must be a string"))

        missing_capability_field = valid_probe()
        missing_capability_field["capabilities"]["excel.com.activation"].pop("error")
        cases.append((missing_capability_field, "capabilities.excel.com.activation.error is required"))

        invalid_capability_type = valid_probe()
        invalid_capability_type["capabilities"]["excel.com.activation"]["error"] = None
        cases.append((invalid_capability_type, "capabilities.excel.com.activation.error must be a string"))

        invalid_timestamp = valid_probe()
        invalid_timestamp["generatedAt"] = "not-a-timestamp"
        cases.append((invalid_timestamp, "generatedAt must be RFC3339"))

        for probe, expected_error in cases:
            with self.subTest(expected_error=expected_error):
                report = self.report_module.build_report(probe, required_capabilities=[])
                self.assertEqual(report["status"], "fail")
                self.assertTrue(any(expected_error in error for error in report["errors"]), report["errors"])

    def test_operation_readiness_distinguishes_blocked_and_user_input(self) -> None:
        blocked_probe = valid_probe()
        blocked_probe["capabilities"]["excel.com.activation"] = capability("fail", "activation")
        blocked = self.report_module.build_report(blocked_probe, required_capabilities=[])
        self.assertEqual(blocked["operations"]["workbook-automation"]["readiness"], "blocked")

        supported = self.report_module.build_report(valid_probe(), required_capabilities=[])
        self.assertEqual(supported["operations"]["adomd-endpoint-query"]["readiness"], "requires-user-input")

    def test_cli_require_pass_returns_one_for_unmet_requirement(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            probe = valid_probe()
            capabilities = probe["capabilities"]
            assert isinstance(capabilities, dict)
            capabilities["excel.vba.project-access"] = capability("fail", "activation")
            probe_json = tmp_dir / "probe.json"
            out_json = tmp_dir / "report.json"
            out_md = tmp_dir / "report.md"
            probe_json.write_text(json.dumps(probe), encoding="utf-8")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(REPORT_SCRIPT),
                    "--probe-json",
                    str(probe_json),
                    "--require-capability",
                    "excel.vba.project-access",
                    "--out-json",
                    str(out_json),
                    "--out-md",
                    str(out_md),
                    "--require-pass",
                ],
                cwd=str(PROJECT_ROOT),
                text=True,
                capture_output=True,
            )
            self.assertEqual(completed.returncode, 1)
            self.assertEqual(json.loads(out_json.read_text(encoding="utf-8"))["status"], "fail")
            self.assertIn("Excel Compatibility Report", out_md.read_text(encoding="utf-8"))

    def test_fixture_generator_runs_all_four_contract_cases_and_matches_manifest(self) -> None:
        fixture_module = load_module(FIXTURE_SCRIPT, "create_excel_capability_fixture")
        with tempfile.TemporaryDirectory() as tmp:
            fixture_dir = Path(tmp)
            manifest = fixture_module.create_fixture(fixture_dir)
            stored_manifest = json.loads(
                (fixture_dir / "excel_capability_fixture_manifest.json").read_text(encoding="utf-8")
            )
            self.assertEqual(stored_manifest, manifest)
            self.assertEqual(set(manifest["cases"]), {"allSupported", "coreBlocked", "partialEvidence", "malformedContract"})
            reports = {}
            for case_id, path_text in manifest["cases"].items():
                path = Path(path_text)
                self.assertTrue(path.is_file())
                reports[case_id] = self.report_module.build_report(json.loads(path.read_text(encoding="utf-8")), [])

            expected = manifest["expected"]
            self.assertEqual(reports["allSupported"]["status"], expected["allSupported"]["status"])
            self.assertEqual(reports["allSupported"]["summary"]["passCount"], expected["allSupported"]["passCount"])
            self.assertEqual(reports["allSupported"]["operations"]["workbook-automation"]["readiness"], expected["allSupported"]["workbookAutomation"])
            self.assertEqual(reports["allSupported"]["operations"]["adomd-endpoint-query"]["readiness"], expected["allSupported"]["adomdEndpointQuery"])
            self.assertEqual(reports["coreBlocked"]["status"], expected["coreBlocked"]["status"])
            self.assertEqual(reports["coreBlocked"]["operations"]["workbook-automation"]["readiness"], expected["coreBlocked"]["workbookAutomation"])
            self.assertEqual(reports["partialEvidence"]["status"], expected["partialEvidence"]["status"])
            self.assertEqual(reports["partialEvidence"]["operations"]["workbook-automation"]["readiness"], expected["partialEvidence"]["workbookAutomation"])
            self.assertEqual(reports["malformedContract"]["status"], expected["malformedContract"]["status"])
            self.assertGreaterEqual(len(reports["malformedContract"]["errors"]), expected["malformedContract"]["minimumErrorCount"])

    def test_schema_files_are_valid_json_and_pin_contract_kind(self) -> None:
        expected = {
            "excel-capability-probe.schema.json": "excel-capability-probe",
            "excel-compatibility-report.schema.json": "excel-compatibility-report",
        }
        for name, kind in expected.items():
            schema = json.loads((PROJECT_ROOT / "schemas" / name).read_text(encoding="utf-8"))
            self.assertEqual(schema["$schema"], "https://json-schema.org/draft/2020-12/schema")
            self.assertEqual(schema["properties"]["kind"]["const"], kind)

    def test_generated_reports_satisfy_fixed_report_schema_contract(self) -> None:
        schema = json.loads((PROJECT_ROOT / "schemas" / "excel-compatibility-report.schema.json").read_text(encoding="utf-8"))
        capability_schema = schema["properties"]["capabilities"]
        operation_schema = schema["properties"]["operations"]
        self.assertEqual(set(capability_schema["required"]), CAPABILITY_IDS)
        self.assertEqual(set(capability_schema["properties"]), CAPABILITY_IDS)
        self.assertFalse(capability_schema["additionalProperties"])
        self.assertEqual(set(operation_schema["required"]), OPERATION_IDS)
        self.assertEqual(set(operation_schema["properties"]), OPERATION_IDS)
        self.assertFalse(operation_schema["additionalProperties"])

        fixture_module = load_module(FIXTURE_SCRIPT, "create_excel_capability_fixture_schema")
        for case_id, probe in fixture_module.fixture_cases().items():
            with self.subTest(case_id=case_id):
                report = self.report_module.build_report(copy.deepcopy(probe), [])
                self.assertTrue(set(schema["required"]).issubset(report))
                self.assertEqual(set(report), set(schema["properties"]))
                self.assertEqual(report["schemaVersion"], "1.0")
                self.assertEqual(report["kind"], "excel-compatibility-report")
                self.assertIn(report["status"], {"pass", "fail"})
                self.assertTrue(self.report_module.is_rfc3339(report["generatedAt"]))
                self.assertIsInstance(report["environment"], dict)
                self.assertEqual(
                    set(report["summary"]),
                    {"passCount", "failCount", "skipCount", "errorCount", "requiredCount", "unmetRequiredCount"},
                )
                self.assertTrue(all(type(value) is int and value >= 0 for value in report["summary"].values()))
                self.assertEqual(
                    report["summary"]["passCount"]
                    + report["summary"]["failCount"]
                    + report["summary"]["skipCount"]
                    + report["summary"]["errorCount"],
                    11,
                )
                self.assertEqual(set(report["capabilities"]), CAPABILITY_IDS)
                self.assertEqual(set(report["operations"]), OPERATION_IDS)
                for item in report["capabilities"].values():
                    self.assertEqual(set(item), set(schema["$defs"]["capability"]["properties"]))
                    self.assertIn(item["status"], schema["$defs"]["capability"]["properties"]["status"]["enum"])
                    self.assertIn(item["evidenceLevel"], schema["$defs"]["capability"]["properties"]["evidenceLevel"]["enum"])
                    self.assertTrue(all(isinstance(value, str) for value in item.values()))
                    self.assertTrue(item["detail"])
                for item in report["operations"].values():
                    self.assertEqual(set(item), set(schema["$defs"]["operation"]["properties"]))
                    self.assertIn(item["readiness"], schema["$defs"]["operation"]["properties"]["readiness"]["enum"])
                    self.assertTrue(item["requires"])
                    self.assertTrue(set(item["requires"]).issubset(CAPABILITY_IDS))
                    self.assertTrue(set(item["missing"]).issubset(set(item["requires"])))
                self.assertTrue(report["boundaries"])
                self.assertTrue(all(isinstance(item, str) and item for item in report["boundaries"]))
                self.assertTrue(all(isinstance(item, str) and item for item in report["probeErrors"]))
                self.assertTrue(all(isinstance(item, str) and item for item in report["errors"]))

    def test_probe_script_reuses_provider_probe_and_supports_both_profiles(self) -> None:
        self.assertTrue(PROBE_SCRIPT.is_file(), "PowerShell capability probe is missing")
        source = PROBE_SCRIPT.read_text(encoding="utf-8-sig")
        self.assertIn("probe_excel_bi_providers.ps1", source)
        self.assertIn('ValidateSet("inventory", "runtime")', source)

    def test_probe_only_terminates_its_owned_excel_pid_after_quit_failure(self) -> None:
        source = PROBE_SCRIPT.read_text(encoding="utf-8-sig")
        self.assertNotIn("Stop-NewExcelProcesses", source)
        self.assertNotIn("Get-ExcelProcessIds", source)
        self.assertIn("GetWindowThreadProcessId", source)
        self.assertIn("OwnedProcessId", source)
        self.assertIn("if ($quitFailed", source)
        self.assertIn("Stop-Process -Id $OwnedProcessId", source)


if __name__ == "__main__":
    unittest.main()
