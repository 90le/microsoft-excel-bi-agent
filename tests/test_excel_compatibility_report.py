from __future__ import annotations

import importlib.util
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

    def test_fixture_generator_emits_four_contract_cases(self) -> None:
        fixture_module = load_module(FIXTURE_SCRIPT, "create_excel_capability_fixture")
        with tempfile.TemporaryDirectory() as tmp:
            manifest = fixture_module.create_fixture(Path(tmp))
            self.assertEqual(set(manifest["cases"]), {"allSupported", "coreBlocked", "partialEvidence", "malformedContract"})
            for path_text in manifest["cases"].values():
                self.assertTrue(Path(path_text).is_file())

    def test_schema_files_are_valid_json_and_pin_contract_kind(self) -> None:
        expected = {
            "excel-capability-probe.schema.json": "excel-capability-probe",
            "excel-compatibility-report.schema.json": "excel-compatibility-report",
        }
        for name, kind in expected.items():
            schema = json.loads((PROJECT_ROOT / "schemas" / name).read_text(encoding="utf-8"))
            self.assertEqual(schema["$schema"], "https://json-schema.org/draft/2020-12/schema")
            self.assertEqual(schema["properties"]["kind"]["const"], kind)

    def test_probe_script_reuses_provider_probe_and_supports_both_profiles(self) -> None:
        self.assertTrue(PROBE_SCRIPT.is_file(), "PowerShell capability probe is missing")
        source = PROBE_SCRIPT.read_text(encoding="utf-8-sig")
        self.assertIn("probe_excel_bi_providers.ps1", source)
        self.assertIn('ValidateSet("inventory", "runtime")', source)


if __name__ == "__main__":
    unittest.main()
