from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
VALIDATOR_PATH = PROJECT_ROOT / "tools" / "validate_observed_usage.py"
SUMMARIZER_PATH = PROJECT_ROOT / "tools" / "summarize_observed_usage.py"


def valid_event() -> dict[str, object]:
    return {
        "schemaVersion": "1.0",
        "eventId": "event-001",
        "recordedAt": "2026-07-14T09:30:00Z",
        "caseId": "sanitized-case-001",
        "requestedSkill": "excel-bi-router",
        "selectedSkill": "power-query-m-engineering",
        "outcome": "completed",
        "durationMs": 25,
        "evidenceLevel": "structural",
        "evidenceBoundary": "local-sanitized-metadata",
    }


class ObservedUsageValidatorTests(unittest.TestCase):
    def run_validator_path(self, path: Path) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(VALIDATOR_PATH), str(path)],
            cwd=str(PROJECT_ROOT),
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
        )

    def run_validator(self, lines: list[str]) -> subprocess.CompletedProcess[str]:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "events.jsonl"
            path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            return self.run_validator_path(path)

    def test_accepts_valid_jsonl_event(self) -> None:
        result = self.run_validator([json.dumps(valid_event())])

        self.assertEqual(0, result.returncode, result.stderr)
        self.assertIn("validation passed", result.stdout.lower())

    def test_rejects_malformed_json(self) -> None:
        result = self.run_validator(["{"])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("events.jsonl:1", result.stderr)
        self.assertIn("malformed JSON", result.stderr)

    def test_rejects_missing_required_field(self) -> None:
        event = valid_event()
        del event["caseId"]
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("caseId", result.stderr)

    def test_rejects_unsupported_evidence_level(self) -> None:
        event = valid_event()
        event["evidenceLevel"] = "telemetry"
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("evidenceLevel", result.stderr)
        self.assertIn("unsupported", result.stderr)

    def test_rejects_negative_duration(self) -> None:
        event = valid_event()
        event["durationMs"] = -1
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("durationMs", result.stderr)

    def test_rejects_duplicate_event_id(self) -> None:
        event = valid_event()
        result = self.run_validator([json.dumps(event), json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("events.jsonl:2", result.stderr)
        self.assertIn("duplicate eventId", result.stderr)

    def test_rejects_invalid_timestamp(self) -> None:
        event = valid_event()
        event["recordedAt"] = "2026-07-14T09:30:00+08:00"
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("recordedAt", result.stderr)
        self.assertIn("UTC ISO-8601", result.stderr)

    def test_rejects_path_in_any_string(self) -> None:
        event = valid_event()
        event["outcome"] = "saved C:\\customer\\report"
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("outcome", result.stderr)
        self.assertIn("path-like", result.stderr)

    def test_rejects_windows_drive_relative_path(self) -> None:
        for value in ("C:customer-data", "D:report"):
            with self.subTest(value=value):
                event = valid_event()
                event["outcome"] = value
                result = self.run_validator([json.dumps(event)])

                self.assertNotEqual(0, result.returncode)
                self.assertIn("outcome", result.stderr)
                self.assertIn("path-like", result.stderr)

    def test_rejects_credential_like_value_in_nested_string(self) -> None:
        event = valid_event()
        event["details"] = {"note": "api_key=super-secret-value"}
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("details.note", result.stderr)
        self.assertIn("credential-like", result.stderr)

    def test_rejects_path_like_object_key(self) -> None:
        event = valid_event()
        event["details"] = {r"C:\\customer\\report": "safe"}
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("details.<key>", result.stderr)
        self.assertIn("path-like", result.stderr)

    def test_rejects_credential_like_object_key(self) -> None:
        event = valid_event()
        event["details"] = {"api_key": "safe"}
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("details.<key>", result.stderr)
        self.assertIn("credential-like", result.stderr)

    def test_redacts_sensitive_key_from_descendant_diagnostics(self) -> None:
        event = valid_event()
        event["details"] = {"api_key": {"child": "C:customer-data"}}
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("details.<key>.child", result.stderr)
        self.assertNotIn("details.api_key.child", result.stderr)

    def test_rejects_representative_credential_strings(self) -> None:
        credentials = (
            "sk-abcdefghijklmnopqrstuvwxyz123456",
            "ghp_abcdefghijklmnopqrstuvwxyz1234567890",
            "xoxb-REDACTED-PLACEHOLDER",
            "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.signaturevalue",
            "token super-secret-value",
        )
        for credential in credentials:
            with self.subTest(credential=credential):
                event = valid_event()
                event["details"] = {"note": credential}
                result = self.run_validator([json.dumps(event)])

                self.assertNotEqual(0, result.returncode)
                self.assertIn("details.note", result.stderr)
                self.assertIn("credential-like", result.stderr)

    def test_rejects_customer_artifact_name_in_any_string(self) -> None:
        event = valid_event()
        event["details"] = {"source": "CustomerSales.xlsx"}
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("details.source", result.stderr)
        self.assertIn("customer artifact", result.stderr)

    def test_rejects_workbook_behavior_without_sanitized_boundary(self) -> None:
        event = valid_event()
        event["evidenceLevel"] = "workbook-behavior"
        event["evidenceBoundary"] = "local-sanitized-metadata"
        result = self.run_validator([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertIn("evidenceBoundary", result.stderr)
        self.assertIn("local-user-supplied-sanitized", result.stderr)

    def test_diagnostics_include_resolved_input_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "events.jsonl"
            path.write_text("{\n", encoding="utf-8")
            result = self.run_validator_path(path)

            self.assertNotEqual(0, result.returncode)
            self.assertIn(str(path.resolve()), result.stderr)

    def test_rejects_invalid_utf8_with_actionable_diagnostic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "events.jsonl"
            path.write_bytes(b"\xff\xfe")
            result = self.run_validator_path(path)

            self.assertNotEqual(0, result.returncode)
            self.assertIn(str(path.resolve()), result.stderr)
            self.assertIn("UTF-8", result.stderr)


class ObservedUsageSummarizerTests(unittest.TestCase):
    def run_summarizer(self, lines: list[str]) -> subprocess.CompletedProcess[str]:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "events.jsonl"
            path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            return subprocess.run(
                [sys.executable, str(SUMMARIZER_PATH), str(path)],
                cwd=str(PROJECT_ROOT),
                text=True,
                encoding="utf-8",
                errors="replace",
                capture_output=True,
            )

    def test_summarizes_valid_events_without_exposing_event_details(self) -> None:
        events = []
        for index, (requested, selected, outcome, duration, evidence) in enumerate(
            (
                ("excel-bi-router", "power-query-m-engineering", "completed", 10, "structural"),
                ("excel-bi-router", "power-query-m-engineering", "completed", 20, "runtime-capability"),
                ("power-query-m-engineering", "power-query-m-engineering", "failed", 30, "workbook-behavior"),
                ("excel-bi-router", "excel-vba-workbook-engineering", "completed", 40, "structural"),
            ),
            start=1,
        ):
            event = valid_event()
            event.update(
                {
                    "eventId": f"event-{index:03d}",
                    "requestedSkill": requested,
                    "selectedSkill": selected,
                    "outcome": outcome,
                    "durationMs": duration,
                    "evidenceLevel": evidence,
                    "evidenceBoundary": (
                        "local-user-supplied-sanitized"
                        if evidence == "workbook-behavior"
                        else "local-sanitized-metadata"
                    ),
                    "details": {"note": "do-not-report-this-note"},
                }
            )
            events.append(event)

        result = self.run_summarizer([json.dumps(event) for event in events])

        self.assertEqual(0, result.returncode, result.stderr)
        self.assertEqual(
            {
                "eventCount": 4,
                "requestedSkillCounts": {
                    "excel-bi-router": 3,
                    "power-query-m-engineering": 1,
                },
                "selectedSkillCounts": {
                    "excel-vba-workbook-engineering": 1,
                    "power-query-m-engineering": 3,
                },
                "outcomeCounts": {"completed": 3, "failed": 1},
                "evidenceLevelCounts": {
                    "runtime-capability": 1,
                    "structural": 2,
                    "workbook-behavior": 1,
                },
                "durationMs": {"median": 25, "max": 40},
            },
            json.loads(result.stdout),
        )
        self.assertNotIn("do-not-report-this-note", result.stdout)
        self.assertNotIn("events.jsonl", result.stdout)

    def test_rejects_invalid_events_before_aggregation(self) -> None:
        event = valid_event()
        event["durationMs"] = -1

        result = self.run_summarizer([json.dumps(event)])

        self.assertNotEqual(0, result.returncode)
        self.assertEqual("", result.stdout)
        self.assertIn("validation failed", result.stderr.lower())


if __name__ == "__main__":
    unittest.main()
