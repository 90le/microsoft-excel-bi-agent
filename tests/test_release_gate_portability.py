from __future__ import annotations

import importlib.util
import json
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RELEASE_GATE_PATH = PROJECT_ROOT / "tools" / "run_release_gate.py"


def load_release_gate():
    spec = importlib.util.spec_from_file_location("run_release_gate", RELEASE_GATE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {RELEASE_GATE_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


release_gate = load_release_gate()


class ReleaseGatePortabilityTests(unittest.TestCase):
    def test_run_command_decodes_utf8_output_on_localized_windows(self) -> None:
        message = "PowerShell 语法通过 ✓ 🚀"
        command = [
            sys.executable,
            "-c",
            f"import sys; sys.stdout.buffer.write({message!r}.encode('utf-8'))",
        ]

        result = release_gate.run_command(command, PROJECT_ROOT, "localized output")

        self.assertEqual(release_gate.PASS, result.status)
        self.assertEqual(message, result.stdout)
        self.assertIsInstance(result.stderr, str)

    def test_find_bash_rejects_system32_wsl_launcher(self) -> None:
        system_bash = os.path.normcase(r"C:\Windows\System32\bash.exe")
        git_bash = os.path.normcase(r"C:\Program Files\Git\bin\bash.exe")

        def fake_which(name: str) -> str | None:
            return system_bash if name == "bash" else None

        with patch.object(release_gate.os, "name", "nt"), patch.object(
            release_gate.shutil, "which", side_effect=fake_which
        ), patch.object(release_gate.Path, "exists", return_value=True):
            selected = release_gate.find_bash()

        self.assertEqual(git_bash, os.path.normcase(str(selected)))

    def test_manifest_has_public_urls_and_three_high_value_prompts(self) -> None:
        manifest = json.loads(
            (PROJECT_ROOT / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8")
        )
        interface = manifest.get("interface", {})

        self.assertEqual(
            "https://90le.github.io/microsoft-excel-bi-agent/intro.html",
            interface.get("websiteURL"),
        )
        self.assertEqual(
            "https://github.com/90le/microsoft-excel-bi-agent/blob/main/docs/privacy-policy.md",
            interface.get("privacyPolicyURL"),
        )
        self.assertEqual(
            "https://github.com/90le/microsoft-excel-bi-agent/blob/main/docs/terms-of-service.md",
            interface.get("termsOfServiceURL"),
        )
        self.assertTrue((PROJECT_ROOT / "docs" / "privacy-policy.md").is_file())
        self.assertTrue((PROJECT_ROOT / "docs" / "terms-of-service.md").is_file())
        prompts = interface.get("defaultPrompt", [])
        self.assertEqual(3, len(prompts))
        for expected_term in ("Inspect", "Diagnose", "client-ready"):
            self.assertTrue(
                any(expected_term in prompt for prompt in prompts),
                f"defaultPrompt must cover {expected_term}",
            )

    def test_windows_ci_runs_structural_gate_from_chinese_space_path(self) -> None:
        workflow = (PROJECT_ROOT / ".github" / "workflows" / "validate.yml").read_text(
            encoding="utf-8"
        )

        self.assertIn("runs-on: windows-latest", workflow)
        self.assertIn("中文 空格路径", workflow)
        self.assertIn("run_release_gate.py --project-root . --profile structural", workflow)


if __name__ == "__main__":
    unittest.main()
