from __future__ import annotations

import hashlib
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
BUILDER_PATH = PROJECT_ROOT / "tools" / "build_runtime_package.py"
DEPLOY_PATH = PROJECT_ROOT / "tools" / "deploy-local-plugin.py"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


runtime_package = load_module(BUILDER_PATH, "test_build_runtime_package")
deploy_plugin = load_module(DEPLOY_PATH, "test_deploy_local_plugin")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_minimal_project(root: Path, *, missing_reference: bool = False, forbidden: bool = False) -> None:
    (root / ".codex-plugin").mkdir(parents=True)
    (root / ".codex-plugin" / "plugin.json").write_text(
        json.dumps(
            {
                "name": "runtime-test-plugin",
                "version": "0.0.1",
                "skills": "./skills/",
                "interface": {"displayName": "Runtime test"},
            }
        ),
        encoding="utf-8",
    )
    skill_dir = root / "skills" / "demo"
    skill_dir.mkdir(parents=True)
    reference = "\nUse `tools/missing_runtime_tool.py`." if missing_reference else ""
    (skill_dir / "SKILL.md").write_text(
        f"---\nname: demo\ndescription: Runtime package fixture.\n---\n# Demo{reference}\n",
        encoding="utf-8",
    )
    (root / "LICENSE").write_text("MIT\n", encoding="utf-8")
    if forbidden:
        (skill_dir / "customer-workbook.xlsx").write_bytes(b"private fixture")


class RuntimePackageTests(unittest.TestCase):
    def test_allowlist_manifest_hashes_and_mirror_status(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "runtime"
            manifest = runtime_package.build_runtime_package(PROJECT_ROOT, out_dir)

            self.assertEqual("pass", manifest["status"])
            paths = [item["path"] for item in manifest["files"]]
            self.assertEqual(sorted(paths), paths)
            self.assertIn(".codex-plugin/plugin.json", paths)
            self.assertIn("LICENSE", paths)
            self.assertIn("README.md", paths)
            self.assertTrue(any(path.startswith("skills/") and path.endswith("/SKILL.md") for path in paths))
            self.assertTrue(any(path.startswith("tools/") for path in paths))
            self.assertTrue(any(path.startswith("fixtures/") for path in paths))

            forbidden_prefixes = (
                ".agents/",
                ".claude/",
                ".opencode/",
                ".git/",
                ".github/",
                "docs/",
                "prompts/",
            )
            self.assertFalse(any(path.startswith(forbidden_prefixes) for path in paths))
            self.assertLess((out_dir / "README.md").stat().st_size, 5000)

            actual_payload = sorted(
                path.relative_to(out_dir).as_posix()
                for path in out_dir.rglob("*")
                if path.is_file() and path.name != "runtime-package-manifest.json"
            )
            self.assertEqual(paths, actual_payload)
            self.assertEqual(sum(item["size"] for item in manifest["files"]), manifest["totalBytes"])
            for item in manifest["files"]:
                packaged = out_dir / item["path"]
                self.assertEqual(packaged.stat().st_size, item["size"])
                self.assertEqual(sha256(packaged), item["sha256"])

            self.assertIn(manifest["skillMirror"]["status"], {"in-sync", "drift", "not-checked"})
            if manifest["skillMirror"]["status"] == "drift":
                self.assertTrue(any("drift" in warning.lower() for warning in manifest["warnings"]))

    def test_zip_is_deterministic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first_zip = root / "first.zip"
            second_zip = root / "second.zip"
            first_manifest = runtime_package.build_runtime_package(PROJECT_ROOT, root / "first", zip_path=first_zip)
            second_manifest = runtime_package.build_runtime_package(PROJECT_ROOT, root / "second", zip_path=second_zip)

            self.assertEqual(first_zip.read_bytes(), second_zip.read_bytes())
            self.assertEqual(first_manifest["sourceBytes"], second_manifest["sourceBytes"])

    def test_python_dependency_closure_and_help_smoke(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "runtime"
            manifest = runtime_package.build_runtime_package(PROJECT_ROOT, out_dir)
            paths = {item["path"] for item in manifest["files"]}

            self.assertIn("tools/mdx_references.py", paths)
            self.assertIn("tools/create_external_dependency_fixture.py", paths)
            self.assertIn("tools/create_pure_deliverable_fixture.py", paths)
            self.assertIn("tools/create_power_query_lineage_fixture.py", paths)

            packaged_python = sorted(path for path in paths if path.startswith("tools/") and path.endswith(".py"))
            smoke = manifest["pythonToolSmoke"]
            self.assertEqual(packaged_python, [item["path"] for item in smoke])
            self.assertTrue(all(item["status"] == "pass" for item in smoke), smoke)

    def test_require_pass_rejects_python_tool_with_missing_dependency(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            write_minimal_project(root)
            tools_dir = root / "tools"
            tools_dir.mkdir()
            (tools_dir / "entry.py").write_text(
                "import definitely_missing_runtime_dependency\n",
                encoding="utf-8",
            )
            skill = root / "skills" / "demo" / "SKILL.md"
            skill.write_text(skill.read_text(encoding="utf-8") + "\nUse `tools/entry.py`.\n", encoding="utf-8")
            out_dir = Path(tmp) / "runtime"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(BUILDER_PATH),
                    "--project-root",
                    str(root),
                    "--out-dir",
                    str(out_dir),
                    "--require-pass",
                ],
                cwd=str(PROJECT_ROOT),
                text=True,
                capture_output=True,
            )

            self.assertEqual(1, completed.returncode)
            manifest = json.loads((out_dir / "runtime-package-manifest.json").read_text(encoding="utf-8"))
            self.assertEqual("fail", manifest["status"])
            self.assertEqual("fail", manifest["pythonToolSmoke"][0]["status"])
            self.assertTrue(any("entry.py" in error for error in manifest["errors"]))

    def test_output_directory_inside_project_is_rejected_without_deleting_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            write_minimal_project(root)
            protected = root / "tools" / "keep-source.txt"
            protected.parent.mkdir()
            protected.write_text("must survive\n", encoding="utf-8")

            for unsafe in [root, root.parent, root / "tools" / "runtime"]:
                with self.subTest(unsafe=unsafe), self.assertRaises(ValueError):
                    runtime_package.prepare_output_dir(root, unsafe)

            self.assertEqual("must survive\n", protected.read_text(encoding="utf-8"))

    def test_zip_inside_project_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            write_minimal_project(root)
            out_dir = Path(tmp) / "runtime"
            zip_path = root / "artifacts" / "runtime.zip"

            for unsafe in [root, root.parent, zip_path]:
                with self.subTest(unsafe=unsafe), self.assertRaises(ValueError):
                    runtime_package.validate_zip_path(root, out_dir, unsafe)

            self.assertFalse(zip_path.exists())

    def test_fixture_allowlist_excludes_private_and_unstructured_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            write_minimal_project(root)
            safe_fixture = root / "fixtures" / "real-sanitized-cases" / "manifest.json"
            safe_fixture.parent.mkdir(parents=True)
            safe_fixture.write_text('{"synthetic": true}\n', encoding="utf-8")
            referenced_fixture = root / "fixtures" / "synthetic" / "Sample.bas"
            referenced_fixture.parent.mkdir()
            referenced_fixture.write_text("Attribute VB_Name = \"Sample\"\n", encoding="utf-8")
            skill = root / "skills" / "demo" / "SKILL.md"
            skill.write_text(
                skill.read_text(encoding="utf-8") + "\nUse `fixtures/synthetic/Sample.bas`.\n",
                encoding="utf-8",
            )
            (root / "fixtures" / ".env").write_text("CUSTOMER_SECRET=1\n", encoding="utf-8")
            (root / "fixtures" / "customers.csv").write_text("name\nprivate\n", encoding="utf-8")
            misc_fixture = root / "fixtures" / "misc" / "opaque.bin"
            misc_fixture.parent.mkdir()
            misc_fixture.write_bytes(b"opaque")
            out_dir = Path(tmp) / "runtime"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(BUILDER_PATH),
                    "--project-root",
                    str(root),
                    "--out-dir",
                    str(out_dir),
                    "--require-pass",
                ],
                cwd=str(PROJECT_ROOT),
                text=True,
                capture_output=True,
            )

            self.assertEqual(1, completed.returncode)
            manifest = json.loads((out_dir / "runtime-package-manifest.json").read_text(encoding="utf-8"))
            paths = {item["path"] for item in manifest["files"]}
            self.assertIn("fixtures/real-sanitized-cases/manifest.json", paths)
            self.assertIn("fixtures/synthetic/Sample.bas", paths)
            self.assertNotIn("fixtures/.env", paths)
            self.assertNotIn("fixtures/customers.csv", paths)
            self.assertNotIn("fixtures/misc/opaque.bin", paths)
            self.assertTrue(any("fixtures/.env" in item for item in manifest["forbiddenFiles"]))
            self.assertTrue(any("fixtures/customers.csv" in item for item in manifest["forbiddenFiles"]))
            self.assertTrue(any("opaque.bin" in warning for warning in manifest["warnings"]))

    def test_require_pass_rejects_unresolved_skill_reference(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            write_minimal_project(root, missing_reference=True)
            out_dir = Path(tmp) / "runtime"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(BUILDER_PATH),
                    "--project-root",
                    str(root),
                    "--out-dir",
                    str(out_dir),
                    "--require-pass",
                ],
                cwd=str(PROJECT_ROOT),
                text=True,
                capture_output=True,
            )

            self.assertEqual(1, completed.returncode)
            manifest = json.loads((out_dir / "runtime-package-manifest.json").read_text(encoding="utf-8"))
            self.assertEqual("fail", manifest["status"])
            self.assertIn("tools/missing_runtime_tool.py", manifest["unresolvedReferences"])

    def test_require_pass_rejects_forbidden_payload_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            write_minimal_project(root, forbidden=True)
            out_dir = Path(tmp) / "runtime"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(BUILDER_PATH),
                    "--project-root",
                    str(root),
                    "--out-dir",
                    str(out_dir),
                    "--require-pass",
                ],
                cwd=str(PROJECT_ROOT),
                text=True,
                capture_output=True,
            )

            self.assertEqual(1, completed.returncode)
            manifest = json.loads((out_dir / "runtime-package-manifest.json").read_text(encoding="utf-8"))
            self.assertEqual("fail", manifest["status"])
            self.assertTrue(any("customer-workbook.xlsx" in item for item in manifest["forbiddenFiles"]))

    def test_local_deployment_consumes_runtime_staging(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            destination = Path(tmp) / "plugins" / "microsoft-excel-bi-agent-pack"
            deploy_plugin.deploy_runtime_plugin(PROJECT_ROOT, destination, replace=True)

            self.assertTrue((destination / "runtime-package-manifest.json").is_file())
            self.assertTrue((destination / ".codex-plugin" / "plugin.json").is_file())
            self.assertTrue((destination / "skills").is_dir())
            self.assertFalse((destination / ".agents").exists())
            self.assertFalse((destination / ".claude").exists())
            self.assertFalse((destination / ".opencode").exists())
            self.assertFalse((destination / "docs").exists())

    def test_installer_check_builds_runtime_staging(self) -> None:
        completed = subprocess.run(
            ["node", str(PROJECT_ROOT / "tools" / "install.mjs"), "--check"],
            cwd=str(PROJECT_ROOT),
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
        )

        self.assertEqual(0, completed.returncode, completed.stderr or completed.stdout)
        self.assertIn("tools/build_runtime_package.py", completed.stdout)


if __name__ == "__main__":
    unittest.main()
