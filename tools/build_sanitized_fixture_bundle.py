#!/usr/bin/env python3
"""Build a safe generic Excel BI fixture bundle.

The bundle is intended for agent training, release-gate smoke tests, and
workflow demonstrations without using customer workbooks. It combines the
package's existing structural fixtures into one generated folder with a
manifest and a short README.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from create_cube_formula_fixture import create_workbook as create_cube_workbook  # noqa: E402
from create_cube_formula_fixture import model_summary  # noqa: E402
from create_external_dependency_fixture import write_fixture as write_external_fixture  # noqa: E402
from create_pure_deliverable_fixture import write_fixture as write_pure_fixture  # noqa: E402
from create_power_query_lineage_fixture import create_fixture as create_power_query_lineage_fixture  # noqa: E402


BUNDLE_VERSION = 2


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def write_json(path: Path, data: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def safe_clean_dir(path: Path) -> None:
    resolved = path.expanduser().resolve()
    home = Path.home().resolve()
    if str(resolved) == resolved.anchor:
        raise ValueError(f"refusing to remove filesystem root: {resolved}")
    if resolved == home:
        raise ValueError(f"refusing to remove user home directory: {resolved}")
    if len(resolved.parts) < 3:
        raise ValueError(f"refusing to remove shallow directory: {resolved}")
    if resolved.exists():
        shutil.rmtree(resolved)


def relative_to_bundle(path: Path, bundle_dir: Path) -> str:
    return path.resolve().relative_to(bundle_dir.resolve()).as_posix()


def build_readme(manifest: dict[str, object]) -> str:
    fixture_lines = []
    for item in manifest.get("fixtures", []):
        if not isinstance(item, dict):
            continue
        target = item.get("workbook") or item.get("queryDirectory") or item.get("directory") or ""
        fixture_lines.append(
            f"- `{item.get('id')}`: `{target}`; purpose: {item.get('purpose')}"
        )
        for evidence in item.get("expectedEvidence", []):
            fixture_lines.append(f"  - expected: {evidence}")

    return "\n".join(
        [
            "# Sanitized Excel BI Fixture Bundle",
            "",
            "This generated bundle contains only generic, structural Excel workbooks.",
            "It is safe for parser, report, and release-gate validation workflows and is not based on customer data.",
            "",
            "## Contents",
            "",
            *fixture_lines,
            "",
            "## Typical Validation",
            "",
            "```powershell",
            "python tools\\inspect_excel_bi_workbook.py \"<bundle>\\cube_formula_fixture.xlsx\" --out-json \"<tmp>\\cube_openxml.json\"",
            "python tools\\inspect_excel_bi_workbook.py \"<bundle>\\external_dependency_fixture.xlsx\" --out-json \"<tmp>\\external_openxml.json\"",
            "python tools\\inspect_excel_bi_workbook.py \"<bundle>\\pure_deliverable_fixture.xlsx\" --out-json \"<tmp>\\pure_openxml.json\"",
            "python tools\\build_power_query_lineage_report.py \"<bundle>\\power_query_lineage\\safe\" --out-json \"<tmp>\\pq_lineage_safe.json\"",
            "python tools\\build_power_query_lineage_report.py \"<bundle>\\power_query_lineage\\risky\" --out-json \"<tmp>\\pq_lineage_risky.json\"",
            "```",
            "",
            "## Boundary",
            "",
            "- These workbooks validate structure, dependency reporting, and cleanup planning.",
            "- They do not prove live Excel calculation, Power Query credentials, or Power Pivot semantic correctness.",
            "- Generate task outputs into temporary or deliverable folders, not into the plugin source tree.",
            "",
        ]
    )


def build_bundle(out_dir: Path, out_json: Path | None = None, out_md: Path | None = None, clean: bool = False) -> dict[str, object]:
    bundle_dir = out_dir.expanduser().resolve()
    if clean:
        safe_clean_dir(bundle_dir)
    bundle_dir.mkdir(parents=True, exist_ok=True)

    generated_at = now_iso()

    cube_workbook = bundle_dir / "cube_formula_fixture.xlsx"
    cube_model_json = bundle_dir / "cube_model_summary.json"
    external_workbook = bundle_dir / "external_dependency_fixture.xlsx"
    external_json = bundle_dir / "external_dependency_fixture.json"
    pure_workbook = bundle_dir / "pure_deliverable_fixture.xlsx"
    pure_json = bundle_dir / "pure_deliverable_fixture.json"
    pq_lineage_dir = bundle_dir / "power_query_lineage"
    pq_lineage_json = bundle_dir / "power_query_lineage_fixture.json"

    create_cube_workbook(cube_workbook)
    cube_summary = model_summary()
    cube_summary["workbookPath"] = str(cube_workbook)
    cube_summary["generatedAt"] = generated_at
    write_json(cube_model_json, cube_summary)

    external_summary = write_external_fixture(external_workbook)
    write_json(external_json, external_summary)

    pure_summary = write_pure_fixture(pure_workbook)
    write_json(pure_json, pure_summary)

    pq_lineage_summary = create_power_query_lineage_fixture(pq_lineage_dir)
    write_json(pq_lineage_json, pq_lineage_summary)

    manifest: dict[str, object] = {
        "bundleVersion": BUNDLE_VERSION,
        "generatedAt": generated_at,
        "fixtureCount": 4,
        "bundleDir": str(bundle_dir),
        "fixtures": [
            {
                "id": "cube-formula",
                "workbook": relative_to_bundle(cube_workbook, bundle_dir),
                "metadata": relative_to_bundle(cube_model_json, bundle_dir),
                "purpose": "CUBE formula, MDX reference, and model-report pipeline smoke tests",
                "expectedEvidence": [
                    "7 CUBE formulas",
                    "2 known measures",
                    "1 intentionally missing measure reference",
                    "hard-coded and dynamic MDX diagnostics",
                ],
            },
            {
                "id": "external-dependency",
                "workbook": relative_to_bundle(external_workbook, bundle_dir),
                "metadata": relative_to_bundle(external_json, bundle_dir),
                "purpose": "External link, connection, credential-indicator, formula, and mashup-like package detection smoke tests",
                "expectedEvidence": [
                    "2 workbook connections",
                    "1 redacted credential-like connection indicator",
                    "1 external link part",
                    "1 external formula",
                    "mashup-like custom XML part",
                ],
            },
            {
                "id": "pure-deliverable",
                "workbook": relative_to_bundle(pure_workbook, bundle_dir),
                "metadata": relative_to_bundle(pure_json, bundle_dir),
                "purpose": "Post-clean pure-xlsx readiness and cleanup assertion smoke tests",
                "expectedEvidence": [
                    "0 workbook connections",
                    "0 external links",
                    "0 external formulas",
                    "no mashup, Data Model, or VBA project markers",
                ],
            },
            {
                "id": "power-query-lineage",
                "queryDirectory": relative_to_bundle(pq_lineage_dir, bundle_dir),
                "metadata": relative_to_bundle(pq_lineage_json, bundle_dir),
                "purpose": "Exported Power Query M dependency, source-risk, and delivery-readiness smoke tests",
                "expectedEvidence": [
                    "safe exported-query set with 4 queries and 0 findings",
                    "risky exported-query set with 11 queries and source-risk findings",
                    "native SQL, credential-like literal, mixed-source lineage, and query cycle diagnostics",
                ],
            },
        ],
        "usage": [
            "Use this bundle for parser/report/release-gate validation when no safe customer workbook can be used.",
            "Use Windows Excel COM separately when a workflow must prove live refresh, VBA execution, or Data Model semantics.",
        ],
        "limitations": [
            "The CUBE workbook has structural formulas but no live Power Pivot model.",
            "The external-dependency workbook contains safe placeholder dependencies only.",
            "The pure-deliverable workbook proves a clean package shape, not a cleaned copy of a customer workbook.",
            "The Power Query lineage fixture contains exported M source text only; it does not refresh real sources.",
        ],
    }

    manifest_path = (out_json.expanduser().resolve() if out_json else bundle_dir / "fixture-bundle.json")
    readme_path = (out_md.expanduser().resolve() if out_md else bundle_dir / "README.md")
    manifest["manifestPath"] = str(manifest_path)
    manifest["readmePath"] = str(readme_path)
    write_json(manifest_path, manifest)
    readme_path.parent.mkdir(parents=True, exist_ok=True)
    readme_path.write_text(build_readme(manifest), encoding="utf-8")
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out-dir", required=True, type=Path, help="Output folder for generated fixture bundle")
    parser.add_argument("--clean", action="store_true", help="Remove the output folder before generating")
    parser.add_argument("--out-json", type=Path, help="Optional manifest JSON path; defaults to <out-dir>/fixture-bundle.json")
    parser.add_argument("--out-md", type=Path, help="Optional README Markdown path; defaults to <out-dir>/README.md")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = build_bundle(args.out_dir, out_json=args.out_json, out_md=args.out_md, clean=args.clean)
    print(json.dumps({"bundleDir": manifest["bundleDir"], "fixtureCount": manifest["fixtureCount"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
