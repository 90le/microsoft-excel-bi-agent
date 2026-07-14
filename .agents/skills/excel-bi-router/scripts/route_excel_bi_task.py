#!/usr/bin/env python3
"""Route an Excel BI task to the right skill and validation path.

This script is deterministic and conservative. It does not inspect workbooks or
prove runtime behavior; it converts a plain-language task description into a
reviewable first-pass routing decision.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


PASS = "pass"
WARN = "warn"


@dataclass(frozen=True)
class Route:
    layer: str
    skill: str
    validation_needed: tuple[str, ...]
    scripts: tuple[str, ...]
    boundaries: tuple[str, ...]
    keywords: tuple[str, ...]


ROUTES: tuple[Route, ...] = (
    Route(
        layer="Workbook/VBA",
        skill="excel-vba-workbook-engineering",
        validation_needed=(
            "Windows desktop Excel for VBA import, macro execution, buttons, refresh-triggered events, and final workbook behavior.",
            "OpenXML/static inspection is acceptable for formulas, links, workbook structure, and non-runtime triage.",
        ),
        scripts=(
            ".agents/skills/excel-vba-workbook-engineering/scripts/inspect_workbook.ps1",
            ".agents/skills/excel-vba-workbook-engineering/scripts/export_vba.ps1",
            ".agents/skills/excel-vba-workbook-engineering/scripts/import_vba.ps1",
            "tools/build_formula_quality_report.py",
            "tools/build_workbook_controls_report.py",
        ),
        boundaries=(
            "Do not claim Linux or macOS validated VBA runtime behavior.",
            "Preserve vbaProject.bin when editing macro-enabled workbooks structurally.",
        ),
        keywords=(
            ".xlsm",
            ".xlsx",
            "vba",
            "macro",
            "button",
            "onaction",
            "shape",
            "form control",
            "worksheet_change",
            "hidden sheet",
            "very hidden",
            "formula",
            "defined name",
            "pivot table",
            "workbook",
            "excel com",
            "macro-enabled",
        ),
    ),
    Route(
        layer="Power Query M",
        skill="power-query-m-engineering",
        validation_needed=(
            "Windows desktop Excel or Power BI for live refresh, load targets, credential prompts, and query completion.",
            "Static M lint and lineage reports can run cross-platform before import or refresh.",
        ),
        scripts=(
            ".agents/skills/power-query-m-engineering/scripts/manage_power_queries_excel_com.ps1",
            ".agents/skills/power-query-m-engineering/scripts/refresh_power_queries_excel_com.ps1",
            ".agents/skills/power-query-m-engineering/scripts/build_power_query_refresh_report.py",
            ".agents/skills/power-query-m-engineering/scripts/classify_power_query_refresh_errors.py",
            "tools/build_power_query_lineage_report.py",
        ),
        boundaries=(
            "Do not repair credential or privacy-firewall failures with silent retry loops.",
            "Do not treat WorkbookQuery.Refresh alone as proof that a loaded worksheet table updated.",
        ),
        keywords=(
            "power query",
            "powerquery",
            "query editor",
            "m code",
            "m formula",
            "table.",
            "excel.workbook",
            "folder.files",
            "csv.document",
            "refresh",
            "formula.firewall",
            "privacy",
            "query folding",
            "changed type",
            "missingfield.usenull",
        ),
    ),
    Route(
        layer="Power Pivot DAX",
        skill="power-pivot-dax-modeling",
        validation_needed=(
            "Excel Power Pivot or compatible model host for measure evaluation, relationships, and filter context.",
            "Static DAX compatibility and dependency lint can run before editing model formulas.",
        ),
        scripts=(
            ".agents/skills/power-pivot-dax-modeling/scripts/lint_dax_compat.py",
            ".agents/skills/power-pivot-dax-modeling/scripts/analyze_dax_dependencies.py",
            "tools/build_excel_bi_model_report.py",
            "tools/analyze_measure_rename_impact.py",
            "tools/build_measure_rename_rewrite_plan.py",
        ),
        boundaries=(
            "Excel Power Pivot compatibility is not identical to modern Power BI DAX.",
            "Static DAX lint does not prove relationship grain or measure results.",
        ),
        keywords=(
            "dax compatibility",
            "power pivot compatibility",
            "excel power pivot compatibility",
            "power pivot",
            "data model",
            "dax",
            "measure",
            "calculate",
            "filter context",
            "relationship",
            "removefilters",
            "selectedvalue",
            "divide",
            "summarize",
        ),
    ),
    Route(
        layer="MDX/CUBE",
        skill="mdx-cubevalue-extraction",
        validation_needed=(
            "Excel formula inspection for CUBE formulas and helper-cell dependencies.",
            "ADOMD endpoint execution only when a real connection string and MDX query are explicitly supplied.",
        ),
        scripts=(
            "tools/build_cube_dependency_report.py",
            "tools/create_cube_formula_fixture.py",
            "tools/analyze_measure_rename_impact.py",
            "tools/build_measure_rename_rewrite_plan.py",
            "tools/test_excel_adomd_query.ps1",
        ),
        boundaries=(
            "CUBE formula parsing does not prove the cube endpoint can answer the query.",
            "When CUBEVALUE reads ThisWorkbookDataModel, inspect report-layer formulas before editing DAX.",
        ),
        keywords=(
            "cubevalue",
            "cubemember",
            "cubeset",
            "mdx",
            "thisworkbookdatamodel",
            "[measures]",
            "member",
            "tuple",
            "olap",
            "cube formula",
        ),
    ),
    Route(
        layer="ADO/SQL",
        skill="excel-ado-sql-data-access",
        validation_needed=(
            "Windows provider probe for ACE/OLEDB/MSOLAP/ADODB/ADOMD availability.",
            "A real endpoint or workbook fixture is required to prove SQL or MDX query execution.",
        ),
        scripts=(
            "tools/probe_excel_bi_providers.ps1",
            "tools/build_provider_environment_report.py",
            "tools/test_excel_ado_sql_access.ps1",
            "tools/test_excel_adomd_query.ps1",
        ),
        boundaries=(
            "Do not assume ACE, MSOLAP, ADODB, or ADOMD is installed without provider evidence.",
            "Connection-string examples must not include secrets or customer credentials.",
        ),
        keywords=(
            "ado",
            "adodb",
            "adomd",
            "oledb",
            "ole db",
            "odbc",
            "sql",
            "connection string",
            "msolap",
            "ace.oledb",
            "provider",
            "driver",
            "recordset",
        ),
    ),
    Route(
        layer="Deliverable publishing",
        skill="excel-deliverable-publisher",
        validation_needed=(
            "Create and verify a copied deliverable; never clean the only source workbook.",
            "Run post-clean structural verification and use Excel runtime validation if formulas, refresh, or macros affect values.",
        ),
        scripts=(
            "tools/build_external_dependency_report.py",
            "tools/build_pure_deliverable_cleanup_plan.py",
            "tools/build_pure_deliverable_verification_report.py",
            "tools/inspect_excel_bi_workbook.py",
        ),
        boundaries=(
            "A cleanup plan is not itself a cleaned workbook.",
            "Value-freezing should happen only after required calculation or refresh completes.",
        ),
        keywords=(
            "deliverable",
            "publish",
            "client-ready",
            "handoff",
            "final xlsx",
            "pure xlsx",
            "pure-value",
            "values only",
            "freeze formulas",
            "remove formulas",
            "remove links",
            "remove connections",
            "remove power query",
            "remove data model",
            "delete config sheet",
            "clean workbook",
        ),
    ),
    Route(
        layer="Workbook QA",
        skill="excel-workbook-qa-auditor",
        validation_needed=(
            "Static QA reports identify workbook risks; live Excel is needed for runtime calculation, refresh, VBA, and Data Model behavior.",
            "Findings should be prioritized and tied to workbook surfaces or follow-up commands.",
        ),
        scripts=(
            "tools/build_workbook_triage_report.py",
            "tools/build_formula_quality_report.py",
            "tools/build_workbook_controls_report.py",
            "tools/build_external_dependency_report.py",
        ),
        boundaries=(
            "Do not modify the workbook in an audit-only task.",
            "A static QA pass does not prove numeric correctness.",
        ),
        keywords=(
            "qa",
            "audit",
            "review",
            "risk",
            "readiness",
            "quality",
            "findings",
            "pre-delivery",
            "formula quality",
            "workbook controls",
            "broken formulas",
            "external dependency",
            "hidden sheets",
        ),
    ),
    Route(
        layer="Office environment",
        skill="office-environment-diagnostics",
        validation_needed=(
            "Separate the execution environment from the target environment before deciding which capability evidence applies.",
            "Use structural evidence cross-platform, Windows capability probes for local runtime readiness, and workbook-specific tests for final behavior.",
        ),
        scripts=(
            "tools/probe_excel_capabilities.ps1",
            "tools/build_excel_compatibility_report.py",
            "tools/probe_excel_bi_providers.ps1",
            "tools/build_provider_environment_report.py",
            "tools/create_provider_environment_fixture.py",
            "tools/test_excel_adomd_query.ps1",
        ),
        boundaries=(
            "Environment readiness is separate from workbook correctness.",
            "A captured probe describes its source machine; it does not automatically describe the workbook recipient's target environment.",
            "Synthetic provider fixtures prove report logic only.",
        ),
        keywords=(
            "compatibility",
            "compatible",
            "platform",
            "supported environment",
            "support matrix",
            "can this run",
            "windows",
            "linux",
            "macos",
            "mac os",
            "excel online",
            "excel web",
            "microsoft 365",
            "office ltsc",
            "offline",
            "excel com",
            "environment",
            "diagnostic",
            "diagnose",
            "provider",
            "ace",
            "msolap",
            "adomd",
            "adodb",
            "bitness",
            "office version",
            "excel installed",
            "trust access",
            "vba project access",
            "machine",
            "provider drift",
        ),
    ),
    Route(
        layer="Excel report building",
        skill="excel-report-builder",
        validation_needed=(
            "Use static workbook generation for layout where possible and Excel runtime validation when charts, pivots, controls, refresh, or macros matter.",
            "Validate final report surfaces before publishing.",
        ),
        scripts=(
            "tools/inspect_excel_bi_workbook.py",
            "tools/build_formula_quality_report.py",
            "tools/build_workbook_controls_report.py",
            "tools/build_workbook_triage_report.py",
        ),
        boundaries=(
            "Report layout work should not silently change model/query logic.",
            "Use publish cleanup only after the report workbook is validated.",
        ),
        keywords=(
            "report",
            "dashboard",
            "layout",
            "format",
            "chart",
            "table",
            "pivot",
            "client-facing",
            "presentation sheet",
            "analysis sheet",
            "build report",
            "create workbook",
            "polished",
        ),
    ),
    Route(
        layer="Power BI semantic model",
        skill="power-bi-semantic-model",
        validation_needed=(
            "Confirm whether the target is Power BI or Excel Power Pivot before writing DAX or model guidance.",
            "Use official Microsoft documentation for current semantic model, TMDL, XMLA, or Fabric behavior.",
        ),
        scripts=(
            ".agents/skills/power-pivot-dax-modeling/scripts/lint_dax_compat.py",
            ".agents/skills/power-pivot-dax-modeling/scripts/analyze_dax_dependencies.py",
            "tools/build_excel_bi_model_report.py",
            "tools/search_official_docs.py",
        ),
        boundaries=(
            "This package does not directly rewrite PBIX binaries.",
            "Excel Power Pivot and Power BI semantic models are not interchangeable runtime hosts.",
        ),
        keywords=(
            "power bi",
            "pbix",
            "semantic model",
            "dataset",
            "fabric",
            "tmdl",
            "xmla",
            "tabular",
            "calculation group",
            "deployment pipeline",
            "measure portability",
        ),
    ),
    Route(
        layer="Testing fixtures",
        skill="excel-testing-fixtures",
        validation_needed=(
            "Use sanitized fixtures to test tooling behavior without customer data.",
            "State which runtime behavior a fixture does and does not prove.",
        ),
        scripts=(
            "tools/build_sanitized_fixture_bundle.py",
            "tools/create_cube_formula_fixture.py",
            "tools/create_workbook_surface_fixture.py",
            "tools/build_cross_agent_forward_test_pack.py",
        ),
        boundaries=(
            "Fixtures are regression evidence for designed cases, not proof for every customer workbook.",
            "Generated sample responses are scorer fixtures, not external-agent proof.",
        ),
        keywords=(
            "fixture",
            "fixtures",
            "regression",
            "smoke test",
            "sample workbook",
            "sanitized",
            "test case",
            "forward-test",
            "safe example",
            "without customer data",
        ),
    ),
)


MIXED_VALIDATION = (
    "Start with workbook structure and dependency inspection, then route each edited layer to its specialized validation.",
    "Use Windows desktop Excel for final refresh, VBA, Data Model, and report-layer runtime behavior.",
)


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.casefold()).strip()


def keyword_score(text: str, keyword: str) -> int:
    normalized_keyword = keyword.casefold()
    if len(normalized_keyword) <= 3:
        return 1 if normalized_keyword in text else 0
    count = text.count(normalized_keyword)
    if count:
        return count * 3
    compact = normalized_keyword.replace(" ", "")
    return 2 if compact and compact in text.replace(" ", "") else 0


def score_routes(task: str) -> list[dict[str, Any]]:
    text = normalize(task)
    scored: list[dict[str, Any]] = []
    for route in ROUTES:
        matched: list[str] = []
        score = 0
        for keyword in route.keywords:
            value = keyword_score(text, keyword)
            if value:
                matched.append(keyword)
                score += value
        scored.append(
            {
                "layer": route.layer,
                "skill": route.skill,
                "score": score,
                "matchedKeywords": matched,
                "validationNeeded": list(route.validation_needed),
                "recommendedScripts": list(route.scripts),
                "boundaries": list(route.boundaries),
            }
        )
    return sorted(scored, key=lambda item: (-int(item["score"]), str(item["skill"])))


def build_report(task: str) -> dict[str, Any]:
    scored = score_routes(task)
    positive = [item for item in scored if int(item["score"]) > 0]
    top = positive[0] if positive else None
    second = positive[1] if len(positive) > 1 else None
    strong_positive = [item for item in positive if int(item["score"]) >= 5]
    is_mixed = len(strong_positive) >= 2 and (
        second is not None and int(second["score"]) >= max(5, int(top["score"]) * 0.75)
    )
    if len(strong_positive) >= 3:
        is_mixed = True

    if top is None:
        status = WARN
        layer = "Unknown"
        skill = "excel-bi-router"
        why = "No strong Excel BI layer signal was found. Inspect workbook shape before choosing a specialized skill."
        validation_needed = ["Inspect the workbook and task details before editing."]
        scripts: list[str] = ["tools/inspect_excel_bi_workbook.py"]
        boundaries = ["This router cannot infer a layer without task evidence."]
    elif is_mixed:
        status = PASS
        layer = "Mixed"
        skill = "excel-bi-router"
        route_names = ", ".join(item["skill"] for item in positive[:5])
        why = f"Multiple Excel BI layers are present or likely: {route_names}."
        validation_needed = list(MIXED_VALIDATION)
        scripts = ["tools/inspect_excel_bi_workbook.py", "tools/build_external_dependency_report.py"]
        for item in positive[:5]:
            for script in item["recommendedScripts"][:2]:
                if script not in scripts:
                    scripts.append(script)
        boundaries = [
            "Do not edit all layers at once; inspect dependencies and validate each changed layer.",
            "Cross-platform structural checks do not prove Excel runtime behavior.",
        ]
    else:
        status = PASS
        layer = str(top["layer"])
        skill = str(top["skill"])
        why = f"Top matched layer is {layer} with keywords: {', '.join(top['matchedKeywords'][:6])}."
        validation_needed = list(top["validationNeeded"])
        scripts = list(top["recommendedScripts"])
        boundaries = list(top["boundaries"])

    return {
        "status": status,
        "task": task,
        "layer": layer,
        "skill": skill,
        "why": why,
        "validationNeeded": validation_needed,
        "recommendedScripts": scripts,
        "boundaries": boundaries,
        "scoredRoutes": scored,
        "outputTemplate": {
            "Layer": layer,
            "Skill": skill,
            "Why": why,
            "Validation needed": " | ".join(validation_needed),
        },
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "# Excel BI Task Route",
        "",
        f"- Status: `{report['status']}`",
        f"- Layer: `{report['layer']}`",
        f"- Skill: `{report['skill']}`",
        f"- Why: {report['why']}",
        "",
        "## Validation Needed",
        "",
    ]
    for item in report["validationNeeded"]:
        lines.append(f"- {item}")
    lines.extend(["", "## Recommended Scripts", ""])
    for item in report["recommendedScripts"]:
        lines.append(f"- `{item}`")
    lines.extend(["", "## Boundaries", ""])
    for item in report["boundaries"]:
        lines.append(f"- {item}")
    lines.extend(["", "## Scores", ""])
    for item in report["scoredRoutes"]:
        if item["score"]:
            matches = ", ".join(item["matchedKeywords"][:8])
            lines.append(f"- `{item['skill']}` score `{item['score']}` from {matches}")
    if not any(item["score"] for item in report["scoredRoutes"]):
        lines.append("- No layer-specific keyword matches.")
    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument("--text", help="Task description to route")
    input_group.add_argument("--input", type=Path, help="Text file containing the task description")
    parser.add_argument("--out-json", type=Path, help="Optional JSON report path")
    parser.add_argument("--out-md", type=Path, help="Optional Markdown report path")
    parser.add_argument("--expect-skill", help="Fail if the selected skill differs")
    parser.add_argument("--expect-layer", help="Fail if the selected layer differs")
    return parser.parse_args()


def write_outputs(report: dict[str, Any], out_json: Path | None, out_md: Path | None) -> None:
    if out_json:
        path = out_json.expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if out_md:
        path = out_md.expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(markdown_report(report), encoding="utf-8")


def main() -> int:
    args = parse_args()
    task = args.text if args.text is not None else args.input.expanduser().resolve().read_text(encoding="utf-8-sig")
    report = build_report(task)
    write_outputs(report, args.out_json, args.out_md)
    print(f"Excel BI route: {report['layer']} -> {report['skill']} ({report['status']})")
    if args.expect_skill and report["skill"] != args.expect_skill:
        print(f"Expected skill {args.expect_skill}, got {report['skill']}", flush=True)
        return 1
    if args.expect_layer and report["layer"] != args.expect_layer:
        print(f"Expected layer {args.expect_layer}, got {report['layer']}", flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
