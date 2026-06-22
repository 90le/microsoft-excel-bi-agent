#!/usr/bin/env python3
"""Score cross-agent forward-test responses against the generated prompt pack."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path


PASS = "pass"
FAIL = "fail"


SKILL_RESPONSE_CHECKS: dict[str, list[dict[str, object]]] = {
    "excel-bi-router": [
        {
            "id": "specialized-skill-routing",
            "all": [
                "excel-vba-workbook-engineering",
                "power-query-m-engineering",
                "power-pivot-dax-modeling",
                "mdx-cubevalue-extraction",
                "excel-ado-sql-data-access",
            ],
        },
        {"id": "runtime-boundary", "any": ["Excel COM", "OpenXML", "static"]},
        {"id": "customer-data-boundary", "any": ["customer", "temp", "outside the plugin"]},
    ],
    "excel-vba-workbook-engineering": [
        {
            "id": "vba-tooling",
            "all": ["export_vba.ps1", "import_vba.ps1", "lint_vba_source.py"],
        },
        {"id": "button-binding", "any": ["build_vba_button_binding_report.py", "OnAction"]},
        {"id": "runtime-boundary", "any": ["Excel COM", "compile", "macro"]},
    ],
    "power-query-m-engineering": [
        {
            "id": "pq-tooling",
            "all": ["manage_power_queries_excel_com.ps1", "lint_power_query_m.py"],
        },
        {"id": "refresh-boundary", "any": ["refresh_power_queries_excel_com.ps1", "Excel COM", "refresh"]},
        {"id": "m-safety", "all": ["row", "join"]},
    ],
    "power-pivot-dax-modeling": [
        {
            "id": "dax-tooling",
            "all": ["lint_dax_compat.py", "analyze_dax_dependencies.py"],
        },
        {"id": "excel-dax-compatibility", "all": ["REMOVEFILTERS", "DIVIDE"]},
        {"id": "excel-filter-patterns", "all": ["ALL", "FILTER"]},
    ],
    "mdx-cubevalue-extraction": [
        {
            "id": "cube-tooling",
            "all": ["build_cube_dependency_report.py", "create_cube_formula_fixture.py"],
        },
        {"id": "cube-formula-surface", "all": ["CUBEVALUE", "measure"]},
        {"id": "live-calculation-boundary", "any": ["live", "calculation", "refresh", "Excel COM"]},
    ],
    "excel-ado-sql-data-access": [
        {
            "id": "provider-tooling",
            "all": ["probe_excel_bi_providers.ps1", "test_excel_ado_sql_access.ps1"],
        },
        {"id": "adomd-boundary", "all": ["ADOMD", "endpoint"]},
        {"id": "ace-workbook-sql", "all": ["ACE", "SQL"]},
    ],
    "excel-deliverable-publisher": [
        {
            "id": "publish-tooling",
            "all": ["build_pure_deliverable_cleanup_plan.py", "build_pure_deliverable_verification_report.py"],
        },
        {"id": "non-destructive-source", "any": ["copy", "source workbook", "not overwritten"]},
        {"id": "runtime-before-freeze", "any": ["refresh", "calculation", "value-freezing"]},
    ],
    "excel-workbook-qa-auditor": [
        {
            "id": "qa-tooling",
            "all": ["build_workbook_triage_report.py", "build_formula_quality_report.py"],
        },
        {"id": "risk-priority", "any": ["high", "medium", "low", "prioritized"]},
        {"id": "static-boundary", "any": ["numeric correctness", "static", "runtime"]},
    ],
    "office-environment-diagnostics": [
        {
            "id": "environment-tooling",
            "all": ["probe_excel_bi_providers.ps1", "build_provider_environment_report.py"],
        },
        {"id": "environment-vs-workbook", "all": ["environment", "workbook"]},
        {"id": "platform-boundary", "any": ["Linux", "macOS", "structural-only", "desktop Excel"]},
    ],
    "excel-report-builder": [
        {
            "id": "report-surfaces",
            "all": ["input", "calculation", "output"],
        },
        {"id": "report-validation", "all": ["build_formula_quality_report.py", "build_workbook_controls_report.py"]},
        {"id": "specialist-routing", "any": ["Power Query", "Data Model", "VBA", "publish"]},
    ],
    "power-bi-semantic-model": [
        {
            "id": "host-distinction",
            "all": ["Excel Power Pivot", "Power BI"],
        },
        {"id": "semantic-model-boundary", "any": ["PBIX", "TMDL", "XMLA", "semantic model"]},
        {"id": "official-docs", "any": ["Microsoft", "official", "documentation"]},
    ],
    "excel-testing-fixtures": [
        {
            "id": "fixture-tooling",
            "any": ["build_sanitized_fixture_bundle.py", "create_cube_formula_fixture.py", "create_workbook_surface_fixture.py"],
        },
        {"id": "fixture-boundary", "all": ["designed cases", "customer"]},
        {"id": "package-hygiene", "any": ["outside the plugin", "out of the plugin", "temp"]},
    ],
}


PLACEHOLDER_RE = re.compile(
    "|".join([r"\b" + "TO" + "DO" + r"\b", r"\[" + "TO" + "DO" + r"\]", r"FIX" + "ME"]),
    re.IGNORECASE,
)


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


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


def read_json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def normalize(text: str) -> str:
    return text.lower()


def contains_all(text: str, terms: list[str]) -> bool:
    lowered = normalize(text)
    return all(term.lower() in lowered for term in terms)


def contains_any(text: str, terms: list[str]) -> bool:
    lowered = normalize(text)
    return any(term.lower() in lowered for term in terms)


def response_path(responses_dir: Path, prompt_entry: dict[str, object]) -> Path:
    agent = str(prompt_entry.get("agent", "")).strip()
    skill = str(prompt_entry.get("skill", "")).strip()
    return responses_dir / agent / f"{skill}.md"


def score_response(path: Path, skill: str) -> dict[str, object]:
    failures: list[str] = []
    passed_checks: list[str] = []
    failed_checks: list[str] = []

    if not path.is_file():
        return {
            "status": FAIL,
            "path": str(path),
            "skill": skill,
            "passedChecks": [],
            "failedChecks": ["missing-response-file"],
            "failures": [f"missing response file: {path}"],
        }

    text = path.read_text(encoding="utf-8")
    if len(text.strip()) < 120:
        failures.append("response is too short to be reviewable")
        failed_checks.append("minimum-length")
    else:
        passed_checks.append("minimum-length")
    if PLACEHOLDER_RE.search(text):
        failures.append("response contains placeholder marker")
        failed_checks.append("placeholder-free")
    else:
        passed_checks.append("placeholder-free")

    for check in SKILL_RESPONSE_CHECKS.get(skill, []):
        check_id = str(check["id"])
        all_terms = [str(term) for term in check.get("all", [])]
        any_terms = [str(term) for term in check.get("any", [])]
        ok = True
        if all_terms and not contains_all(text, all_terms):
            ok = False
            failures.append(f"{check_id}: missing all terms {all_terms}")
        if any_terms and not contains_any(text, any_terms):
            ok = False
            failures.append(f"{check_id}: missing any of {any_terms}")
        if ok:
            passed_checks.append(check_id)
        else:
            failed_checks.append(check_id)

    return {
        "status": PASS if not failures else FAIL,
        "path": str(path),
        "skill": skill,
        "passedChecks": passed_checks,
        "failedChecks": failed_checks,
        "failures": failures,
    }


def sample_response(agent: str, skill: str, failing: bool = False) -> str:
    checks = SKILL_RESPONSE_CHECKS[skill]
    terms: list[str] = []
    for check in checks:
        terms.extend(str(term) for term in check.get("all", []))
        any_terms = [str(term) for term in check.get("any", [])]
        if any_terms:
            terms.append(any_terms[0])
    if failing:
        return "\n".join(
            [
                f"# Forward-test response for {agent} / {skill}",
                "",
                "This intentionally incomplete response is used to verify that the scorer fails missing evidence.",
                "It names the task but omits the expected package tooling and runtime boundary proof.",
                "",
            ]
        )
    return "\n".join(
        [
            f"# Forward-test response for {agent} / {skill}",
            "",
            "Concrete action path:",
            "Read the referenced SKILL.md, use package fixtures or a task-local temp folder, and keep customer workbooks outside the plugin source tree.",
            "",
            "Evidence terms covered:",
            ", ".join(sorted(set(terms))),
            "",
            "Runtime boundary:",
            "Static OpenXML/source checks are separated from live Excel COM, Power Query refresh, VBA execution, and ADOMD endpoint validation.",
            "",
        ]
    )


def write_sample_responses(manifest: dict[str, object], responses_dir: Path, failing: bool = False) -> None:
    responses_dir.mkdir(parents=True, exist_ok=True)
    prompts = [item for item in manifest.get("prompts", []) if isinstance(item, dict)]
    for index, prompt in enumerate(prompts):
        skill = str(prompt.get("skill"))
        agent = str(prompt.get("agent"))
        path = response_path(responses_dir, prompt)
        path.parent.mkdir(parents=True, exist_ok=True)
        make_failing = failing and index == 0
        path.write_text(sample_response(agent, skill, failing=make_failing), encoding="utf-8")


def score_manifest(manifest: dict[str, object], responses_dir: Path) -> dict[str, object]:
    prompts = [item for item in manifest.get("prompts", []) if isinstance(item, dict)]
    results: list[dict[str, object]] = []
    failures: list[str] = []
    expected_skills = set(SKILL_RESPONSE_CHECKS)
    manifest_skills = {str(item.get("skill")) for item in prompts}
    missing_skill_checks = sorted(manifest_skills - expected_skills)
    for skill in missing_skill_checks:
        failures.append(f"no response checks configured for skill: {skill}")

    for prompt in prompts:
        skill = str(prompt.get("skill"))
        result = score_response(response_path(responses_dir, prompt), skill)
        result["agent"] = str(prompt.get("agent"))
        results.append(result)
        if result["status"] != PASS:
            failures.append(f"{result['agent']}/{skill}: {', '.join(result['failures'])}")

    passed_count = sum(1 for item in results if item["status"] == PASS)
    failed_count = len(results) - passed_count
    return {
        "generatedAt": now_iso(),
        "status": PASS if not failures else FAIL,
        "packVersion": manifest.get("packVersion"),
        "manifestStatus": manifest.get("status"),
        "responsesDir": str(responses_dir),
        "expectedResponseCount": len(prompts),
        "scoredResponseCount": len(results),
        "passedCount": passed_count,
        "failedCount": failed_count,
        "results": results,
        "failures": failures,
        "boundaries": [
            "This scorer evaluates response evidence text; it does not prove that an external agent actually ran unless the responses came from fresh sessions.",
            "Use generated sample responses only to validate scorer mechanics, not as real forward-test evidence.",
        ],
    }


def markdown_report(report: dict[str, object]) -> str:
    lines = [
        "# Cross-Agent Forward-Test Result Score",
        "",
        f"- Status: `{report.get('status')}`",
        f"- Expected responses: `{report.get('expectedResponseCount')}`",
        f"- Passed: `{report.get('passedCount')}`",
        f"- Failed: `{report.get('failedCount')}`",
        "",
        "| Agent | Skill | Status | Failed Checks |",
        "|---|---|---|---|",
    ]
    for item in report.get("results", []):
        if not isinstance(item, dict):
            continue
        failed = ", ".join(str(value) for value in item.get("failedChecks", []))
        lines.append(f"| {item.get('agent', '')} | {item.get('skill', '')} | {item.get('status', '')} | {failed} |")
    if report.get("failures"):
        lines.extend(["", "## Failures", ""])
        for failure in report["failures"]:
            lines.append(f"- {failure}")
    lines.extend(
        [
            "",
            "## Boundary",
            "",
            "Generated sample responses validate scorer mechanics only. Real forward-test evidence requires fresh agent-session outputs saved into the response folder.",
            "",
        ]
    )
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest-json", required=True, type=Path, help="Manifest JSON from build_cross_agent_forward_test_pack.py")
    parser.add_argument("--responses-dir", required=True, type=Path, help="Directory containing responses/<agent>/<skill>.md files")
    parser.add_argument("--out-json", type=Path, help="Optional score JSON path")
    parser.add_argument("--out-md", type=Path, help="Optional score Markdown path")
    parser.add_argument("--clean-responses", action="store_true", help="Remove responses dir before writing sample responses")
    parser.add_argument("--write-passing-fixture", action="store_true", help="Write synthetic passing responses before scoring")
    parser.add_argument("--write-failing-fixture", action="store_true", help="Write one intentionally failing synthetic response before scoring")
    parser.add_argument("--require-pass", action="store_true", help="Return non-zero when score status is not pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.write_passing_fixture and args.write_failing_fixture:
        print("choose only one of --write-passing-fixture or --write-failing-fixture", file=sys.stderr)
        return 2
    manifest_path = args.manifest_json.expanduser().resolve()
    responses_dir = args.responses_dir.expanduser().resolve()
    manifest = read_json(manifest_path)

    if args.clean_responses:
        safe_clean_dir(responses_dir)
    if args.write_passing_fixture or args.write_failing_fixture:
        write_sample_responses(manifest, responses_dir, failing=args.write_failing_fixture)

    report = score_manifest(manifest, responses_dir)
    report["manifestJson"] = str(manifest_path)

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_report(report), encoding="utf-8")

    print(
        json.dumps(
            {
                "status": report["status"],
                "expectedResponseCount": report["expectedResponseCount"],
                "passedCount": report["passedCount"],
                "failedCount": report["failedCount"],
            },
            ensure_ascii=False,
        )
    )
    if args.require_pass and report["status"] != PASS:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
