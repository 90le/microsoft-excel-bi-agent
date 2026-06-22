#!/usr/bin/env python3
"""Build a customer-data-free cross-agent forward-test prompt pack."""

from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path


PACK_VERSION = 1

AGENT_TARGETS = [
    {
        "id": "codex",
        "label": "Codex",
        "skillPathPattern": "skills/{skill}/SKILL.md",
        "invocation": "Use ${skill-name} from the installed plugin or the local skill path.",
    },
    {
        "id": "claude",
        "label": "Claude",
        "skillPathPattern": ".claude/skills/{skill}/SKILL.md",
        "invocation": "Use the mirrored skill folder as the task-specific operating guide.",
    },
    {
        "id": "opencode",
        "label": "OpenCode",
        "skillPathPattern": ".opencode/skills/{skill}/SKILL.md",
        "invocation": "Use the mirrored skill folder and keep outputs in a temp folder.",
    },
    {
        "id": "generic",
        "label": "Generic Agent",
        "skillPathPattern": ".agents/skills/{skill}/SKILL.md",
        "invocation": "Read the referenced SKILL.md first, then use only task-relevant references.",
    },
]

SKILL_TASKS = [
    {
        "skill": "excel-bi-router",
        "title": "Route an Excel BI workbook request",
        "request": (
            "Given a workbook request that mentions VBA macros, Power Query refresh, "
            "Power Pivot measures, CUBEVALUE formulas, and ADO access, identify which "
            "specialized skills and package tools should be used first. Produce a short "
            "ordered routing plan and state which claims require Windows Excel COM."
        ),
        "expectedEvidence": [
            "routes to all relevant specialized skills",
            "separates OpenXML/static checks from Excel COM runtime checks",
            "keeps customer files out of the plugin source tree",
        ],
    },
    {
        "skill": "excel-vba-workbook-engineering",
        "title": "Audit workbook automation surfaces",
        "request": (
            "Use the package's generic fixture/report tools to explain how you would "
            "audit formulas, workbook controls, VBA public entry points, and button "
            "bindings in an .xlsm workbook. Do not use a customer workbook."
        ),
        "expectedEvidence": [
            "mentions VBA export/import/lint and button binding reports",
            "mentions formula quality and workbook controls reports",
            "states that VBA compile/run evidence needs Excel on Windows",
        ],
    },
    {
        "skill": "power-query-m-engineering",
        "title": "Review exported Power Query M lineage",
        "request": (
            "Generate or describe the sanitized Power Query lineage fixture and review "
            "safe versus risky exported .m queries. Call out local paths, web/database/"
            "cloud-service sources, native SQL, credential-like literals, mixed-source "
            "lineage, and query cycles."
        ),
        "expectedEvidence": [
            "uses exported M source-lineage/source-risk workflow",
            "does not claim refresh success without Excel COM refresh evidence",
            "preserves row-order and join-cardinality concerns when discussing M edits",
        ],
    },
    {
        "skill": "power-pivot-dax-modeling",
        "title": "Check Excel Power Pivot DAX compatibility",
        "request": (
            "Review DAX formulas for Excel Power Pivot compatibility. Flag Power BI-only "
            "or version-sensitive functions, prefer ALL/FILTER/DIVIDE patterns for Excel, "
            "and describe how dependency analysis should be validated."
        ),
        "expectedEvidence": [
            "flags REMOVEFILTERS as not safe for Excel Power Pivot by default",
            "prefers DIVIDE for ratios and ALL/FILTER for filter control",
            "uses dependency analysis before measure rename/delete rewrites",
        ],
    },
    {
        "skill": "mdx-cubevalue-extraction",
        "title": "Trace CUBE formulas to model measures",
        "request": (
            "Use the generic CUBE formula fixture workflow to inspect CUBEVALUE and "
            "related formulas, map measure references, flag missing measures, and "
            "explain why structural formula parsing is not live cube calculation."
        ),
        "expectedEvidence": [
            "mentions CUBE formula dependency report",
            "flags missing measures and dynamic MDX/helper-cell patterns",
            "states that CUBEVALUE result correctness requires refresh/calculation evidence",
        ],
    },
    {
        "skill": "excel-ado-sql-data-access",
        "title": "Validate workbook SQL and ADOMD boundaries",
        "request": (
            "Explain how to validate ADO workbook SQL and ADOMD access using the package "
            "probe tools and fixtures. Separate ACE workbook SQL smoke evidence from "
            "ADOMD endpoint query evidence."
        ),
        "expectedEvidence": [
            "mentions provider probe and provider environment report",
            "mentions ACE/ADO workbook SQL fixture",
            "states that real ADOMD endpoint execution needs explicit connection string and MDX",
        ],
    },
    {
        "skill": "excel-deliverable-publisher",
        "title": "Publish a clean Excel deliverable",
        "request": (
            "Explain the non-destructive workflow for producing a client-ready workbook copy "
            "with formulas frozen to values, external links/connections removed, process sheets "
            "cleaned, and post-clean verification recorded. Do not modify a real workbook."
        ),
        "expectedEvidence": [
            "uses cleanup plan and post-clean verification reports",
            "states that source workbook must not be overwritten",
            "separates value-freezing from required refresh or calculation evidence",
        ],
    },
    {
        "skill": "excel-workbook-qa-auditor",
        "title": "Audit workbook delivery readiness",
        "request": (
            "Describe how to audit a workbook for formula, controls, hidden sheets, external "
            "dependency, Power Query, Data Model, CUBE, and VBA-button risks before delivery. "
            "Return prioritized findings and runtime-boundary language."
        ),
        "expectedEvidence": [
            "mentions workbook triage plus formula, controls, and external dependency reports",
            "prioritizes QA findings by delivery risk",
            "states that static QA does not prove numeric correctness",
        ],
    },
    {
        "skill": "office-environment-diagnostics",
        "title": "Diagnose Office provider readiness",
        "request": (
            "Explain how to diagnose whether the current machine can run Excel automation, "
            "Power Query refresh, ADO workbook SQL, MSOLAP, and ADOMD tasks. Include the "
            "provider probe and environment report path."
        ),
        "expectedEvidence": [
            "mentions provider probe and provider environment report",
            "separates Office/provider readiness from workbook correctness",
            "states Linux/macOS structural-only boundaries when desktop Excel is unavailable",
        ],
    },
    {
        "skill": "excel-report-builder",
        "title": "Build a polished Excel report workbook",
        "request": (
            "Describe how to create or revise a client-facing Excel report workbook with "
            "visible report sheets, stable inputs, formulas/tables/charts, and validation before "
            "final publishing."
        ),
        "expectedEvidence": [
            "separates input, calculation, output, and QA surfaces",
            "mentions formula quality, controls, and triage validation",
            "routes query/model/VBA changes to specialist skills before report publishing",
        ],
    },
    {
        "skill": "power-bi-semantic-model",
        "title": "Review Power BI semantic model portability",
        "request": (
            "Review a DAX/modeling request that might target either Excel Power Pivot or Power BI "
            "semantic models. Explain how to distinguish host assumptions, DAX portability, and "
            "validation evidence."
        ),
        "expectedEvidence": [
            "distinguishes Excel Power Pivot from Power BI semantic models",
            "mentions PBIX/TMDL/XMLA or semantic model boundaries",
            "uses official Microsoft documentation for current Power BI behavior",
        ],
    },
    {
        "skill": "excel-testing-fixtures",
        "title": "Create sanitized Excel BI regression fixtures",
        "request": (
            "Choose safe fixtures to test Excel BI parser/report workflows without customer files. "
            "Cover workbook surface, CUBE/model, Power Query lineage, provider drift, and "
            "cross-agent prompt fixtures."
        ),
        "expectedEvidence": [
            "mentions sanitized fixture bundle or specific fixture generators",
            "states fixtures prove designed cases only",
            "keeps customer workbooks and machine reports out of the plugin package",
        ],
    },
]


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


def read_skill_description(skill_dir: Path) -> str:
    skill_md = skill_dir / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    in_frontmatter = False
    for line in text.splitlines():
        stripped = line.strip()
        if stripped == "---":
            in_frontmatter = not in_frontmatter
            continue
        if in_frontmatter and stripped.startswith("description:"):
            return stripped.split(":", 1)[1].strip().strip('"')
    return ""


def prompt_text(agent: dict[str, str], task: dict[str, object], project_placeholder: str) -> str:
    skill = str(task["skill"])
    skill_path = f"{project_placeholder}/{agent['skillPathPattern'].format(skill=skill)}"
    expected = "\n".join(f"- {item}" for item in task["expectedEvidence"])
    return "\n".join(
        [
            f"# Forward Test: {agent['label']} / ${skill}",
            "",
            "## Setup",
            "",
            f"- Skill path: `{skill_path}`",
            f"- Invocation: {agent['invocation'].replace('${skill-name}', '$' + skill)}",
            "- Use only generic fixtures or a temp output folder.",
            "- Do not use customer workbooks, local screenshots, or machine-specific reports as source material.",
            "- If Excel COM, provider, or ADOMD endpoint evidence is unavailable, state the boundary instead of claiming success.",
            "",
            "## User Request",
            "",
            str(task["request"]),
            "",
            "## Expected Evidence",
            "",
            expected,
            "",
            "## Response Contract",
            "",
            "- Start with the concrete action path.",
            "- Name the package tools or references you would use.",
            "- Distinguish static OpenXML/source checks from live Excel runtime checks.",
            "- Keep outputs outside the plugin package unless explicitly asked to update the plugin.",
            "",
        ]
    )


def write_json(path: Path, data: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def build_readme(manifest: dict[str, object]) -> str:
    lines = [
        "# Cross-Agent Forward Test Pack",
        "",
        "This generated pack gives Codex, Claude, OpenCode, and generic agents the same customer-data-free tasks.",
        "Use it to forward-test skill behavior without leaking intended answers or customer workbooks.",
        "",
        "## Coverage",
        "",
    ]
    for skill in manifest["skills"]:
        lines.append(f"- `{skill['name']}`: {skill['title']}")
    lines.extend(
        [
            "",
            "## Agent Targets",
            "",
        ]
    )
    for agent in manifest["agentTargets"]:
        lines.append(f"- `{agent['id']}`: {agent['label']} using `{agent['skillPathPattern']}`")
    lines.extend(
        [
            "",
            "## Validation Boundary",
            "",
            "- This pack validates prompt coverage and transferable task framing.",
            "- It does not prove that an external agent actually ran the tasks.",
            "- Run fresh agent sessions with these prompts for human-reviewed forward-testing.",
            "",
        ]
    )
    return "\n".join(lines)


def validate_manifest(manifest: dict[str, object], out_dir: Path) -> list[str]:
    failures: list[str] = []
    if manifest.get("packVersion") != PACK_VERSION:
        failures.append(f"packVersion={manifest.get('packVersion')}")
    if manifest.get("agentTargetCount") != len(AGENT_TARGETS):
        failures.append(f"agentTargetCount={manifest.get('agentTargetCount')}")
    if manifest.get("skillCount") != len(SKILL_TASKS):
        failures.append(f"skillCount={manifest.get('skillCount')}")
    expected_prompt_count = len(AGENT_TARGETS) * len(SKILL_TASKS)
    if manifest.get("promptCount") != expected_prompt_count:
        failures.append(f"promptCount={manifest.get('promptCount')}")
    if manifest.get("status") != "pass":
        failures.append(f"status={manifest.get('status')}")

    prompt_paths = [out_dir / str(item["path"]) for item in manifest.get("prompts", []) if isinstance(item, dict)]
    if len(prompt_paths) != expected_prompt_count:
        failures.append(f"prompt path count={len(prompt_paths)}")
    placeholder_marker = "TO" + "DO"
    for prompt in prompt_paths:
        if not prompt.is_file():
            failures.append(f"missing prompt={prompt}")
            continue
        text = prompt.read_text(encoding="utf-8")
        if "Expected Evidence" not in text:
            failures.append(f"prompt missing Expected Evidence={prompt.name}")
        if "customer workbooks" not in text:
            failures.append(f"prompt missing customer-data boundary={prompt.name}")
        if placeholder_marker in text:
            failures.append(f"prompt contains placeholder marker={prompt.name}")
    return failures


def build_pack(
    project_root: Path,
    out_dir: Path,
    clean: bool = False,
    project_placeholder: str = "<PLUGIN_ROOT>",
) -> dict[str, object]:
    project_root = project_root.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    if clean:
        safe_clean_dir(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    prompts_dir = out_dir / "prompts"
    prompts_dir.mkdir(parents=True, exist_ok=True)

    skills: list[dict[str, object]] = []
    prompts: list[dict[str, object]] = []
    failures: list[str] = []

    for task in SKILL_TASKS:
        skill = str(task["skill"])
        skill_dir = project_root / ".agents" / "skills" / skill
        if not (skill_dir / "SKILL.md").is_file():
            failures.append(f"missing canonical skill: {skill}")
            description = ""
        else:
            description = read_skill_description(skill_dir)
        skills.append(
            {
                "name": skill,
                "title": task["title"],
                "description": description,
                "expectedEvidence": task["expectedEvidence"],
            }
        )
        for agent in AGENT_TARGETS:
            agent_dir = prompts_dir / str(agent["id"])
            agent_dir.mkdir(parents=True, exist_ok=True)
            prompt_path = agent_dir / f"{skill}.md"
            prompt_path.write_text(prompt_text(agent, task, project_placeholder), encoding="utf-8")
            prompts.append(
                {
                    "agent": agent["id"],
                    "skill": skill,
                    "path": prompt_path.relative_to(out_dir).as_posix(),
                    "title": task["title"],
                }
            )

    manifest: dict[str, object] = {
        "packVersion": PACK_VERSION,
        "generatedAt": now_iso(),
        "status": "pass" if not failures else "fail",
        "projectRootPlaceholder": project_placeholder,
        "agentTargetCount": len(AGENT_TARGETS),
        "skillCount": len(SKILL_TASKS),
        "promptCount": len(prompts),
        "agentTargets": AGENT_TARGETS,
        "skills": skills,
        "prompts": prompts,
        "failures": failures,
        "boundaries": [
            "Forward-test prompts are evaluation inputs, not proof that an external agent ran them.",
            "Prompts use generic fixtures and temp folders only; customer workbook paths are intentionally absent.",
            "Excel COM, Power Query refresh, VBA execution, and ADOMD endpoint behavior still need runtime evidence.",
        ],
    }
    validation_failures = validate_manifest(manifest, out_dir)
    if validation_failures:
        manifest["status"] = "fail"
        manifest["failures"] = failures + validation_failures
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", default=".", type=Path, help="Plugin project root")
    parser.add_argument("--out-dir", required=True, type=Path, help="Output folder for generated prompt pack")
    parser.add_argument("--clean", action="store_true", help="Remove the output folder before generating")
    parser.add_argument("--out-json", type=Path, help="Optional manifest JSON path")
    parser.add_argument("--out-md", type=Path, help="Optional README Markdown path")
    parser.add_argument("--project-placeholder", default="<PLUGIN_ROOT>", help="Placeholder used inside generated prompts")
    parser.add_argument("--require-pass", action="store_true", help="Return non-zero when validation fails")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = build_pack(args.project_root, args.out_dir, clean=args.clean, project_placeholder=args.project_placeholder)
    out_dir = args.out_dir.expanduser().resolve()
    out_json = args.out_json.expanduser().resolve() if args.out_json else out_dir / "forward-test-pack.json"
    out_md = args.out_md.expanduser().resolve() if args.out_md else out_dir / "README.md"
    manifest["manifestPath"] = str(out_json)
    manifest["readmePath"] = str(out_md)
    write_json(out_json, manifest)
    out_md.parent.mkdir(parents=True, exist_ok=True)
    out_md.write_text(build_readme(manifest), encoding="utf-8")
    print(json.dumps({"status": manifest["status"], "promptCount": manifest["promptCount"]}, ensure_ascii=False))
    if args.require_pass and manifest["status"] != "pass":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
