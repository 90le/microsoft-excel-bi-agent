# Microsoft Excel BI Agent Pack

## Purpose

This repository provides cross-agent skills for Microsoft Excel BI engineering: Excel workbooks, VBA, Power Query M, Power Pivot DAX, MDX/CUBE formulas, ADO/SQL data access, deliverable publishing, workbook QA, Office environment diagnostics, Excel report building, Power BI semantic model review, and sanitized testing fixtures.

Use these skills with Claude, Codex, OpenCode, or other agents that understand `AGENTS.md` and `SKILL.md`.

## Skill Location

Canonical skills live in:

```text
.agents/skills/
```

Codex plugin packaging also needs a root-level `skills/` mirror because plugin validation expects that conventional path. Treat `skills/` as generated output, not as the source of truth.

Do not maintain divergent copies in `skills/`, `.claude/skills`, `.opencode/skills`, or `~/.codex/skills`. If a tool needs a different install path, copy from `.agents/skills` using `tools/sync-skills.py`.

One-command prompt sync for Codex plugin mirror, Codex user skills, Claude project skills, and OpenCode project skills:

```bash
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Install or refresh the Codex plugin:

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
```

Use `--update-cachebuster` only when behavior changed and a new installed Codex plugin version is required. Do not hand-edit marketplace JSON.

For the constrained install contract, read `docs/install-and-sync.md`.

## Routing

Start with `.agents/skills/excel-bi-router/SKILL.md` unless the user explicitly names a skill.

Use:

- `excel-vba-workbook-engineering` for workbook files, VBA, macro buttons, Excel COM, and cross-platform workbook inspection.
- `power-query-m-engineering` for Power Query M queries, refresh errors, file ingestion, joins, grouping, expansion, and type handling.
- `power-pivot-dax-modeling` for Data Model relationships, DAX measures, context, and Power Pivot calculations.
- `mdx-cubevalue-extraction` for `CUBEVALUE`, `CUBEMEMBER`, MDX member references, and `ThisWorkbookDataModel` report formulas.
- `excel-ado-sql-data-access` for VBA ADO/OLEDB/ADOMD, SQL against Excel tables, and query templates.
- `excel-deliverable-publisher` for copied clean deliverables, values-only outputs, link/query/model cleanup, and post-clean verification.
- `excel-workbook-qa-auditor` for pre-delivery workbook QA, visual QA, and prioritized risk findings.
- `office-environment-diagnostics` for Excel COM, provider, bitness, Trust Center, and local automation readiness.
- `excel-report-builder` for polished Excel report/dashboard workbook surfaces.
- `power-bi-semantic-model` for Power BI semantic model concepts and Excel Power Pivot portability boundaries.
- `excel-testing-fixtures` for sanitized workbook fixtures, regression cases, and forward-test prompts.

For the current operational snapshot, use `docs/current-status.md`. For end-to-end examples, use `docs/task-recipes.md`. For goal tracking, completion, and validation evidence, use `docs/goal-tracking.md`, `docs/completion-evidence.md`, and `docs/validation.md`.
For installation and prompt/agent sync, use `docs/install-and-sync.md`.
For the real/sanitized regression library, use `docs/real-case-regression.md`.
For rendered report-surface evidence, use `tools/export_visual_qa_render_evidence.ps1` on Windows desktop Excel only; it exports task-local PDFs and JSON/Markdown evidence for visible `Report*` sheets and does not belong in committed package artifacts.

## Fast Profiles

Use `tools/run_task_profile.py` when a request fits a common workflow and the agent needs a command plan before choosing individual scripts:

- `audit`: workbook QA audit.
- `publish`: pure deliverable planning and verification.
- `pq-refresh`: Power Query refresh plus status report.
- `dax-review`: exported model/DAX review.
- `cube-trace`: CUBE/MDX report-layer tracing.
- `env-diagnostics`: Office/provider diagnostics.
- `report-build`: report-surface validation.
- `fixture`: sanitized fixture bundle.
- `case-regression`: real/sanitized regression case library validation.
- `release-structural` / `release-full`: package validation.

## Cross-Platform Rules

- Windows desktop Excel + PowerShell can perform full Excel COM validation, VBA import/export, macro execution, query refresh, and Solver checks.
- Windows desktop Excel + PowerShell can also export rendered Visual QA PDF evidence for sanitized report sheets.
- Windows Git Bash must call PowerShell/Excel COM through wrapper scripts.
- Linux/macOS can inspect OpenXML workbook structure and generate code, but cannot truthfully validate Excel VBA execution, Power Query refresh, Power Pivot Data Model contents, Solver, rendered Excel PDFs, or button click behavior without desktop Excel.
- Always state which environment was used for validation.

## Editing Rules

- Work on copies by default. Do not overwrite a source workbook unless explicitly requested.
- Keep code source files outside binary workbooks when possible.
- Prefer deterministic scripts for repeated inspection/export/import.
- Do not put customer-specific business rules in reusable skills.
- Do not paste full Microsoft function documentation into `SKILL.md`; use references and official links.
- Do not add broad new skills or long duplicate docs unless a repeated workflow cannot be handled by an existing skill plus script.

## Validation

Preferred full package gate:

```bash
python tools/run_release_gate.py --project-root .
```

Portable wrapper for Git Bash, Linux, and macOS:

```bash
tools/run_release_gate.sh
```

Structural-only gate for Linux/macOS or environments without Windows Excel COM:

```bash
tools/run_release_gate.sh --profile structural
```

For official-documentation index changes, also use:

```bash
python tools/validate_official_docs_index.py --project-root .
```

For every skill change:

1. Validate `SKILL.md` frontmatter: `name` and `description` are present and the name matches the folder.
2. Scan for unfinished placeholder markers.
3. Check scripts for syntax on the relevant platform.
4. Run a smoke test when a script is changed.
5. Record platform limitations in the final response.

## Current Status

This package now includes a Codex plugin manifest at `.codex-plugin/plugin.json`. The plugin `skills/` mirror and other agent mirrors should be regenerated from `.agents/skills` before validation or packaging. Treat `docs/progress.md` as a historical ledger; read `docs/current-status.md` and `docs/install-and-sync.md` first for current state and install behavior.
