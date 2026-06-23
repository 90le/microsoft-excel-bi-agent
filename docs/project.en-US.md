# Microsoft Excel BI Agent Project Overview

## Objective

Microsoft Excel BI Agent is an open-source cross-agent Excel BI skill pack for AI agents that work on Excel workbooks, VBA, Power Query M, Power Pivot DAX, MDX/CUBE formulas, ADO/SQL, workbook QA, clean deliverables, report building, Office diagnostics, and sanitized fixtures.

The project turns recurring Excel BI risks into reusable agent workflows: read the workbook surface, choose the right Excel BI layer, make narrow edits, validate the result, and document what was verified or skipped.

Maintainer: **Qiu Binbin (丘彬彬)**. WeChat: **binstudy**. Blog: **https://90le.cn**.

## Who It Is For

- Teams using Codex, Claude, OpenCode, or similar agents to work on Excel BI files.
- Analysts and automation engineers who need safer workbook editing workflows.
- Delivery teams that must produce clean `.xlsx` or `.xlsm` client files without leaking process sheets, links, credentials, or local paths.
- Maintainers who need repeatable checks for VBA, Power Query, DAX, MDX/CUBE, ADO/SQL, and Office runtime boundaries.

## What It Contains

```text
microsoft-excel-bi-agent/
  README.md                    # English repository entry
  README.zh-CN.md              # Chinese repository entry
  LICENSE
  .codex-plugin/plugin.json
  marketplace.json             # Codex marketplace metadata
  .agents/skills/              # source of truth
  skills/                      # generated Codex plugin mirror
  .claude/skills/              # generated Claude mirror
  .opencode/skills/            # generated OpenCode mirror
  docs/
  fixtures/
  prompts/
  tools/
```

## Core Capabilities

| Area | Capability |
| --- | --- |
| Excel/VBA workbook engineering | Export/import modules, bind buttons, debug compile/runtime errors, preserve workbook structure |
| Power Query M | Read, edit, refresh, diagnose errors, track dependencies, and wait for refresh completion |
| Power Pivot DAX | Review measures, context behavior, relationships, and Excel Data Model boundaries |
| MDX / CUBE formulas | Audit `CUBEVALUE`, `CUBEMEMBER`, measures, members, and helper-cell references |
| ADO / SQL | Query workbook tables, external files, and Data Model sources through ADO/OLEDB/ADOMD patterns |
| Client deliverables | Freeze formulas, remove links/queries/Data Model dependencies, delete config/process sheets |
| Workbook QA | Audit formulas, hidden sheets, controls, external dependencies, and delivery risk |
| Cross-agent distribution | Sync one canonical skill source into Codex, Claude, and OpenCode style folders |

## Installation Paths

Codex marketplace:

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

Local one-command install:

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

Manual install:

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

## Public Repository Boundary

The public repository contains source skills, scripts, fixtures, install prompts, and recipient-facing documentation. Maintainer-only release ledgers and machine-specific runtime reports are intentionally not included.

Windows desktop Excel is required for Excel COM, VBA execution, Power Query refresh, and Power Pivot/Data Model runtime validation. macOS and Linux support structural checks, OpenXML inspection, documentation, and non-COM scripts.

## Validation

Public validation:

```bash
python tools/validate-skills.py .
python tools/validate_project_docs.py --project-root .
python tools/validate_github_community_health.py --project-root .
python tools/validate_task_recipes.py --project-root .
python tools/validate_official_docs_index.py --project-root .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
python tools/build_goal_coverage_report.py --project-root . --require-pass
node tools/install.mjs --check
```

Full runtime validation, Windows desktop Excel only:

```powershell
python tools\run_release_gate.py --project-root .
```

## Related Pages

- [Chinese project overview](project.zh-CN.md)
- [Maintenance goals and risk backlog](maintenance-goals.en-US.md)
- [Public growth goals](growth-goals.en-US.md)
- [Repository governance goals](repository-governance-goals.en-US.md)
- [Marketing copy pack](marketing-copy.en-US.md)
- [Release notes](release-notes.en-US.md)
- [Contributing guide](https://github.com/90le/microsoft-excel-bi-agent/blob/main/CONTRIBUTING.md)
- [Security policy](https://github.com/90le/microsoft-excel-bi-agent/blob/main/SECURITY.md)
- [English site](intro.html)
- [Chinese site](intro.zh-CN.html)
- [Install and sync guide](install-and-sync.md)
- [Compatibility boundaries](compatibility.md)
