# Microsoft Excel BI Agent

![Microsoft Excel BI Agent hero](assets/readme-hero.png)

[![Release](https://img.shields.io/github/v/release/90le/microsoft-excel-bi-agent?include_prereleases&style=flat-square)](https://github.com/90le/microsoft-excel-bi-agent/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg?style=flat-square)](LICENSE)
[![Skills](https://img.shields.io/badge/skills-12_excel_bi_workflows-217346?style=flat-square)](.agents/skills)
[![Agents](https://img.shields.io/badge/agents-Codex%20%7C%20Claude%20%7C%20OpenCode-blue?style=flat-square)](docs/install-and-sync.md)

[English](README.md) | [中文](README.zh-CN.md) | [Website](https://90le.github.io/microsoft-excel-bi-agent/intro.html)

**Make AI agents reliable on real Microsoft Excel BI workbooks.**

Microsoft Excel BI Agent is an open-source, cross-agent skill pack for teams that use AI agents to inspect, modify, debug, validate, and publish Excel BI workbooks. It covers **Excel VBA**, **Power Query M**, **Power Pivot DAX**, **MDX/CUBE formulas**, **ADO/SQL**, workbook QA, clean deliverables, Office diagnostics, report building, semantic model review, and sanitized testing fixtures.

It is built for the messy Excel work that generic coding agents usually mishandle: hidden sheets, macro-enabled workbooks, Power Query refresh timing, Data Model boundaries, CUBEVALUE formulas, external links, client-ready `.xlsx` publishing, and Windows Excel COM validation.

Maintained by **Qiu Binbin (丘彬彬)**. WeChat: **binstudy**. Blog: **https://90le.cn**.

Current release: **v0.2.1** (`0.2.1+codex.20260714`). This release shortens the three starter prompts, makes all 12 skill descriptions trigger-only, and adds a 36-case trigger corpus plus three real plugin-eval benchmark scenarios. The 12 published skill IDs and Excel feature scope are unchanged.

## Use It When

- An AI agent needs to inspect or modify a real workbook with formulas, VBA, Power Query, Data Model, CUBE formulas, links, or hidden process sheets.
- A delivery workbook must be cleaned before it is shared with a client.
- A team wants one Excel BI workflow source that can be reused across Codex, Claude, OpenCode, and similar agents.
- A maintainer needs public validation that does not require private workbooks or desktop Excel.

## What It Helps Agents Do

| Area | What the skill pack adds |
| --- | --- |
| Excel/VBA workbook engineering | Export/import VBA modules, bind buttons, debug compile/runtime errors, preserve workbook structure |
| Power Query M | Read, edit, refresh, diagnose errors, track dependencies, and wait for refresh completion |
| Power Pivot / DAX | Review measures, context behavior, relationships, and Excel Data Model boundaries |
| MDX / CUBE formulas | Explain and audit `CUBEVALUE`, `CUBEMEMBER`, measures, members, and helper-cell references |
| ADO / SQL | Query workbook tables, files, and Data Model sources through ADO/OLEDB/ADOMD patterns |
| Client deliverables | Freeze formulas, remove links/queries/Data Model dependencies, delete config/process sheets |
| Workbook QA | Audit formulas, hidden sheets, controls, external dependencies, and delivery risk |
| Cross-agent installation | Sync the same skills into Codex, Claude, and OpenCode style folders |

## Install

### Option A: Codex Marketplace

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

### Option B: One-Command Local Install

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

Windows shortcuts:

```powershell
.\install.ps1
```

```cmd
install.cmd
```

macOS, Linux, and Git Bash:

```bash
sh install.sh
```

### Option C: Manual Install

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

## npm / npx Status

This repository does not publish an npm package yet, so it intentionally does **not** advertise a fake `npx` command. The current one-command installer is:

```bash
node tools/install.mjs
```

## Verify

Public validation, suitable for Windows, macOS, Linux, and Git Bash:

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

Full runtime validation requires Windows desktop Excel:

```powershell
python tools\run_release_gate.py --project-root .
```

Build the compact runtime package without development docs or duplicated agent mirrors:

```powershell
python tools\build_runtime_package.py --project-root . --out-dir "$env:TEMP\excel-bi-runtime" --zip "$env:TEMP\excel-bi-runtime.zip" --require-pass
```

The staged static plugin-eval comparison measured `trigger_cost_tokens` at 682, down from 1,161 by 41.26%, while `invoke_cost_tokens` moved from 15,365 to 14,886 (-479). This uses synthetic/generated package analysis and does not prove real task success; observed usage is separate evidence. See [Task recipes](docs/task-recipes.md) for the reproducible staging, analysis, comparison, and three-scenario benchmark commands.

## Compatibility Evidence

Compatibility is target-specific. Name the authoring, automation, consumer, and recipient environments, then report one of three levels: **structural evidence**, **runtime capability evidence**, or **workbook behavior evidence**. Windows desktop Excel can provide COM and provider capability probes; macOS, Excel for web, Linux, offline, legacy Office, and third-party spreadsheet targets require their own host-specific evidence. See [Compatibility boundaries](docs/compatibility.md) for the version, bitness, confidence, and target matrix.

## Included Skills

| Area | Skill |
| --- | --- |
| Routing | `excel-bi-router` |
| VBA and workbook automation | `excel-vba-workbook-engineering` |
| Power Query M | `power-query-m-engineering` |
| Power Pivot DAX | `power-pivot-dax-modeling` |
| MDX / CUBE formulas | `mdx-cubevalue-extraction` |
| ADO / SQL data access | `excel-ado-sql-data-access` |
| Clean Excel deliverables | `excel-deliverable-publisher` |
| Workbook QA | `excel-workbook-qa-auditor` |
| Report building | `excel-report-builder` |
| Office environment diagnostics | `office-environment-diagnostics` |
| Power BI semantic model context | `power-bi-semantic-model` |
| Sanitized test workbooks | `excel-testing-fixtures` |

## Documentation

- [English project overview](docs/project.en-US.md)
- [Chinese project overview](docs/project.zh-CN.md)
- [Install and sync guide](docs/install-and-sync.md)
- [Task recipes](docs/task-recipes.md)
- [Maintenance goals and risk backlog](docs/maintenance-goals.en-US.md)
- [Public growth goals](docs/growth-goals.en-US.md)
- [Repository governance goals](docs/repository-governance-goals.en-US.md)
- [Marketing copy pack](docs/marketing-copy.en-US.md)
- [Release notes](docs/release-notes.en-US.md)
- [Contributing guide](CONTRIBUTING.md)
- [Security policy](SECURITY.md)
- [Compatibility boundaries](docs/compatibility.md)
- [Distribution checklist](docs/distribution-checklist.md)
- [One-click install prompt EN](prompts/one-click-install-prompt.en-US.md)
- [One-click install prompt CN](prompts/one-click-install-prompt.zh-CN.md)
- [English website](docs/intro.html)
- [Chinese website](docs/intro.zh-CN.html)

## Boundaries

- Do not store customer workbooks, screenshots, PDFs, credentials, local machine paths, or generated QA reports inside the plugin package.
- `.agents/skills/` is the source of truth. `skills/`, `.claude/skills/`, and `.opencode/skills/` are generated mirrors.
- macOS and Linux can validate structure, prompts, OpenXML, and non-COM scripts. They do not prove Excel COM, VBA, Power Query refresh, or Power Pivot runtime behavior.
- This package improves agent operating discipline. It does not replace workbook-specific business review.

## License

MIT
