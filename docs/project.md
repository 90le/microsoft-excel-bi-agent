# Project

## Objective

Microsoft Excel BI Agent is an open-source cross-agent Excel BI skill pack for agents that work on Excel workbooks, VBA, Power Query M, Power Pivot DAX, MDX/CUBE formulas, ADO/SQL, workbook QA, clean deliverables, report building, Office diagnostics, and sanitized fixtures.

Current release: `0.1.0+codex.20260622060709`.

## Start Here

- Recipient README: `README.md`
- Recipient guide: `docs/recipient-guide.zh-CN.md`
- One-click install prompt: `prompts/one-click-install-prompt.zh-CN.md`
- HTML introduction page: `docs/intro.html`
- Install and sync: `docs/install-and-sync.md`
- Distribution checklist: `docs/distribution-checklist.md`
- Task recipes: `docs/task-recipes.md`
- Case regression: `docs/real-case-regression.md`
- Open-source publishing notes: `docs/open-source-publishing.md`

## Canonical Structure

```text
microsoft-excel-bi-agent/
  README.md
  LICENSE
  .codex-plugin/plugin.json
  .agents/skills/          # source of truth
  skills/                  # generated Codex plugin mirror
  .claude/skills/          # generated Claude mirror
  .opencode/skills/        # generated OpenCode mirror
  docs/
  fixtures/
  prompts/
  tools/
```

## Public Repository Boundary

The public repository contains source skills, scripts, fixtures, install prompts, and recipient-facing documentation. Maintainer-only release ledgers and machine-specific runtime reports are intentionally not included.

Windows desktop Excel is required for Excel COM, VBA execution, Power Query refresh, and Power Pivot/Data Model runtime validation. macOS and Linux support structural checks, OpenXML inspection, documentation, and non-COM scripts.

## Validation

```powershell
python tools\validate-skills.py .
python <plugin-creator-skill-root>\scripts\validate_plugin.py .
python tools\build_artifact_hygiene_report.py --project-root . --require-pass
```
