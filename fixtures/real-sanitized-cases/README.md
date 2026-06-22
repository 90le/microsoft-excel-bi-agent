# Real/Sanitized Case Regression Library

This folder stores reusable regression case definitions learned from real Excel BI work.
It must not contain customer workbooks, screenshots, credentials, private paths, or business-specific rules.

## Goal

- Preserve high-risk workbook problem shapes as sanitized case specs.
- Give future agents a checklist before changing Power Query, DAX, CUBE formulas, VBA, deliverable cleanup, or visual QA behavior.
- Keep each case small, deterministic, and explicit about what it does not prove.

## Boundary

- Case specs can reference package tools and expected evidence.
- Case specs do not prove a private workbook is correct.
- Live workbook proof belongs in task-local temp output or a separate sanitized workbook supplied by the user.

Validate with:

```powershell
python tools\run_case_regression.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_case_regression.json" `
  --out-md "$env:TEMP\excel_bi_case_regression.md" `
  --require-pass
```
