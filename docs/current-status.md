# Current Status

This is the short operational entry point for Microsoft Excel BI Agent.
Use `docs/maintenance-goals.en-US.md` or `docs/maintenance-goals.zh-CN.md` for the public maintenance goals, risk register, and optimization backlog.

## Current Release

```text
v0.1.3
```

Package cachebuster version:

```text
0.1.3+codex.20260623171436
```

## Current Capability Shape

- 12 canonical skills.
- 76 package tools/scripts, including task profiles, the real/sanitized case regression runner, visual QA fixture tools, and Excel COM render-evidence export.
- 14 cataloged workflows.
- 53 release-gate check functions.
- Core Excel BI skills, six upper-layer scenario skills, and maintenance task profiles are complete.
- The real/sanitized case regression library V1 is complete for `0.1.0+codex.20260622033808`.
- The workbook-backed sanitized Visual QA case V1 is complete for `0.1.0+codex.20260622045441`.
- The rendered Visual QA evidence chain V1 is complete for `0.1.0+codex.20260622060709`.
- Public maintenance goals, risk backlog, and CI-backed structural validation are complete for `v0.1.3`.

## Daily Entry Points

| Need | Start Here |
|---|---|
| Give the pack to another user | `README.md`, then `docs/distribution-checklist.md` |
| Explain the pack to a recipient | `docs/recipient-guide.zh-CN.md` or `docs/intro.html` |
| Ask another agent to install it | `prompts/one-click-install-prompt.zh-CN.md` |
| Install/sync all prompts | `docs/install-and-sync.md` |
| Review maintenance goals and remaining risks | `docs/maintenance-goals.en-US.md` or `docs/maintenance-goals.zh-CN.md` |
| Review release notes | `docs/release-notes.en-US.md` or `docs/release-notes.zh-CN.md` |
| Choose a skill or workflow | `excel-bi-router` or `tools/run_task_profile.py` |
| Audit a workbook | `tools/run_task_profile.py --profile audit` |
| Publish a pure deliverable | `tools/run_task_profile.py --profile publish` |
| Refresh Power Query | `tools/run_task_profile.py --profile pq-refresh` |
| Review DAX/model dependencies | `tools/run_task_profile.py --profile dax-review` |
| Run real/sanitized case regression | `tools/run_task_profile.py --profile case-regression --execute` |
| Generate/check visual QA fixture | `tools/create_visual_qa_fixture.py` then `tools/build_visual_qa_report.py` |
| Export rendered Visual QA evidence | `tools/export_visual_qa_render_evidence.ps1 -CreateFixture` for smoke tests, or pass a sanitized workbook and task-local output paths |
| Run package validation | `tools/run_task_profile.py --profile release-structural` or `--profile release-full` |

## Current Boundaries

- The V1 case library validates regression definitions, coverage, safety boundaries, and package-tool references.
- The visual QA fixture validates static workbook-backed readability risks; it does not render pixels.
- The rendered Visual QA evidence chain proves only that desktop Excel COM opened a sanitized workbook and exported visible `Report*` sheets to PDF on that Windows machine.
- It does not prove that a private workbook calculates correctly.
- Customer workbook reports, generated PDFs, machine-specific evidence, screenshots, credentials, and local paths must stay outside the plugin package.

## Maintenance Rule

Constrain before expanding. Prefer shorter entry docs, fixed install commands, existing skills/scripts, and regression evidence. Add new skills or broad tools only when a repeated workflow cannot be handled by the current package.
