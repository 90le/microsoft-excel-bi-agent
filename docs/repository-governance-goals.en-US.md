# Repository Governance Goals

This page defines the public GitHub governance target for Microsoft Excel BI Agent. It covers issue intake, pull-request review, security reporting, community health files, and repository settings that affect public trust.

Maintainer: **Qiu Binbin (丘彬彬)**<br>
WeChat: **binstudy**<br>
Blog: **https://90le.cn**

## Objective

Make the GitHub repository safe to participate in without increasing the chance that users upload customer workbooks, screenshots, PDFs, credentials, local private paths, generated QA reports, or unsanitized runtime evidence.

The goal is not more process. The goal is a narrow public collaboration path that protects users, keeps install claims truthful, and keeps every contribution reviewable with the existing public checks.

## Constraints

- Public issue templates must discourage customer files and private artifacts before the user writes a report.
- PR templates must require validation results, bilingual-document awareness, and source/mirror discipline.
- Security and sensitive-data reports must not be routed through public issues when private details are involved.
- GitHub Wiki must remain disabled because `docs/` and GitHub Pages are the maintained documentation surfaces.
- GitHub Discussions can remain disabled until there is real community volume to justify moderation.
- New governance files must stay concise and should not create a separate process that conflicts with existing validation.

## Boundaries

- Repository governance may add community-health files, issue templates, PR templates, security guidance, and validation scripts.
- Repository governance may update GitHub repository settings when they reduce drift or leakage risk.
- Repository governance does not replace workbook-specific business review.
- Repository governance does not add runtime proof for Excel COM, VBA, Power Query refresh, or Power Pivot.

## Can Do

- Add issue forms that require sanitized reproduction steps and environment boundaries.
- Add a PR checklist covering public checks, bilingual docs, install truth, and artifact hygiene.
- Add `SECURITY.md` and `CONTRIBUTING.md` for public contributor orientation.
- Disable Wiki to prevent unreviewed parallel documentation.
- Add CI validation for community-health files and template safety language.

## Cannot Do

- Ask users to attach customer workbooks, screenshots, PDFs, or generated QA reports in public issues.
- Accept PRs that only update generated skill mirrors while leaving `.agents/skills/` unchanged.
- Accept install or marketing changes that mention unsupported npm/npx commands.
- Treat a public issue template as permission to share private workbook data.
- Put maintainer-only runtime evidence into public repository governance files.

## Detailed Goals

| Goal | Why it matters | Done when |
| --- | --- | --- |
| Safe issue intake | Public issues are the most likely place for accidental workbook/data leakage. | Issue forms require sanitized repros and prohibit private artifacts. |
| PR review discipline | Open-source changes need the same release discipline as maintainer changes. | PR template requires validation commands, bilingual docs awareness, and source/mirror checks. |
| Security routing | Sensitive reports need a private path. | `SECURITY.md` tells users not to disclose secrets or private workbook details in public issues. |
| Documentation source control | Wiki pages bypass normal review and can drift from Pages. | Repository Wiki is disabled and docs point to `docs/`/Pages. |
| CI-backed governance | Governance files should not silently regress. | Public checks include community-health validation. |

## High-Value Backlog

| Priority | Item | Boundary |
| --- | --- | --- |
| P0 | Add issue forms, PR template, `SECURITY.md`, and `CONTRIBUTING.md`. | Keep forms short; no customer data requests. |
| P0 | Validate community-health files in CI. | Validation is structural and does not call GitHub APIs. |
| P0 | Disable Wiki. | Docs remain in repository files and Pages. |
| P1 | Add GitHub Discussions later only if real community volume appears. | Do not create an unmoderated support forum prematurely. |
| P2 | Add a short contributor quickstart video later. | Must use sanitized fixtures only. |

## Required Public Checks

Run these after changing governance files, issue templates, PR templates, docs, install flows, or validation scripts:

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
