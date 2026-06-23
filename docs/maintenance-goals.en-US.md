# Maintenance Goals

This page defines the public maintenance target for Microsoft Excel BI Agent. It is a repository-facing contract for future changes, risk reviews, and release decisions. It does not contain maintainer-only runtime ledgers or machine-specific evidence.

## Objective

Keep Microsoft Excel BI Agent reliable for real Excel BI work while making the public repository easier to install, validate, review, and extend across Codex, Claude, OpenCode, and similar agents.

The goal is not to make agents edit Excel files freely. The goal is to force safer operating discipline: identify the workbook surface, choose the correct BI layer, make narrow changes, validate the result, and state exactly which runtime checks were performed or skipped.

## Must-Worthy Optimization Rule

Only prioritize changes that reduce release risk, prevent false installation claims, protect user data, improve public validation, or make the package easier for another agent or maintainer to operate. Avoid cosmetic churn, broad rewrites, new skills without repeated demand, or release artifacts that would need private evidence inside the public repository.

## Constraints

- Public install commands must be executable by the current project. Do not advertise npm or npx commands unless an npm package is actually published.
- English and Chinese user-facing docs must stay independently maintained. Do not mix bilingual long-form content into a single README or site page.
- `.agents/skills/` is the source of truth for skills. `skills/`, `.claude/skills/`, and `.opencode/skills/` are generated mirrors.
- Windows desktop Excel is required for Excel COM, VBA execution, Power Query refresh, and Power Pivot runtime behavior.
- macOS and Linux validation proves only structure, OpenXML inspection, documentation, and non-COM script behavior.
- Customer files, screenshots, PDFs, credentials, local private paths, and generated QA reports must remain outside the public package.

## Boundaries

- This repository ships source skills, generated mirrors, scripts, docs, prompts, and sanitized fixtures.
- This repository does not ship customer workbooks, private release ledgers, machine-specific runtime reports, or local Excel evidence.
- Structural validation does not prove that a private workbook calculates correctly.
- Release notes may summarize runtime evidence, but the raw runtime artifacts should stay in task-local or release-local storage unless deliberately sanitized.

## Can Do

- Improve public docs, website pages, install instructions, and validation commands.
- Update `.agents/skills/` when a repeated Excel BI workflow needs a clearer agent procedure.
- Sync generated skill mirrors after source skill changes.
- Add sanitized fixtures, regression cases, and static OpenXML checks.
- Add CI checks that run public structural validation without requiring desktop Excel.
- Tighten artifact hygiene and documentation consistency checks.

## Cannot Do

- Commit customer workbooks, screenshots, PDFs, generated QA reports, credentials, or private local paths.
- Claim Excel COM, VBA, Power Query refresh, or Power Pivot behavior was validated without Windows desktop Excel evidence.
- Add fake npm, npx, Claude, OpenCode, or marketplace commands that are not supported by the project.
- Manually edit generated skill mirrors while leaving `.agents/skills/` unchanged.
- Treat a structural gate as workbook-specific business validation.

## Detailed Goals

| Goal | Why it matters | Done when |
| --- | --- | --- |
| Install truth | Adoption fails quickly when install docs contain non-working commands. | README, website, and install docs expose only real install paths. |
| Bilingual independence | Mixed pages are hard to review and break search/user expectations. | English and Chinese docs can be edited and reviewed independently. |
| Skill source discipline | Mirror drift creates different behavior across agents. | Source skills are edited first and mirror drift checks pass. |
| Public validation | Contributors need fast checks that do not require private files or Excel COM. | Public checks pass locally and in GitHub Actions. |
| Runtime boundary clarity | Users must know what macOS/Linux checks cannot prove. | Docs and release notes separate structural checks from Windows Excel runtime checks. |
| Artifact hygiene | Public repos can leak sensitive workbook artifacts by accident. | Hygiene checks fail on customer files, local reports, lock files, and private paths. |
| Risk backlog visibility | Maintainers need a shared list of valuable next improvements. | High-risk and high-value backlog items are documented with clear boundaries. |

## Risk Register

| Risk | Severity | Mitigation |
| --- | --- | --- |
| Public docs mention commands that do not exist. | High | Validate install docs and keep `node tools/install.mjs` as the local one-command path. |
| Maintainer-only evidence leaks into the public package. | High | Keep release ledgers and generated reports ignored; run artifact hygiene before release. |
| Generated skill mirrors drift from `.agents/skills/`. | High | Run `tools/sync-skills.py --check-drift` before release and after skill edits. |
| Linux/macOS users over-trust structural validation. | High | Repeat the Windows Excel runtime boundary in docs and validation output. |
| Release confidence depends on manual local checks only. | Medium | Run public structural checks in GitHub Actions on pushes and pull requests. |
| Goal/roadmap docs become stale after release. | Medium | Keep this page tied to project docs validation and review it before tagged releases. |

## Optimization Backlog

| Priority | Item | Boundary |
| --- | --- | --- |
| P0 | Keep public validation green in CI. | CI must not require desktop Excel or private artifacts. |
| P0 | Keep install commands truthful across README, site, and install docs. | No npm/npx claim until a package is published. |
| P1 | Continue expanding sanitized regression cases from recurring workbook failures. | Cases must not reveal customer data or local paths. |
| P1 | Keep website release and validation visibility current. | Do not embed raw machine evidence or generated QA reports. |
| P1 | Add clearer contributor guidance for source skills versus generated mirrors. | Do not encourage manual mirror edits. |
| P2 | Add optional deeper Windows Excel runtime release evidence templates. | Templates may describe evidence shape; raw evidence remains outside the repo. |

## Required Public Checks

Run these after public docs, install flow, validation scripts, or skill packaging changes:

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

If plugin structure changes, also run the plugin validator from the Codex plugin creator skill.
