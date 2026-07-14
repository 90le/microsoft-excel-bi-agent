# Release Notes

## v0.2.0 - Capability-Aware Excel Compatibility And Compact Runtime

Release focus: make compatibility claims evidence-based across legacy, desktop, offline, Mac/web, Microsoft 365, and recipient environments while reducing the installed Codex runtime payload.

### Changed

- Added a stable Windows capability probe contract covering Excel COM, workbook roundtrip, VBA project access, Power Query object-model/wait behavior, Data Model access, PDF export, ADO/ACE, MSOLAP, and ADOMD.
- Added a cross-platform compatibility report that separates structural evidence, runtime capability evidence, and workbook behavior evidence; explicit capability requirements can block a handoff without treating an unavailable optional component as a package crash.
- Added capability-first routing: platform/host/runtime availability requests go to `office-environment-diagnostics`, while DAX and Power Pivot formula/function compatibility remains with `power-pivot-dax-modeling`.
- Added four synthetic capability fixtures and the `excel-capability-routing` sanitized regression case. No private workbook is required or shipped.
- Added an allowlisted, deterministic runtime-package builder with SHA-256 inventory, dependency closure, compact README, deterministic zip, and private-artifact exclusions.
- Added structural release checks for compatibility fixture/report behavior, runtime-package contents, release manifest version, and exactly three inspect/diagnose/publish starter prompts. The full Windows gate adds an optional live capability probe.
- Documented Windows, macOS, Excel for web, Linux, Excel 2007/2010/2013/2016/2019, Office LTSC, Microsoft 365, 32-bit/64-bit, offline, WPS/LibreOffice structural-only boundaries, and authoring/automation/consumer/recipient targets.
- Synchronized generated `skills/`, `.claude/skills/`, and `.opencode/skills/` mirrors from `.agents/skills/`.
- Bumped the plugin manifest to `0.2.0+codex.20260714`.

### Validation

```bash
python -m unittest discover -s tests -v
python tools/validate-skills.py .
python tools/run_release_gate.py --project-root . --profile structural
```

On Windows desktop Excel, the full gate also attempts the live capability probe:

```powershell
python tools\run_release_gate.py --project-root .
```

### Boundary

Structural evidence does not prove Excel execution. Runtime capability evidence applies only to the probed machine. High-confidence compatibility requires representative workbook behavior evidence in each materially different target environment; third-party spreadsheet applications remain structural-only until tested in that host.

## v0.1.5 - GitHub Community Health And Safe Intake

Release focus: reduce public collaboration risk by making issue intake, PR review, security reporting, and repository documentation surfaces safer.

### Changed

- Added bilingual repository governance goals with objective, constraints, boundaries, can-do/cannot-do rules, detailed goals, and high-value backlog.
- Added `CONTRIBUTING.md`, `SECURITY.md`, issue forms, and a pull request template.
- Added `tools/validate_github_community_health.py` and included it in `node tools/install.mjs --check` and GitHub Actions.
- Updated public checks across docs so community-health validation is part of the release gate.
- Bumped the plugin manifest to `0.1.5+codex.20260623175347`.

### Validation

Public checks:

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

### Boundary

This release changes GitHub governance and public intake safety. It does not claim new Excel COM, VBA, Power Query refresh, or Power Pivot runtime proof.

## v0.1.4 - Public Growth Goals And Marketing Readiness

Release focus: improve public trust, adoption clarity, social sharing, and marketing reuse without changing Excel workbook behavior.

### Changed

- Added bilingual public growth goals with objective, constraints, boundaries, can-do/cannot-do rules, detailed goals, and high-value optimization backlog.
- Added bilingual marketing copy pack for launch posts, short taglines, ad directions, channel variants, and do-not-claim rules.
- Added maintainer attribution across README, project docs, Pages, and plugin manifest: Qiu Binbin (丘彬彬), WeChat `binstudy`, blog `90le.cn`.
- Updated the website with use-case conversion cards, proof metrics, absolute Open Graph/Twitter image URLs, canonical URLs, author metadata, and v0.1.4 release visibility.
- Updated project documentation validation and goal coverage checks so public growth, marketing copy, maintainer signature, and social metadata stay covered by CI.
- Bumped the plugin manifest to `0.1.4+codex.20260623173419`.

### Validation

Public checks:

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

Maintainer structural gate:

```bash
python tools/run_release_gate.py --project-root . --profile structural
```

### Boundary

This release changes public positioning, site layout, metadata, and documentation. It does not claim new Excel COM, VBA, Power Query refresh, or Power Pivot runtime proof.

## v0.1.3 - Public Maintenance Goals And CI Validation

Release focus: reduce public repository maintenance risk without changing Excel workbook behavior.

### Changed

- Added public maintenance goals, constraints, boundaries, can-do/cannot-do rules, detailed goals, risk register, and optimization backlog.
- Added GitHub Actions public structural validation for pushes and pull requests.
- Reworked project-document validation so it checks public repository docs instead of maintainer-only local ledgers.
- Reworked goal coverage and completion readiness audits around public maintenance coverage and active backlog state.
- Expanded `node tools/install.mjs --check` so it runs the full public structural check set.
- Updated the website to expose release, public validation, and runtime boundary status.
- Updated artifact hygiene expectations so the public package contains no Office workbooks by default.
- Updated task recipes, distribution, open-source publishing, and real/sanitized regression docs to remove stale local-evidence assumptions.
- Bumped the plugin manifest to `0.1.3+codex.20260623171436`.

### Validation

Public checks:

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

Maintainer structural gate:

```bash
python tools/run_release_gate.py --project-root . --profile structural
```

### Boundary

This release does not claim new Excel COM, VBA, Power Query refresh, or Power Pivot runtime proof. Those checks still require Windows desktop Excel and task-specific runtime evidence outside the public repository.

## v0.1.2 - Split Bilingual Docs And Site

- Split README and project overview into independent English and Chinese entry points.
- Split the website into English and Chinese pages with browser-language redirect.
- Kept install commands limited to supported Codex marketplace and local installer paths.
