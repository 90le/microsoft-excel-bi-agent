# Release Notes

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
