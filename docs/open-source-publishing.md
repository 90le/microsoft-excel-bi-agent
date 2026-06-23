# Open Source Publishing Notes

This project is intended to be published as a reusable local plugin and cross-agent skill pack.

## Recommended Repository Name

```text
microsoft-excel-bi-agent
```

## Public Scope

Include:

- `.codex-plugin/plugin.json`
- `.agents/skills/`
- generated skill mirrors that are useful for recipients: `skills/`, `.claude/skills/`, `.opencode/skills/`
- `tools/`
- `fixtures/`
- `prompts/`
- recipient-facing docs and setup docs
- `README.md`
- `LICENSE`

Exclude or regenerate locally:

- local release ledgers with machine paths
- generated PDF/image evidence
- task-local generated smoke fixture folders
- Python caches
- Excel lock files
- private customer workbooks or screenshots

## Release Checklist

Before publishing:

```powershell
python tools\validate-skills.py .
python tools\validate_project_docs.py --project-root .
python tools\validate_task_recipes.py --project-root .
python tools\validate_official_docs_index.py --project-root .
python tools\build_artifact_hygiene_report.py --project-root . --require-pass
python tools\build_goal_coverage_report.py --project-root . --require-pass
node tools\install.mjs --check
```

Run `python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift` after skill source changes. Run `python tools\run_release_gate.py --project-root . --profile structural` when preparing a deeper maintainer validation pass.

For a public repository staging copy, also scan staged files for machine paths:

```powershell
rg -n "C:\\Users|AppData|Documents\\Codex|token|secret|password|api[_-]?key" .
```

False positives are expected in security scanners and documentation that describes credential checks. Private absolute paths should be removed or excluded.

## GitHub Publishing

Create a public repository, then from the clean staging directory:

```powershell
git init
git branch -M main
git add .
git commit -m "Initial open source release"
git remote add origin https://github.com/<owner>/microsoft-excel-bi-agent.git
git push -u origin main
```

If GitHub CLI is available and authenticated:

```powershell
gh repo create <owner>/microsoft-excel-bi-agent --public --source . --remote origin --push
```

## Boundary Statement

The public repository can prove source structure and static validation. It cannot prove a private workbook calculates correctly. Excel COM, VBA execution, Power Query refresh, and Power Pivot/Data Model runtime behavior require Windows desktop Excel and local validation.

