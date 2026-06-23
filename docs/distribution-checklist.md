# Distribution Checklist

Use this before sending Microsoft Excel BI Agent to another user or another machine.

GitHub Actions runs public structural validation on pushes and pull requests. These local checks mirror that public CI boundary and do not require desktop Excel.

## Package Shape

- Send the whole plugin folder, not only `skills/`.
- Keep `.codex-plugin/plugin.json`.
- Keep `.agents/skills/` as the source of truth.
- Keep generated mirrors only if they were produced by `tools/sync-skills.py`.
- Include `README.md` and `docs/install-and-sync.md`.
- Include `docs/recipient-guide.zh-CN.md`, `prompts/one-click-install-prompt.zh-CN.md`, and `docs/intro.html` for recipient onboarding.

## Required Pre-Send Checks

```powershell
python tools\validate-skills.py .
python tools\validate_project_docs.py --project-root .
python tools\validate_task_recipes.py --project-root .
python tools\validate_official_docs_index.py --project-root .
python tools\build_artifact_hygiene_report.py --project-root . --require-pass
python tools\build_goal_coverage_report.py --project-root . --require-pass
node tools\install.mjs --check
```

Run this additional drift check after changing `.agents/skills/`:

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
```

Run the full gate only on a Windows machine with desktop Excel when Excel runtime behavior changed:

```powershell
python tools\run_release_gate.py --project-root .
```

## Do Not Include

- Customer workbooks.
- Screenshots.
- PDFs or rendered visual evidence.
- Generated QA reports.
- Credentials or tokens.
- Private local paths.
- `__pycache__` folders.
- Excel lock files such as `~$*.xlsx`.
- Temporary zip/build folders.

## Recipient Install Commands

Codex plugin:

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
```

Prompt/skill mirrors:

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Drift check:

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
```

## Support Boundary

- Windows + desktop Excel is required for Excel COM, VBA execution, Power Query refresh, and Data Model runtime validation.
- macOS/Linux/Git Bash support covers structural checks, skill sync, docs, OpenXML inspection, and non-COM scripts.
- A successful structural gate does not prove a private workbook calculates correctly.
