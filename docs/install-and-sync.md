# Install And Sync

This file is the constrained install path for `microsoft-excel-bi-agent-pack`.
Do not create new install flows unless one of these commands cannot support the target runtime.

For a recipient who receives the whole plugin folder, start with `README.md` or `docs/recipient-guide.zh-CN.md`, then use the commands in this file. If another agent will perform the installation, copy `prompts/one-click-install-prompt.zh-CN.md` to that agent.

## Source Of Truth

```text
.agents/skills/
```

Generated mirrors:

- `skills/` for Codex plugin packaging.
- `.claude/skills/` for Claude project usage.
- `.opencode/skills/` for OpenCode project usage.
- `~/.codex/skills/` only when user-level Codex skills are explicitly wanted.

Do not edit generated mirrors directly.

## One-Command Prompt Sync

PowerShell:

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Bash:

```bash
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

This syncs Codex plugin mirror, Claude project skills, OpenCode project skills, and Codex user skills from the canonical `.agents/skills/` source.

`~/.codex/skills/` is a shared user directory. Drift checks validate this package's 12 skill folders there and intentionally ignore unrelated system or personal skills.

## Codex Plugin Install

Use this when the Codex plugin itself must be installed or refreshed:

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
```

Use `--update-cachebuster` only when behavior changed and a new installed plugin version is required:

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install --update-cachebuster
```

Do not hand-edit `~/.agents/plugins/marketplace.json`.

## Checks

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
python tools\validate_project_docs.py --project-root .
python tools\build_artifact_hygiene_report.py --project-root . --require-pass
```

For a release candidate:

```powershell
python tools\run_release_gate.py --project-root . --profile structural
```

Run the full gate only when Excel COM behavior, plugin installation, or provider behavior changed.

## Hard Boundaries

- No customer workbooks, screenshots, PDFs, generated QA reports, credentials, or private local paths in the plugin package.
- No divergent manual edits in `skills/`, `.claude/skills/`, `.opencode/skills/`, or `~/.codex/skills/`.
- No new skills or broad tools unless a repeated workflow cannot be handled by an existing skill plus script.
- Linux/macOS structural checks are not proof of Excel COM refresh, VBA execution, Data Model behavior, or rendered Excel output.
