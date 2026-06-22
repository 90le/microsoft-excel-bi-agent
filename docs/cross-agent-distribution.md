# Cross-Agent Distribution

This document defines the constrained distribution model for Codex, Claude, OpenCode, and compatible agents.

## Rule

`.agents/skills/` is the only source of truth.

Generated mirrors are disposable:

- `skills/`
- `.claude/skills/`
- `.opencode/skills/`
- `~/.codex/skills/`

Do not edit generated mirrors directly.

## One-Command Sync

PowerShell:

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Bash:

```bash
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

This updates:

- Codex plugin skill mirror: `skills/`
- Codex user skills: `~/.codex/skills/`
- Claude project skills: `.claude/skills/`
- OpenCode project skills: `.opencode/skills/`

For `~/.codex/skills/`, unrelated system or personal skills are allowed to remain. Drift checks validate only this package's 12 skill folders inside that shared user directory.

## Drift Check

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
```

Use this before release or after any skill edit.

## Codex Plugin Install

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
```

Use cachebuster only when behavior changed:

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install --update-cachebuster
```

Do not edit `~/.agents/plugins/marketplace.json` by hand.

## Metadata Requirement

Every canonical skill must keep:

```text
agents/openai.yaml
```

Required fields:

- `interface.display_name`
- `interface.short_description`
- `interface.default_prompt`

The default prompt must mention the matching `$skill-name`.

## What Not To Claim

- Mirror sync proves distribution consistency, not skill quality.
- Generated forward-test prompts prove coverage framing, not real external-agent performance.
- Generated response stubs prove scaffolding only, not execution.
- Real external-agent behavior requires fresh-session outputs and separate scoring.

## Validation

```powershell
python tools\validate-skills.py .
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
python tools\run_release_gate.py --project-root . --profile structural
```
