# One-Click Install Prompt

Copy the prompt below into Codex, Claude, OpenCode, or another local-file-capable agent.

```text
You are my local plugin installation assistant. Please install Microsoft Excel BI Agent.

Goals:
1. Install or refresh the Codex plugin.
2. Sync skills for Codex, Claude, and OpenCode.
3. Run structural validation.
4. Summarize install results, skipped checks, and any failures.

Plugin location:
Use the currently opened microsoft-excel-bi-agent directory. If I provide a zip file, unzip it first, then enter the extracted microsoft-excel-bi-agent directory.

Preferred install command:
node tools/install.mjs

If Node is unavailable, use the manual commands.

Windows PowerShell:
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace

Git Bash / macOS / Linux:
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace

Structural validation:
python tools/validate-skills.py .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
node tools/install.mjs --check

If the environment is Windows with desktop Excel installed, also run the full runtime gate:
python tools\run_release_gate.py --project-root .

Required boundaries:
- Do not manually edit generated mirrors: skills/, .claude/skills/, .opencode/skills/.
- Do not delete unrelated user skills under ~/.codex/skills.
- Do not write customer workbooks, screenshots, PDFs, credentials, or private paths into the plugin directory.
- If this is not Windows with desktop Excel, do not claim that Excel COM, VBA, Power Query refresh, or Power Pivot runtime behavior was validated.
- If full Excel validation is skipped, state why.

Output:
- Tell me whether the plugin install succeeded.
- Tell me whether Codex / Claude / OpenCode skills were synced.
- Tell me whether public structural validation passed.
- If anything failed, report the failing command, error summary, and recommended next step.
```

## Short Version

```text
Enter the microsoft-excel-bi-agent directory, run:
node tools/install.mjs
node tools/install.mjs --check
Then summarize the result. Do not manually edit generated skills mirrors.
```
