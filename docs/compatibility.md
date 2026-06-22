# Compatibility

## Supported Agents

The pack targets the common `AGENTS.md` + `SKILL.md` model used by modern coding agents. It avoids making any agent-specific metadata required.

| Agent | Primary entry | Notes |
|---|---|---|
| Codex | `AGENTS.md`, `.agents/skills/*/SKILL.md` | Can also install skills under `~/.codex/skills`. |
| Codex plugin | `.codex-plugin/plugin.json`, `skills/*/SKILL.md` | `skills/` is a generated mirror from `.agents/skills`. |
| Claude | `SKILL.md`, `.claude/skills` if synced | Keep source canonical in `.agents/skills`. |
| OpenCode | `AGENTS.md`, `.agents/skills`, `.opencode/skills` if synced | Supports OpenCode rules and skills. |
| Other agents | `AGENTS.md` and Markdown instructions | Use scripts manually if skills are not auto-discovered. |

## Skill Mirror Policy

- `.agents/skills` is the source of truth.
- `skills/` is generated for Codex plugin packaging.
- `.claude/skills` and `.opencode/skills` are optional generated project mirrors.
- `~/.codex/skills` is an optional generated user-level mirror.
- Use `tools/sync-skills.py --check-drift` before claiming any mirror is current.

## Platform Matrix

| Platform | Excel COM | OpenXML inspection | VBA import/export | Macro execution |
|---|---:|---:|---:|---:|
| Windows PowerShell + Excel | yes | yes | yes | yes |
| Windows Git Bash + Excel | via wrapper | yes | via wrapper | via wrapper |
| Linux | no | yes | no VBE automation | no |
| macOS | limited/non-COM | yes | no Windows VBE automation | only manual Excel validation |

## Rule

Never report Linux/macOS structural checks as proof that Excel VBA, Power Query refresh, Power Pivot, Solver, or button clicks work in desktop Excel.

Use the structural release gate on Linux/macOS or non-Excel environments:

```bash
tools/run_release_gate.sh --profile structural
```

Structural mode validates package shape, skill files, script syntax where the local shell can parse it, official documentation indexes, mirror drift, OpenXML/CUBE formula parser fixtures, and model-report fixtures. It intentionally skips Excel process checks, local installed-plugin cache validation, and `codex plugin list`.

Use the portable structural helper when an agent needs a smaller static proof point rather than the full release gate:

```bash
tools/excel_bi_structural.sh sanitized-bundle --out-dir /tmp/excel_bi_sanitized_fixtures --clean --validate
tools/excel_bi_structural.sh pq-lineage --query-dir src/m --out-json tmp/pq-lineage.json --out-md tmp/pq-lineage.md --fail-on-high-risk
tools/excel_bi_structural.sh provider-baseline-fixture --out-dir /tmp/excel_bi_provider_baseline_fixture --clean
```

This helper validates OpenXML/static-source evidence and synthetic provider-baseline comparison behavior only. It is suitable for Git Bash, Linux, and macOS, and it must not be reported as proof of Excel COM, Power Query refresh, VBA execution, Power Pivot calculation, ADO workbook SQL, ADOMD endpoint query behavior, or live Office provider availability.

## CUBE / MDX Specifics

- `build_cube_dependency_report.py` can map CUBE formulas to sheets, cells, measures, member references, and helper cell references without Excel.
- Missing-measure checks are strongest when paired with Windows Excel COM Data Model export.
- CUBE dependency reports are structural and do not prove that `CUBEVALUE` returns the expected value after refresh.

## ADO / OLEDB Specifics

- `test_excel_ado_sql_access.ps1` requires Windows and a compatible ADO/OLEDB provider such as `Microsoft.ACE.OLEDB.12.0`.
- Creating the built-in Excel fixture requires desktop Excel COM.
- Querying a saved workbook through ACE OLEDB does not prove that unsaved workbook changes, Power Query refreshes, or Data Model calculations are current.
- Linux/macOS can draft SQL and inspect OpenXML workbook structure, but cannot validate ACE OLEDB provider behavior.

## Power Query Specifics

- Exact `Workbook.Queries` formula export requires Windows desktop Excel COM.
- Query add/update/delete through `Workbook.Queries` requires Windows desktop Excel COM.
- Refresh completion waiting uses Excel object model behavior and must be validated in desktop Excel.
- `inspect_power_query_openxml.py` can detect `connections.xml`, query tables, external links, custom XML, and mashup-like package parts without Excel.
- OpenXML inspection does not fully decode Excel DataMashup binaries and does not refresh queries.
