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

## Evidence And Confidence Contract

Every compatibility statement must name its target and one of these evidence tiers:

1. **Structural evidence**: file/package inspection, source lint, schema checks, synthetic fixtures, or generated plans. It can show that a workbook or formula has a compatible shape; it cannot prove that Excel executed it.
2. **Runtime capability evidence**: a capability probe on one identified machine and host, such as Excel COM activation, workbook roundtrip, provider activation, Power Query object-model access, Data Model access, or PDF export. It applies only to the probed environment.
3. **Workbook behavior evidence**: a representative workbook was opened and the required refresh, macro, model calculation, rendered output, or user path was exercised in the target host. Business correctness still needs domain review.

Use explicit confidence labels:

| Confidence | Minimum evidence | Meaning |
|---|---|---|
| `low` | Structural evidence only | Candidate-compatible; runtime behavior is unknown. |
| `medium` | Runtime capability evidence on the named environment | The machine/host can perform the operation; the workbook itself may still fail. |
| `high` | Workbook behavior evidence in the intended target environment | The tested workbook path worked under the recorded conditions; do not generalize to other recipients or versions. |

`Unsupported`, `blocked`, and `unknown` are results, not low-confidence success. Missing capability evidence must remain unknown.

## Compatibility Targets

Record each distinct target before implementation:

| Target | Question to answer |
|---|---|
| **Authoring target** | Where will formulas, queries, VBA, model objects, and report surfaces be created or edited? |
| **Automation target** | Which machine, Office build, bitness, policy, providers, and credentials will execute automation? |
| **Consumer target** | Which Excel host will open, refresh, calculate, and interact with the workbook? |
| **Recipient target** | What environment will each recipient actually receive, including offline/cloud restrictions? |

The agent's execution environment is not automatically any of these targets. Evidence from the automation target must not be transferred to a recipient target without verification.

## Platform And Host Matrix

| Platform/host | Structural evidence | Runtime capability evidence | Workbook behavior evidence |
|---|---|---|---|
| Windows desktop Excel + PowerShell | Yes | Full capability probe, Excel COM, providers, bitness, Trust Center policy | Automated representative-workbook tests are available |
| Windows Git Bash + desktop Excel | Yes | Via the packaged PowerShell wrappers | Same Windows host, with wrapper commands recorded |
| macOS desktop Excel | Yes | No Windows COM proof; use host-specific/manual evidence | Validate manually or with a supported Mac automation path |
| Excel for web | Package/source inspection only | No desktop COM/VBE/provider proof | Test the workbook in Excel for web, including refresh and unsupported-feature behavior |
| Linux | OpenXML/source/synthetic checks | No Microsoft desktop Excel runtime probe | No Excel workbook behavior proof without a separate supported host |
| Third-party spreadsheet applications such as WPS or LibreOffice | **Structural compatibility only** until tested | Microsoft Excel capability probes do not apply | Require application-specific workbook behavior evidence |

## Excel Version Policy

Version labels describe support posture, not proof. Use the exact target build when possible.

| Excel family | Policy |
|---|---|
| Excel 2007 | Structural inspection and conservative workbook compatibility only; Power Query and modern Data Model automation must be treated as unavailable unless separately demonstrated. |
| Excel 2010 | Best effort; Power Query/Power Pivot may depend on separately installed legacy add-ins and cannot be inferred from the version number. |
| Excel 2013 | Best effort; probe actual add-ins, object model, providers, and workbook behavior. |
| Excel 2016 | Legacy-compatible target; use lifecycle warnings and probe the exact MSI/C2R build rather than forcing an upgrade. |
| Excel 2019 | Legacy-compatible target; modern Microsoft 365 functions and service behavior are not implied. |
| Office LTSC | Supported as a fixed-feature target when the exact LTSC release/build and capability probe are recorded. |
| Microsoft 365 | Active channel target; channel/build drift means a successful probe is time- and machine-specific. |

Excel 2007/2010/2013/2016/2019, LTSC, and Microsoft 365 can expose different functions, object models, providers, and security defaults. DAX or Power Pivot formula/function compatibility stays with `power-pivot-dax-modeling`; platform, host, run, support, and availability compatibility belongs to `office-environment-diagnostics`.

## Bitness, Offline, And Policy Profiles

- Record Windows OS bitness, Office **32-bit** or **64-bit**, the agent process bitness, and provider bitness. ACE/OLEDB, COM, and add-in readiness can fail when architectures differ.
- Treat **offline** and cloud-prohibited operation as first-class recipient policies. Local files, cached credentials, gateways, web connectors, SharePoint/OneDrive, licensing, and sign-in-dependent features must be assessed separately.
- An offline structural gate can validate package shape and source rules. It cannot prove cloud-source refresh or subscription-backed functionality.
- Trust Center access to the VBA project object model, macros, protected view, signed code, and external-content policy are environment evidence, not workbook-format facts.

## Capability Workflow

1. Classify authoring, automation, consumer, and recipient targets.
2. Run structural checks everywhere.
3. On Windows, capture `tools/probe_excel_capabilities.ps1` output or use an explicitly supplied captured probe.
4. Build the compatibility report with `tools/build_excel_compatibility_report.py`; add `--require-capability` only for operations the task truly requires.
5. Route implementation to the specialist skill, then gather workbook behavior evidence in every materially different target environment.

Never report Linux/macOS/web/third-party structural checks as proof that Excel VBA, Power Query refresh, Power Pivot, Solver, provider access, PDF rendering, or button clicks work in desktop Excel.

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
