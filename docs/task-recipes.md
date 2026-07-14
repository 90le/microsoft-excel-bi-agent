# Task Recipes

These recipes show how an agent should use Microsoft Excel BI Agent on real Excel BI work without shipping customer-specific workbooks or reports inside the plugin.

The release gate validates this document with `tools/validate_task_recipes.py`. Keep package paths real, mention every canonical skill, and keep the required recipe headings intact so another agent can trust this as an executable routing guide rather than stale prose.

Installation and prompt/agent mirror sync are intentionally not repeated here. Use `docs/install-and-sync.md` as the single install contract.

## Routing First

Start with `excel-bi-router` when the task spans more than one Excel layer.

## Fast Profiles

Use `tools/run_task_profile.py` when the request matches a common workflow and the agent needs a repeatable command plan before running specialist scripts:

```powershell
python tools\run_task_profile.py `
  --profile audit `
  --workbook "workbook.xlsx" `
  --out-dir "$env:TEMP\excel_bi_audit" `
  --out-json "$env:TEMP\excel_bi_audit\profile.json" `
  --out-md "$env:TEMP\excel_bi_audit\profile.md"
```

Supported profiles: `audit`, `publish`, `pq-refresh`, `dax-review`, `cube-trace`, `env-diagnostics`, `report-build`, `fixture`, `case-regression`, `release-structural`, and `release-full`.

Expected evidence: the generated plan lists commands, output paths, and runtime boundaries. Use `--execute` only when the command plan points at copied workbooks or non-destructive reports.

For a machine-readable first pass, route the user request before choosing a specialized skill:

```powershell
python .agents\skills\excel-bi-router\scripts\route_excel_bi_task.py `
  --text "Workbook has VBA buttons, Power Query refresh, Data Model measures, and CUBEVALUE formulas." `
  --out-json "tmp\excel-bi-route.json" `
  --out-md "tmp\excel-bi-route.md"
```

Expected evidence: the route report names the layer, selected skill, matched signals, validation boundary, and recommended package scripts. It is a routing aid only; it does not inspect or validate the workbook.

```text
Layer: Mixed
Skill: excel-bi-router
Why: workbook contains formulas, Power Query, Data Model measures, and report formulas
Validation needed: Windows desktop Excel for refresh/evaluation, OpenXML for structural scan
```

Then move to the narrow skill for the layer being edited.

When a fresh agent needs a machine-readable map of this plugin before choosing scripts, build the capability catalog:

```powershell
python tools\build_capability_catalog.py `
  --project-root . `
  --out-json "tmp\capability-catalog.json" `
  --out-md "tmp\capability-catalog.md" `
  --require-pass
```

Expected evidence: the catalog lists all 12 canonical skills, package scripts from both `tools/` and `.agents/skills/*/scripts`, official documentation index counts, release-gate check inventory, core workflows, validation commands, and boundary text. It is discovery metadata, not proof that a workbook was validated.

When a fresh agent needs a single handoff folder before doing any workbook-specific work, build the agent bootstrap bundle:

```powershell
python tools\build_agent_bootstrap_bundle.py `
  --project-root . `
  --out-dir "$env:TEMP\excel_bi_agent_bootstrap" `
  --clean `
  --zip `
  --require-pass
```

Expected evidence: the bundle contains `BOOTSTRAP.md`, `bootstrap-manifest.json`, `capability-catalog.json/md`, `release-evidence.json/md`, `task-recipes.md`, `validation-commands.md`, and optional `agent-bootstrap-bundle.zip`. It is onboarding infrastructure for a new agent; it is not proof of external-agent behavior or workbook-specific validation.

## Recipe 1: Triage A Mixed Excel BI Workbook

Use when the workbook may contain VBA, Power Query, Power Pivot, PivotTables, and CUBE formulas.

On a new machine or before debugging provider-specific failures, capture the environment first:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\probe_excel_bi_providers.ps1 `
  -RunExcelComSmoke `
  -RunAdoWorkbookSmoke `
  -OutJson "tmp\provider-probe.json"
```

The full release gate runs the same provider baseline in Windows `full` profile and checks provider rows, COM activation rows, Excel COM smoke, generated-workbook ACE/ADO smoke, and interpretation text. Structural profile skips it so Linux/macOS package validation remains usable.

To turn the raw probe into a reviewer-facing capability and drift report, run:

```powershell
python tools\build_provider_environment_report.py `
  --project-root . `
  --probe-json "tmp\provider-probe.json" `
  --excel-com `
  --ado-workbook-smoke `
  --out-json "$env:TEMP\excel_bi_provider_environment.json" `
  --out-md "$env:TEMP\excel_bi_provider_environment.md" `
  --require-pass
```

Use `--baseline-json` to compare against a previous environment report. Use `--run-probe --excel-com --ado-workbook-smoke` when the report should create fresh local evidence directly. Keep generated provider reports outside the plugin package because they include machine-specific Office, COM, and provider details.

To validate baseline-comparison behavior without depending on the current machine, generate the synthetic provider fixture:

```powershell
python tools\create_provider_environment_fixture.py `
  --out-dir "$env:TEMP\excel_bi_provider_environment_fixture" `
  --out-json "$env:TEMP\excel_bi_provider_environment_fixture.json"

python tools\build_provider_environment_report.py `
  --project-root . `
  --probe-json "$env:TEMP\excel_bi_provider_environment_fixture\provider_fixture_probe.json" `
  --baseline-json "$env:TEMP\excel_bi_provider_environment_fixture\provider_matching_baseline.json" `
  --fail-on-drift `
  --out-json "$env:TEMP\provider_matching_report.json" `
  --require-pass
```

Use the generated `provider_drifting_baseline.json` when testing that `--fail-on-drift` blocks changed provider capability evidence. This fixture validates report logic only; it does not prove local Office/provider availability.

```powershell
python tools\inspect_excel_bi_workbook.py "workbook.xlsx" --markdown --out-json "tmp\openxml.json"
```

Before reviewing output sheets or creating a static handoff, build a formula quality report from the same structural scan:

```powershell
python tools\build_formula_quality_report.py `
  --openxml-json "tmp\openxml.json" `
  --out-json "tmp\formula-quality.json" `
  --out-md "tmp\formula-quality.md" `
  --fail-on-high-risk
```

Use this as an early warning for cached formula errors, `#REF!`, local path formulas, volatile functions, and dynamic reference functions. A clean static report still does not prove numeric correctness; it only means the selected static risks were not found.

Before handoff, also review workbook controls and visibility from the structural scan:

```powershell
python tools\build_workbook_controls_report.py `
  --openxml-json "tmp\openxml.json" `
  --out-json "tmp\workbook-controls.json" `
  --out-md "tmp\workbook-controls.md"
```

Use this to check hidden or very hidden sheets, workbook/sheet protection, filters, frozen panes, and data validation rules. Findings may be intentional, but they should be known before the workbook reaches a reviewer or client.

Before a pure-value deliverable or link-cleanup task, convert the structural scan into a readiness report:

```powershell
python tools\build_external_dependency_report.py `
  --openxml-json "tmp\openxml.json" `
  --out-json "tmp\external-dependencies.json" `
  --out-md "tmp\external-dependencies.md"
```

This report flags workbook connections, external links, external formulas, external defined names, mashup/model markers, CUBE formulas, and credential-like key names in connection metadata. Credential-like evidence is redacted to connection names, source kind, and indicator names; it does not print connection strings and does not replace a dedicated secret scanner.

When the workbook has several surfaces and you need one entry-point report for reviewer handoff, aggregate the structural scan plus any specialized reports already produced:

```powershell
python tools\build_workbook_triage_report.py `
  --inspection-json "tmp\openxml.json" `
  --formula-report-json "tmp\formula-quality.json" `
  --controls-report-json "tmp\workbook-controls.json" `
  --external-report-json "tmp\external-dependencies.json" `
  --out-json "tmp\workbook-triage.json" `
  --out-md "tmp\workbook-triage.md"
```

Use the triage report as the first delivery-readiness page: it summarizes workbook surfaces, supplied report status, missing coverage such as Power Query lineage or Data Model inspection, and the next commands to run. It is an aggregator, not a replacement for formula, external-dependency, CUBE, Power Query, or live Excel runtime validation.

Then build a non-destructive cleanup plan before editing the workbook copy:

```powershell
python tools\build_pure_deliverable_cleanup_plan.py `
  --readiness-json "tmp\external-dependencies.json" `
  --target pure-xlsx `
  --out-json "tmp\pure-cleanup-plan.json" `
  --out-md "tmp\pure-cleanup-plan.md"
```

After cleaning a copied workbook, re-run the structural scan/readiness report on the cleaned deliverable, then verify the cleanup plan assertions:

```powershell
python tools\build_pure_deliverable_verification_report.py `
  --cleanup-plan-json "tmp\pure-cleanup-plan.json" `
  --post-readiness-json "tmp\post-clean-external-dependencies.json" `
  --out-json "tmp\pure-cleanup-verification.json" `
  --out-md "tmp\pure-cleanup-verification.md" `
  --fail-on-fail
```

If Windows desktop Excel is available and Data Model metadata is needed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\inspect_excel_data_model_com.ps1 `
  -WorkbookPath "workbook.xlsx" `
  -IncludeColumns `
  -OutJson "tmp\model.json"
```

Expected evidence:

- workbook file type and macro presence
- Excel COM, ACE, MSOLAP, ADOMD, and bitness evidence when provider probing was run
- connection and table inventory
- external dependency readiness, including connections, credential-like connection string indicators, external links, external formulas, external defined names, CUBE formulas, mashup-like parts, and Data Model-like parts
- pure-deliverable cleanup plan with ordered copy, refresh, value-freeze, cleanup, and post-clean audit actions when a static deliverable is requested
- pure-deliverable verification report proving post-clean counts/markers are zero or explicitly documented for manual review
- CUBE formula count and referenced measures
- Data Model table, relationship, and measure inventory when Excel COM is available

Do not conclude that DAX or CUBE values are correct from OpenXML alone. OpenXML proves structure, not live Excel calculation.

## Recipe 2: Edit And Validate VBA In An `.xlsm`

Use `excel-vba-workbook-engineering`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .agents\skills\excel-vba-workbook-engineering\scripts\export_vba.ps1 `
  -WorkbookPath "source.xlsm" `
  -OutDir "src\vba"
```

Edit source files under `src/vba`, then import into a copy:

```powershell
python .agents\skills\excel-vba-workbook-engineering\scripts\lint_vba_source.py `
  "src\vba" `
  --strict-option-explicit `
  --out-json "tmp\vba-lint.json"
```

The lint step is a source-level guard only. It does not replace Excel VBA compile or macro execution validation.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .agents\skills\excel-vba-workbook-engineering\scripts\import_vba.ps1 `
  -WorkbookPath "source.xlsm" `
  -SourceDir "src\vba" `
  -OutputWorkbookPath "deliverables\source_macro_update.xlsm"
```

If the workbook has shape buttons or form-control buttons, inspect the output workbook and cross-check `OnAction` values against exported public VBA macros:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .agents\skills\excel-vba-workbook-engineering\scripts\inspect_workbook.ps1 `
  -WorkbookPath "deliverables\source_macro_update.xlsm" `
  -OutJson "tmp\workbook-inventory.json"
python tools\build_vba_button_binding_report.py `
  --workbook-inventory-json "tmp\workbook-inventory.json" `
  --vba-lint-json "tmp\vba-lint.json" `
  --out-json "tmp\vba-button-bindings.json" `
  --out-md "tmp\vba-button-bindings.md" `
  --fail-on-unresolved
```

This catches stale button assignments such as renamed macros, deleted macros, or workbook/module-prefixed `OnAction` values that no longer resolve to a public entry macro.

Expected evidence:

- source workbook was not overwritten unless explicitly requested
- VBA exported before modification
- edited VBA source passed source lint before import when source files were changed
- button/form-control `OnAction` bindings were checked when the workbook uses button-driven macros
- package maintenance gate includes a generic live import/run/export smoke test for this VBA path
- output remains `.xlsm`
- representative macro or compile check was run in Excel
- no orphaned hidden Excel process remains

## Recipe 3: Add, Update, Delete, Or Refresh Power Query

Use `power-query-m-engineering`.

List queries:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .agents\skills\power-query-m-engineering\scripts\manage_power_queries_excel_com.ps1 `
  -WorkbookPath "workbook.xlsx" `
  -Action List `
  -OutJson "tmp\queries.before.json"
```

Before importing changed M source, run static lint:

```powershell
python .agents\skills\power-query-m-engineering\scripts\lint_power_query_m.py `
  "src\m" `
  --out-json "tmp\m-lint.json"
```

Use `--warnings-as-errors` when folder-ingestion filters, join-cardinality risk, order-restoration gaps, hard-coded expand columns, or unguarded `List.Max` should block a delivery build.

After exporting workbook queries, build a lineage/source-risk report before large rewrites or delivery review:

```powershell
python tools\build_power_query_lineage_report.py `
  "src\m" `
  --out-json "tmp\pq-lineage.json" `
  --out-md "tmp\pq-lineage.md" `
  --fail-on-high-risk
```

This static report maps exported query dependencies and flags query cycles, hard-coded local paths, web/database/cloud-service endpoints, `Value.NativeQuery` native SQL pass-through, credential-like literals or authorization keys, and mixed-source lineage that can cause privacy-firewall or credential issues. It recognizes common enterprise connectors such as OData, Azure Storage, Power Platform Dataflows, Dataverse, additional database connectors, SharePoint, and Web sources. Credential-like evidence is redacted to indicators and counts. It does not refresh sources, prove credentials are available, or replace a dedicated secret scanner.

Preferred source pattern: keep environment-specific file/folder paths in a workbook table or named configuration range, then read them with `Excel.CurrentWorkbook`. The lineage report records that as `workbook-config` plus the actual external source kind, and does not treat `workbook-config + one external source` as a mixed-source privacy risk by itself.

Update a query into a copy:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .agents\skills\power-query-m-engineering\scripts\manage_power_queries_excel_com.ps1 `
  -WorkbookPath "workbook.xlsx" `
  -Action Update `
  -QueryName "Query1" `
  -FormulaPath "src\m\Query1.m" `
  -OutputWorkbookPath "deliverables\workbook_pq_update.xlsx" `
  -OutJson "tmp\queries.update.json"
```

Refresh and wait for completion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .agents\skills\power-query-m-engineering\scripts\refresh_power_queries_excel_com.ps1 `
  -WorkbookPath "deliverables\workbook_pq_update.xlsx" `
  -OutputWorkbookPath "deliverables\workbook_pq_refreshed.xlsx" `
  -DisableBackgroundRefresh `
  -OutJson "tmp\refresh.json"
```

When a query is loaded to a worksheet table, validate the refreshed output workbook's loaded table values or row counts. A named `WorkbookQuery.Refresh` is not enough evidence by itself after the M formula changes; the refresh helper refreshes matching `ListObject.QueryTable` load targets and saves the refreshed copy when `-OutputWorkbookPath` is supplied.

Build a refresh timing/status report before downstream VBA, Power Pivot, or delivery steps continue:

```powershell
python .agents\skills\power-query-m-engineering\scripts\build_power_query_refresh_report.py `
  "tmp\refresh.json" `
  --require-completed `
  --max-elapsed-seconds 300 `
  --out-json "tmp\refresh-status.json" `
  --out-md "tmp\refresh-status.md" `
  --fail-on-error
```

Use `--fail-on-warning` only when slow refresh, missing elapsed time, or background-refresh settings should block the delivery build. The report is evidence summarization only; it does not refresh Excel by itself.

If refresh fails or captures errors, classify the report before rewriting the query:

```powershell
python .agents\skills\power-query-m-engineering\scripts\classify_power_query_refresh_errors.py `
  "tmp\refresh.json" `
  --out-json "tmp\refresh-diagnosis.json" `
  --out-md "tmp\refresh-diagnosis.md"
```

Expected evidence:

- query inventory before/after
- changed M source passed static lint before import
- query lineage/source risks reviewed when the workbook has multiple queries or external sources
- refresh completed without captured errors
- refresh timing/status report reviewed before dependent automation continued
- refresh failures were classified when refresh did not complete
- row counts and schema checked for the changed output
- join/group/expand logic validated for cardinality and order

## Recipe 4: Inspect Power Pivot And DAX

Use `power-pivot-dax-modeling`.

Before adding or revising Excel-targeted DAX source, run the compatibility lint when formulas are available as `.dax` or JSON:

```powershell
python .agents\skills\power-pivot-dax-modeling\scripts\lint_dax_compat.py `
  "src\dax" `
  --profile excel `
  --warn-division `
  --out-json "tmp\dax-lint.json"
```

This catches Power BI-style functions that are risky in Excel Power Pivot, such as `REMOVEFILTERS`, and can also surface review warnings for version-sensitive functions such as `SELECTEDVALUE` and raw `/` ratio division when `--warn-division` is enabled. Use `--warnings-as-errors` only when warnings should block the handoff. This does not prove the model evaluates correctly.

When model JSON or DAX source includes multiple measures, run dependency analysis before renaming, deleting, or handing formulas back to Excel:

```powershell
python .agents\skills\power-pivot-dax-modeling\scripts\analyze_dax_dependencies.py `
  "tmp\model.json" `
  --out-json "tmp\dax-dependencies.json" `
  --out-md "tmp\dax-dependencies.md"
```

This catches missing measure references, direct self references, dependency cycles, and duplicate measure names. It does not validate table or column names and does not evaluate DAX.

Before renaming or deleting a measure, check report-layer impact when OpenXML metadata is available:

```powershell
python tools\analyze_measure_rename_impact.py `
  --model-json "tmp\model.json" `
  --openxml-json "tmp\openxml.json" `
  --rename "Old Measure=New Measure" `
  --out-json "tmp\measure-rename-impact.json" `
  --out-md "tmp\measure-rename-impact.md"
```

This identifies dependent DAX measures and CUBE formulas that still reference the old `[Measures].[...]` path.

When the impact report shows formulas that must change, build a reviewable replacement plan before editing the workbook:

```powershell
python tools\build_measure_rename_rewrite_plan.py `
  --model-json "tmp\model.json" `
  --openxml-json "tmp\openxml.json" `
  --rename "Old Measure=New Measure" `
  --out-json "tmp\measure-rename-rewrite-plan.json" `
  --out-md "tmp\measure-rename-rewrite-plan.md"
```

This proposes static DAX and CUBE formula replacements without changing the workbook. Deletions and ambiguous cases are kept as manual-review items.

For measure deletion, do not expect a replacement formula. Treat direct DAX/CUBE references and downstream formulas that depend on a deleted-measure helper cell as blockers:

```powershell
python tools\build_measure_rename_rewrite_plan.py `
  --model-json "tmp\model.json" `
  --openxml-json "tmp\openxml.json" `
  --delete "Old Measure" `
  --out-json "tmp\measure-delete-review-plan.json" `
  --out-md "tmp\measure-delete-review-plan.md" `
  --fail-on-manual-review
```

Use the generated manual-review rows to either remove the report surface, remap it to a new approved measure, or change the workbook design before applying the model deletion.

For helper-cell layouts, the plan distinguishes two cases:

- formulas that should be rewritten, such as `CUBEMEMBER(...,"[Measures].[Old]")`
- downstream formulas that do not need text replacement but depend on a rewritten helper cell, such as `CUBEVALUE(...,$H$2,$A$2)`

Dynamic period/member helper references such as `$A$5` are retained in the plan so reviewers can see which business selector remains part of the formula.

MDX bracket identifiers can contain a literal closing bracket escaped as `]]`, for example `[Measures].[Revenue ]] Special]`. The CUBE dependency report, model report, rename impact analyzer, and rewrite planner share the same parser for this pattern, so a model measure named `Revenue ] Special` is matched consistently across report extraction and rewrite planning.

When validating the report-building pipeline without a customer workbook, generate a structural fixture first:

```powershell
python tools\create_cube_formula_fixture.py `
  --workbook "tmp\cube_formula_fixture.xlsx" `
  --model-json "tmp\cube_model_summary.json"

python tools\inspect_excel_bi_workbook.py "tmp\cube_formula_fixture.xlsx" --out-json "tmp\openxml.json"

python tools\build_excel_bi_model_report.py `
  --model-json "tmp\cube_model_summary.json" `
  --openxml-json "tmp\openxml.json" `
  --out-md "tmp\model-report.md" `
  --out-json "tmp\model-report.json"
```

This validates the metadata/report pipeline and CUBE-reference mapping. It does not validate a live Excel Data Model.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\inspect_excel_data_model_com.ps1 `
  -WorkbookPath "workbook.xlsx" `
  -IncludeColumns `
  -OutJson "tmp\model.json"

python tools\inspect_excel_bi_workbook.py "workbook.xlsx" --out-json "tmp\openxml.json"

python tools\build_excel_bi_model_report.py `
  --model-json "tmp\model.json" `
  --openxml-json "tmp\openxml.json" `
  --out-md "tmp\model-report.md" `
  --out-json "tmp\model-report.json"
```

Expected evidence:

- model tables, relationships, measures, and connection metadata
- measure formulas captured from Excel COM where available
- DAX source lint passed for Excel-targeted measures when source formulas were changed
- DAX measure dependencies were checked when multiple measures were changed or renamed
- measure rename/delete impact was checked when report-layer CUBE formulas are present
- reviewable measure rewrite plans were generated before applying report-layer formula replacements
- helper-cell downstream impacts and dynamic period/member helper references were reviewed before applying formula changes
- report-layer CUBE formulas mapped to model measures
- for the generic fixture, 3 tables, 2 relationships, 2 measures, 7 CUBE formulas, and one intentionally missing CUBE measure reference
- DAX checks tested at expected grains, not only grand totals

## Recipe 5: Trace CUBE Formulas And MDX References

Use `mdx-cubevalue-extraction`.

When no safe workbook is available and the task is to validate the parser/report pipeline itself, create a generic structural fixture:

```powershell
python tools\create_cube_formula_fixture.py `
  --workbook "tmp\cube_formula_fixture.xlsx" `
  --model-json "tmp\cube_model_summary.json"
```

This fixture is not a live Power Pivot model. It is for validating OpenXML CUBE formula discovery, measure reference extraction, helper-cell tracing, and diagnostic flags.

```powershell
python tools\inspect_excel_bi_workbook.py "workbook.xlsx" --out-json "tmp\openxml.json"

python tools\build_cube_dependency_report.py `
  --openxml-json "tmp\openxml.json" `
  --model-json "tmp\model.json" `
  --out-md "tmp\cube-report.md" `
  --out-json "tmp\cube-report.json" `
  --out-mermaid "tmp\cube-report.mmd"
```

Expected evidence:

- formula cell addresses by sheet
- referenced measures and MDX member paths
- helper-cell dependencies
- missing model measure references flagged before formula rewrites
- measure rename/delete impact checked before changing model measure names used by CUBE formulas
- measure rewrite plan generated before applying static `[Measures].[...]` replacements
- downstream `CUBEVALUE` formulas flagged when they depend on rewritten measure helper cells
- for the generic fixture, exactly 7 CUBE formulas and expected diagnostics for missing measure, hard-coded period marker, and dynamic MDX string

The full release gate runs `tools\test_excel_adomd_query.ps1 -ProbeOnly` to verify local ADODB/ADOMD COM activation where Windows runtime support exists. When a real ADOMD/MSOLAP endpoint connection string exists, validate the endpoint MDX separately:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\test_excel_adomd_query.ps1 `
  -ProbeOnly `
  -OutJson "tmp\adomd-probe.json"
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\test_excel_adomd_query.ps1 `
  -ConnectionString "Provider=MSOLAP;Data Source=server;Initial Catalog=model;" `
  -Mdx "SELECT [Measures].[Sales] ON 0 FROM [Model]" `
  -MaxCells 100 `
  -OutJson "tmp\adomd-query.json"
```

Probe-only validation confirms local COM activation only. Endpoint validation confirms one cube/model endpoint and MDX query. Neither path directly recalculates Excel `CUBEVALUE` formulas.

## Recipe 6: Validate ADO/OLEDB Workbook SQL

Use `excel-ado-sql-data-access`.

Probe provider availability:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\probe_excel_bi_providers.ps1 `
  -RunAdoWorkbookSmoke `
  -OutJson "tmp\provider-probe.json"
```

Create and query a generic fixture:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\test_excel_ado_sql_access.ps1 `
  -WorkbookPath "tmp\ado fixture.xlsx" `
  -CreateFixture `
  -SqlText "SELECT Category, SUM(Amount) AS TotalAmount FROM [Data$] GROUP BY Category ORDER BY Category" `
  -IncludeSchema `
  -OutJson "tmp\ado-smoke.json"
```

Expected evidence:

- ACE OLEDB provider can open the workbook
- SQL returns expected fields and rows
- release-gate fixture query `SELECT * FROM [Data$]` returns 4 rows, fields `Region`, `Category`, `Amount`, `Period`, schema table `Data$`, and Amount total `500`
- file paths with spaces and non-ASCII characters are tested when relevant
- recordsets/connections are closed
- structural profile skips this runtime check because it requires Windows Excel COM plus a compatible ACE OLEDB provider

## Recipe 7: Git Bash On Windows

Use `tools/invoke_excel_bi_com.sh` when the terminal is Git Bash/MSYS/Cygwin and the task needs Windows Excel COM.

```bash
tools/invoke_excel_bi_com.sh provider-probe --excel-com --ado-workbook-smoke --out-json "tmp/provider-probe.json"
tools/invoke_excel_bi_com.sh model-report -w "workbook.xlsx" --include-columns --out-md "tmp/model-report.md"
tools/invoke_excel_bi_com.sh cube-report -w "workbook.xlsx" --include-model --out-md "tmp/cube-report.md"
tools/invoke_excel_bi_com.sh ado-query -w "tmp/ado fixture.xlsx" --create-fixture --sql "SELECT * FROM [Data$]"
tools/invoke_excel_bi_com.sh adomd-query --probe-only --out-json "tmp/adomd-probe.json"
```

Expected evidence is the same as the PowerShell path. The wrapper only handles shell/path translation; it does not reduce the need for Excel COM validation.

## Recipe 8: Linux Or macOS Structural Review

Use this only when desktop Excel is not available.

```bash
python tools/inspect_excel_bi_workbook.py workbook.xlsx --markdown --out-json tmp/openxml.json
tools/excel_bi_structural.sh sanitized-bundle --out-dir /tmp/excel_bi_sanitized_fixtures --clean --validate
tools/excel_bi_structural.sh provider-baseline-fixture --out-dir /tmp/excel_bi_provider_baseline_fixture --clean
```

Safe claims:

- workbook package structure
- macro binary presence
- formulas and CUBE formula strings visible in OpenXML
- workbook connections, pivot caches, and table metadata visible in package XML
- provider baseline comparison logic, using synthetic matching/drifting provider reports without live Office provider dependency

Unsafe claims without Excel:

- VBA compiles or buttons work
- Power Query refresh succeeds
- Data Model measures evaluate correctly
- CUBE formulas return the expected values
- Solver or Excel-only automation behaves correctly
- local Office provider availability; the provider baseline fixture proves drift-report logic only

## Recipe 9: Publish A Clean Excel Deliverable

Use `excel-deliverable-publisher` when the requested output is a client-ready copy rather than an editable working model.

```powershell
python tools\inspect_excel_bi_workbook.py "workbook.xlsx" --out-json "tmp\openxml.json"
python tools\build_external_dependency_report.py `
  --openxml-json "tmp\openxml.json" `
  --out-json "tmp\external-dependencies.json" `
  --out-md "tmp\external-dependencies.md"
python tools\build_pure_deliverable_cleanup_plan.py `
  --readiness-json "tmp\external-dependencies.json" `
  --target pure-xlsx `
  --out-json "tmp\pure-cleanup-plan.json" `
  --out-md "tmp\pure-cleanup-plan.md"
```

Expected evidence: the source workbook is not overwritten, required refresh/calculation happens before value-freezing, and the post-clean copy is verified with `build_pure_deliverable_verification_report.py`.

## Recipe 10: Audit Workbook QA Before Delivery

Use `excel-workbook-qa-auditor` for review-first requests.

```powershell
python tools\inspect_excel_bi_workbook.py "workbook.xlsx" --out-json "tmp\openxml.json"
python tools\build_formula_quality_report.py --openxml-json "tmp\openxml.json" --out-json "tmp\formula-quality.json"
python tools\build_workbook_controls_report.py --openxml-json "tmp\openxml.json" --out-json "tmp\workbook-controls.json"
python tools\build_workbook_triage_report.py `
  --inspection-json "tmp\openxml.json" `
  --formula-report-json "tmp\formula-quality.json" `
  --controls-report-json "tmp\workbook-controls.json" `
  --out-json "tmp\workbook-triage.json" `
  --out-md "tmp\workbook-triage.md"
```

Expected evidence: findings are prioritized by delivery risk, and static QA is not presented as proof of live numeric correctness.

## Recipe 11: Diagnose Office Environment Readiness

Use `office-environment-diagnostics` when errors may come from local Office, COM, provider, or shell setup.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\probe_excel_bi_providers.ps1 `
  -RunExcelComSmoke `
  -RunAdoWorkbookSmoke `
  -OutJson "tmp\provider-probe.json"
python tools\build_provider_environment_report.py `
  --project-root . `
  --probe-json "tmp\provider-probe.json" `
  --excel-com `
  --ado-workbook-smoke `
  --out-json "tmp\provider-environment.json" `
  --out-md "tmp\provider-environment.md"
```

Expected evidence: Office/provider readiness is separated from workbook correctness, and Linux/macOS work is limited to structural checks unless Excel runtime is available.

## Recipe 12: Build A Polished Excel Report Workbook

Use `excel-report-builder` when the primary work is a readable workbook or dashboard surface.

```powershell
python tools\inspect_excel_bi_workbook.py "report.xlsx" --out-json "tmp\report-openxml.json"
python tools\build_formula_quality_report.py --openxml-json "tmp\report-openxml.json" --out-json "tmp\report-formula-quality.json"
python tools\build_workbook_controls_report.py --openxml-json "tmp\report-openxml.json" --out-json "tmp\report-controls.json"
python tools\build_visual_qa_report.py --workbook "report.xlsx" --out-json "tmp\report-visual-qa.json" --out-md "tmp\report-visual-qa.md"
```

When rendered Visual QA evidence chain V1 is required on Windows desktop Excel, export task-local PDF evidence:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\export_visual_qa_render_evidence.ps1 `
  -WorkbookPath "report.xlsx" `
  -OutDir "$env:TEMP\excel_bi_visual_render" `
  -OutJson "$env:TEMP\excel_bi_visual_render.json" `
  -OutMd "$env:TEMP\excel_bi_visual_render.md"
```

For sanitized smoke tests only, use `-CreateFixture` so the tool creates a customer-data-free Excel-COM-native workbook before exporting:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\export_visual_qa_render_evidence.ps1 `
  -WorkbookPath "$env:TEMP\excel_bi_visual_render\visual_render_fixture.xlsx" `
  -CreateFixture `
  -OutDir "$env:TEMP\excel_bi_visual_render\pdf" `
  -OutJson "$env:TEMP\excel_bi_visual_render\render.json" `
  -OutMd "$env:TEMP\excel_bi_visual_render\render.md"
```

Expected evidence: inputs, calculations, outputs, and QA areas are separated; visual QA records visible report-sheet risks such as blank report sheets, missing print areas, and long text in narrow cells; rendered evidence records `Windows Excel COM PDF export`, `readiness=rendered`, exported `Report*` sheet count, PDF byte sizes, and boundaries; Power Query, Data Model, and VBA logic changes are routed to their specialist skills before final report publishing.

## Recipe 13: Review Power BI Semantic Model Portability

Use `power-bi-semantic-model` when a request may target Power BI instead of Excel Power Pivot.

```powershell
python tools\search_official_docs.py --project-root . --query "Power BI semantic model XMLA TMDL"
python .agents\skills\power-pivot-dax-modeling\scripts\lint_dax_compat.py `
  "src\dax" `
  --profile excel `
  --warn-division `
  --out-json "tmp\dax-excel-compat.json"
```

Expected evidence: the response distinguishes Power BI semantic models from Excel `ThisWorkbookDataModel`, flags portability risks, and uses official Microsoft documentation when current Power BI behavior matters.

## Recipe 14: Create Sanitized Testing Fixtures

Use `excel-testing-fixtures` when a workflow needs reproducible evidence without customer files.

```powershell
python tools\build_sanitized_fixture_bundle.py `
  --out-dir "$env:TEMP\excel_bi_sanitized_fixtures" `
  --clean
python tools\build_cross_agent_forward_test_pack.py `
  --project-root . `
  --out-dir "$env:TEMP\excel_bi_forward_test_pack" `
  --clean `
  --require-pass
```

Expected evidence: fixtures prove the designed parser/report/scorer path only, not every customer workbook shape. Keep generated customer-workbook reports outside the plugin package.

## Recipe 15: Run Real/Sanitized Case Regression

Use `excel-testing-fixtures` when a future agent needs to confirm that the package still tracks recurring Excel BI problem shapes without storing customer workbooks.

Preferred profile entry point:

```powershell
python tools\run_task_profile.py `
  --profile case-regression `
  --out-dir "$env:TEMP\excel_bi_case_regression" `
  --out-json "$env:TEMP\excel_bi_case_regression\profile.json" `
  --out-md "$env:TEMP\excel_bi_case_regression\profile.md" `
  --execute
```

Direct runner:

```powershell
python tools\run_case_regression.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_case_regression.json" `
  --out-md "$env:TEMP\excel_bi_case_regression.md" `
  --require-pass
```

Expected evidence:

- the manifest under `fixtures/real-sanitized-cases/manifest.json` is readable.
- seven seed cases cover Power Query, DAX, CUBE/MDX, VBA, clean deliverable publishing, visual QA, and environment/capability routing.
- every case has required fields, evidence mode, package-tool references, expected evidence, boundaries, and a next live-workbook requirement.
- no forbidden local/customer markers are present.
- the visual QA case is fixture-backed through `tools/create_visual_qa_fixture.py` and `tools/build_visual_qa_report.py`.
- this validates case definitions and package-tool references only; it does not prove any private workbook is correct.

## Recipe 16: Establish Excel Compatibility Before Implementation

Use `office-environment-diagnostics` when the question is whether an operation can run on a platform, host, Office version, bitness, offline machine, or recipient environment. DAX/Power Pivot formula and function compatibility remains with `power-pivot-dax-modeling` unless the request is explicitly about platform/host runtime availability.

First record the authoring target, automation target, consumer target, and recipient target. Do not assume the agent's current machine represents any of them.

Generate a Windows capability probe and report:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\probe_excel_capabilities.ps1 `
  -OutJson "$env:TEMP\excel-capabilities.json" `
  -Profile runtime
python tools\build_excel_compatibility_report.py `
  --probe-json "$env:TEMP\excel-capabilities.json" `
  --out-json "$env:TEMP\excel-compatibility.json" `
  --out-md "$env:TEMP\excel-compatibility.md" `
  --require-capability excel.com.activation `
  --require-pass
```

When the agent is not on the target Windows machine, use a captured probe without running local COM:

```powershell
python tools\run_task_profile.py `
  --profile env-diagnostics `
  --probe-json "C:\evidence\recipient-capabilities.json" `
  --require-capability excel.com.activation `
  --out-dir "$env:TEMP\excel_env_plan" `
  --execute
```

Expected evidence:

- structural evidence identifies package, source, formula, and workbook-shape compatibility with `low` confidence;
- runtime capability evidence records the exact Windows/Office build, 32-bit or 64-bit process/provider context, policy, and operation readiness with `medium` confidence;
- workbook behavior evidence exercises the representative workbook in the intended host and supports `high` confidence only for that recorded target;
- macOS, Excel for web, Linux, offline, WPS, LibreOffice, Excel 2007/2010/2013/2016/2019, Office LTSC, and Microsoft 365 targets retain explicit unknown/blocked boundaries until target-specific evidence exists.

## Delivery Rule

Generated reports from customer workbooks should live in task-specific temporary or deliverable folders, not inside this plugin package. Generic smoke fixtures should be created in task-local or release-local output folders unless they are deliberately sanitized and documented for public fixtures.

## Sanitized Fixture Bundle

When no customer workbook can be shared, generate a safe bundle of structural examples for parser, report, and cleanup-plan validation:

```powershell
python tools\build_sanitized_fixture_bundle.py `
  --out-dir "$env:TEMP\excel_bi_sanitized_fixtures" `
  --clean
```

From Git Bash, Linux, or macOS, use the portable structural wrapper when the bundle should be generated and immediately validated without Excel COM:

```bash
tools/excel_bi_structural.sh sanitized-bundle \
  --out-dir /tmp/excel_bi_sanitized_fixtures \
  --clean \
  --validate
```

The generated bundle contains CUBE-formula, external-dependency, and pure-deliverable `.xlsx` fixtures, plus safe/risky exported Power Query M source folders with metadata and a README. The external-dependency fixture includes two safe structural workbook connections, including one redacted credential-like connection indicator, so parser/report workflows can validate both ordinary connection detection and credential-indicator reporting without customer files. The Power Query lineage fixture validates clean parameterized M sources and risky M sources covering local paths, web/database/cloud-service sources, native SQL, credential-like literals, mixed-source lineage, and dependency cycles. Use the bundle to validate package behavior without moving client files into the plugin source tree.

For provider/environment drift reporting, use the portable fixture wrapper when a Linux/macOS/Git Bash agent needs to prove comparison behavior without touching live Office providers:

```bash
tools/excel_bi_structural.sh provider-baseline-fixture \
  --out-dir /tmp/excel_bi_provider_baseline_fixture \
  --clean
```

Expected evidence: matching baseline returns zero changes and passes; drifting baseline returns a non-zero status under `--fail-on-drift` and reports required changed paths for Excel COM smoke, workbook SQL smoke, MSOLAP, ADOMD COM, ADOMD.NET, and readiness flags. This validates baseline-comparison logic only; it does not prove the current machine has those providers installed.

## Cross-Agent Forward-Test Pack

When you need another agent to test this plugin without leaking customer workbooks or prior conclusions, generate the portable prompt pack:

```bash
python tools/build_cross_agent_forward_test_pack.py \
  --project-root . \
  --out-dir /tmp/excel_bi_forward_test_pack \
  --clean \
  --require-pass
```

Expected evidence:

- 48 generated prompt files.
- 4 target styles: Codex, Claude, OpenCode, and generic agents.
- all 12 canonical skills represented.
- each prompt includes expected evidence and runtime-boundary language.

This pack is an evaluation input, not an execution result. Use fresh agent sessions for real forward-testing, then review their outputs against the prompt's expected-evidence section.

For the normal fresh-session handoff, generate the complete bundle in one command:

```bash
python tools/build_cross_agent_forward_test_handoff_bundle.py \
  --project-root . \
  --out-dir /tmp/excel_bi_forward_test_handoff \
  --clean \
  --zip-path /tmp/excel_bi_forward_test_handoff.zip \
  --require-pass
```

Expected handoff evidence:

- a `pack/` prompt set with 48 prompts.
- a `runbook/` folder with assignment matrix and scoring command.
- `responses/<agent>/<skill>.md` stubs for all 48 expected responses.
- `HANDOFF.md` and `handoff-manifest.json` with boundary language.
- an optional zip archive that contains the prompt pack, runbook, response stubs, handoff manifest, and handoff Markdown.

Generate a collection runbook before sending prompts out:

```bash
python tools/build_cross_agent_forward_test_runbook.py \
  --manifest-json /tmp/excel_bi_forward_test_pack/forward-test-pack.json \
  --responses-dir /tmp/excel_bi_forward_test_responses \
  --out-dir /tmp/excel_bi_forward_test_runbook \
  --clean \
  --write-response-stubs \
  --require-pass
```

Expected runbook evidence:

- 48 assignment rows.
- response stubs for `responses/<agent>/<skill>.md`.
- a scoring command for the collected responses.
- clear boundary text that the runbook is not external-agent proof.

Save real fresh-session responses under `responses/<agent>/<skill>.md`, then score them:

```bash
python tools/score_cross_agent_forward_test_results.py \
  --manifest-json /tmp/excel_bi_forward_test_pack/forward-test-pack.json \
  --responses-dir /tmp/excel_bi_forward_test_responses \
  --out-json /tmp/excel_bi_forward_test_score.json \
  --out-md /tmp/excel_bi_forward_test_score.md \
  --require-pass
```

Expected scoring evidence:

- 48 response files when every agent/skill prompt has been run.
- package-tool and runtime-boundary evidence in every response.
- incomplete or placeholder-like responses are blocked.
- generated sample responses are only scorer fixtures, not external-agent proof.

After scoring, build a collection report before claiming real external-agent behavior:

```bash
python tools/build_cross_agent_response_collection_report.py \
  --manifest-json /tmp/excel_bi_forward_test_pack/forward-test-pack.json \
  --responses-dir /tmp/excel_bi_forward_test_responses \
  --score-json /tmp/excel_bi_forward_test_score.json \
  --out-json /tmp/excel_bi_forward_test_collection.json \
  --out-md /tmp/excel_bi_forward_test_collection.md \
  --require-pass
```

Expected collection evidence:

- stubs are reported as `collecting`, not proof.
- generated scorer fixtures are reported as `fixture-only`, not proof.
- candidate fresh-session responses are counted separately.
- `--require-external-proof` only passes after all 48 responses are candidate fresh-session outputs and the scorer passes.

## Workbook Surface Fixture

When the task is to validate workbook delivery-surface inspection without a customer file, generate a safe structural workbook with normal formulas, workbook-defined names, a worksheet table, and a chart package part:

```powershell
python tools\create_workbook_surface_fixture.py `
  --workbook "$env:TEMP\workbook_surface_fixture.xlsx" `
  --out-json "$env:TEMP\workbook_surface_fixture.json"

python tools\inspect_excel_bi_workbook.py `
  "$env:TEMP\workbook_surface_fixture.xlsx" `
  --out-json "$env:TEMP\workbook_surface_openxml.json" `
  --markdown
```

Use this fixture to validate structural inspection logic only. It does not prove Excel formula results, chart rendering, VBA button behavior, or Power Query refresh.

## Formula Quality Fixture

When the task is to validate formula-quality reporting without a customer file, generate safe OpenXML-inspection JSON fixtures:

```powershell
python tools\create_formula_quality_fixture.py `
  --out-dir "$env:TEMP\formula_quality_fixture" `
  --out-json "$env:TEMP\formula_quality_fixture_manifest.json"

python tools\build_formula_quality_report.py `
  --openxml-json "$env:TEMP\formula_quality_fixture\formula_quality_safe_openxml.json" `
  --out-json "$env:TEMP\formula_quality_safe.json" `
  --fail-on-high-risk

python tools\build_formula_quality_report.py `
  --openxml-json "$env:TEMP\formula_quality_fixture\formula_quality_risky_openxml.json" `
  --out-json "$env:TEMP\formula_quality_risky.json" `
  --out-md "$env:TEMP\formula_quality_risky.md"
```

The risky fixture intentionally includes cached Excel error values, `#REF!`, a local path formula, `INDIRECT`, and `NOW`. It validates reporting logic only; it does not calculate workbook results.

## Workbook Controls Fixture

When the task is to validate workbook visibility and protection reporting without a customer file, generate a safe workbook-controls fixture:

```powershell
python tools\create_workbook_controls_fixture.py `
  --workbook "$env:TEMP\workbook_controls_fixture.xlsx" `
  --out-json "$env:TEMP\workbook_controls_fixture.json"

python tools\inspect_excel_bi_workbook.py `
  "$env:TEMP\workbook_controls_fixture.xlsx" `
  --out-json "$env:TEMP\workbook_controls_openxml.json"

python tools\build_workbook_controls_report.py `
  --openxml-json "$env:TEMP\workbook_controls_openxml.json" `
  --out-json "$env:TEMP\workbook_controls_report.json" `
  --out-md "$env:TEMP\workbook_controls_report.md"
```

The fixture intentionally includes one hidden sheet, one very hidden sheet, workbook structure protection, one protected sheet, one filtered sheet, one frozen pane, and one data validation rule. It validates static reporting only; it does not test Excel passwords or user interaction.

## Artifact Hygiene Audit

Before release or handoff, run the artifact hygiene report to confirm the plugin package does not contain customer workbooks, Excel lock files, generated machine reports, Python bytecode/cache folders, local screenshot paths, or customer-specific markers:

```powershell
python tools\build_artifact_hygiene_report.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_artifact_hygiene.json" `
  --out-md "$env:TEMP\excel_bi_artifact_hygiene.md" `
  --require-pass
```

The report fails on unexpected Office files, customer artifacts, generated reports, lock files, and local path markers. Write the JSON/Markdown report to a temp path, not into the plugin package.

## Goal Coverage Audit

Before claiming that the package still covers the active goal after a broad change, run the coverage report:

```powershell
python tools\build_goal_coverage_report.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_goal_coverage.json" `
  --out-md "$env:TEMP\excel_bi_goal_coverage.md" `
  --require-pass
```

This checks evidence presence across the core goal areas. It does not replace the full release gate; use it as an early audit before deployment.

## Completion Readiness Audit

Before considering whether the active thread goal can be closed, run the completion-readiness audit:

```powershell
python tools\build_completion_readiness_audit.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_completion_readiness.json" `
  --out-md "$env:TEMP\excel_bi_completion_readiness.md" `
  --require-pass
```

Expected current evidence: `status=pass`, `coverage.status=pass`, and either an in-progress blocker state while the public optimization backlog remains active or `completionReady=true` after every public backlog item is closed or explicitly accepted. Use `--require-complete` only for a deliberate final closure audit.

## Official Documentation Drift Report

Use the drift report when reviewing the bundled Microsoft documentation indexes or preparing a handoff where official-source coverage matters:

```powershell
python tools\build_official_docs_drift_report.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_official_docs_drift.json" `
  --out-md "$env:TEMP\excel_bi_official_docs_drift.md" `
  --require-pass
```

The default mode is offline and deterministic. It inventories every bundled `official-docs-index.json`, summarizes official Microsoft URLs by skill, host, and category, and can compare against a prior report with `--baseline-json`. For periodic live drift checks, add `--check-online --online-limit 8` to sample Microsoft official URLs; keep those reports in a temp folder because HTTP status and redirects are environment-specific.

## Release Evidence Bundle

Before handing the package to another agent or reviewer, build a consolidated evidence bundle:

```powershell
python tools\build_release_evidence_bundle.py `
  --project-root . `
  --release-gate-json "$env:TEMP\excel_bi_release_gate_final.json" `
  --out-json "$env:TEMP\excel_bi_release_evidence.json" `
  --out-md "$env:TEMP\excel_bi_release_evidence.md" `
  --require-pass
```

If no release gate report exists yet, omit `--release-gate-json`; the bundle will still summarize project-docs, task-recipe, official-docs, and goal-coverage evidence, while clearly marking the release gate attachment as not supplied.
