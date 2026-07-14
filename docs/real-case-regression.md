# Real/Sanitized Case Regression V1

## Goal

Create a reusable regression library for Excel BI work learned from real task patterns without storing customer workbooks.

V1 is complete when:

- `fixtures/real-sanitized-cases/manifest.json` defines the case library.
- Six seed case specs cover Power Query, DAX, CUBE/MDX, VBA, deliverable cleanup, and visual QA.
- The visual QA case is workbook-backed through a generated sanitized fixture.
- `tools/run_case_regression.py` validates schema, coverage, package-tool references, evidence modes, and safety boundaries.
- `tools/run_task_profile.py --profile case-regression` points agents to the runner.
- `docs/task-recipes.md`, `docs/current-status.md`, and the maintenance goal docs record the public entry point, validation boundary, and remaining optimization backlog.

## Boundaries

- No customer workbooks, screenshots, private paths, credentials, or business-specific rules are stored in this package.
- A case definition proves that a problem shape is tracked and has a validation plan. It does not prove a private workbook is correct.
- Live workbook evidence must be generated into a task-local temp folder or from a user-supplied sanitized workbook.
- Visual QA is intentionally tracked as a case even though a full automated render QA engine is still a future improvement.

## V1 Case Matrix

| Case | Layer | Why It Exists | Current Evidence Mode |
|---|---|---|---|
| `pq-folder-dynamic-expand-order` | Power Query | Folder ingestion, temp-file filtering, dynamic expansion, duplicate-safe joins, and order restoration are recurring breakpoints. | sanitized spec |
| `dax-excel-powerpivot-compat` | DAX | Excel Power Pivot compatibility differs from Power BI DAX; ratio and dependency checks need guardrails. | sanitized spec |
| `cube-zero-result-debug` | CUBE/MDX | CUBEVALUE zero results require checking MDX members, measure grain, and display formatting separately. | sanitized spec |
| `vba-button-binding-runtime` | VBA | Buttons can look valid while OnAction or macro runtime behavior is broken. | sanitized spec |
| `deliverable-clean-copy` | Deliverable | Clean client copies need pre-clean dependency scan, non-destructive plan, and post-clean verification. | sanitized spec |
| `visual-report-readability` | Visual QA | Structurally valid workbooks can still fail presentation readiness due to clipping, blank report surfaces, or unreadable output areas. | fixture-backed |
| `excel-capability-routing` | Environment | Platform and host questions need capability-first routing and explicit evidence levels. | sanitized spec |
| `pq-refresh-lineage-boundary` | Power Query | Refresh lineage boundaries need source-transition and shape evidence before live refresh. | sanitized spec |
| `dax-measure-rename-impact` | DAX | Measure renames can leave synthetic downstream references unresolved. | sanitized spec |
| `cube-mdx-escaped-member-grain` | CUBE/MDX | Escaped members and measure grain require separate parsing and live-model checks. | sanitized spec |
| `vba-onaction-scope-mismatch` | VBA | OnAction targets can exist in source yet remain uncallable from a workbook control. | sanitized spec |
| `deliverable-external-dependency-readiness` | Deliverable | A copied deliverable needs dependency readiness evidence before release. | sanitized spec |

The five additional cases extend known Excel BI boundaries while remaining deterministic and synthetic. They track repeatable problem shapes and package-tool checks; they cannot replace workbook-behavior proof from a sanitized live workbook in its target environment.

## Validation

```powershell
python tools\run_case_regression.py `
  --project-root . `
  --out-json "$env:TEMP\excel_bi_case_regression.json" `
  --out-md "$env:TEMP\excel_bi_case_regression.md" `
  --require-pass
```

Expected V1 result:

- status: `pass`
- case count: `12`
- covered layers: `cube-mdx`, `dax`, `deliverable`, `environment`, `power-query`, `vba`, `visual-qa`

## Optional Local Observed Usage

The regression library tracks sanitized problem shapes. It is separate from optional, local observed-usage evidence: users who choose to record sanitized workflow events can validate and summarize them with the commands in [Local observed-usage evidence](observed-usage.md). No usage log, Excel installation, or execution telemetry is required for this case-regression workflow or the release gate.

## Next Best Work

Add pixel-level screenshot comparison or more report-surface patterns for the same sanitized fixture. PDF-backed Excel COM export proof is now covered by rendered Visual QA evidence chain V1; exact pixel/render quality comparison is still follow-up hardening.
