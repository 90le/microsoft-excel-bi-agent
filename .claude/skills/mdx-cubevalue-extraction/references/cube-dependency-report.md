# CUBE Dependency Report

Use this reference when a workbook has `CUBEVALUE`, `CUBEMEMBER`, `CUBESET`, or related CUBE formulas that must be traced back to model measures, member paths, helper cells, or report sheets.

## Purpose

The dependency report answers:

- Which sheets contain CUBE formulas?
- Which cells call which model measures?
- Which member references appear most often?
- Which formulas use dynamic helper-cell references?
- Which formulas reference measures that are missing from the Data Model?
- Which formulas use hard-coded latest/previous period markers that may become stale?

## Standard Workflow

1. Inspect workbook structure.

```bash
python tools/inspect_excel_bi_workbook.py workbook.xlsx --out-json openxml.json
```

2. Optional but recommended: export Data Model metadata on Windows desktop Excel.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/inspect_excel_data_model_com.ps1 `
  -WorkbookPath workbook.xlsx `
  -OutJson data_model.json
```

3. Build the CUBE dependency report.

```bash
python tools/build_cube_dependency_report.py \
  --openxml-json openxml.json \
  --model-json data_model.json \
  --out-md cube_dependencies.md \
  --out-json cube_dependencies.json \
  --out-mermaid cube_dependencies.mmd
```

4. In Windows Git Bash, use the combined wrapper.

```bash
tools/invoke_excel_bi_com.sh cube-report \
  -w workbook.xlsx \
  --include-model \
  --out-md cube_dependencies.md \
  --out-json cube_dependencies.json \
  --out-mermaid cube_dependencies.mmd
```

## Diagnostic Flags

| Flag | Meaning | Action |
|---|---|---|
| `measure_not_found_in_model` | Formula references a measure name absent from exported model metadata. | Fix the formula measure path or restore/create the model measure. |
| `hard_coded_period_marker` | Formula contains direct markers such as `[All].[new]` or `[All].[-1]`. | Use visible helper cells or named ranges when business users need to audit period selection. |
| `dynamic_mdx_string` | Formula builds MDX strings through `&` concatenation. | Prefer helper `CUBEMEMBER` cells for reusable or business-facing selectors. |
| `long_formula_without_helper_cells` | Long CUBE formula has no obvious helper cell references. | Split member paths into helper cells to make the report easier to audit. |
| `cubevalue_without_measure` | `CUBEVALUE` has no `[Measures].[...]` argument. | Confirm whether it should be a member formula or add the correct measure. |
| `error_cached_value` | Cached OpenXML value is an Excel error. | Refresh in Excel and validate member paths, model refresh, and measure names. |

## Rewrite Principles

- Keep measure names stable and verify against the Data Model before changing formulas.
- Put reusable period, product, sponsor, region, or category members into helper `CUBEMEMBER` cells.
- Use direct `CUBEVALUE` strings only for simple formulas that are unlikely to be reused or audited.
- Treat OpenXML cached values as clues, not proof; final formula correctness requires desktop Excel validation.

## Validation Boundary

The report is structural. It does not refresh Power Query, recalculate the Data Model, or evaluate CUBE formulas. Use it to identify dependencies and risks, then validate final values in Excel.
