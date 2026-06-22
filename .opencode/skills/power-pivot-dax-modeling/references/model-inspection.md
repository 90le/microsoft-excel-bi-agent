# Power Pivot Model Inspection

Use this reference when a workbook contains a Power Pivot Data Model, DAX measures, PivotTables, or CUBE formulas that depend on `ThisWorkbookDataModel`.

## Inspection Levels

| Level | Tool | Works on | Purpose |
|---|---|---|---|
| Workbook structure | `tools/inspect_excel_bi_workbook.py` | Windows, Linux, macOS | Detect sheets, connections, PivotCaches, formulas, CUBE formulas, Power Query-like parts, and Data Model-like package parts. |
| Data Model metadata | `tools/inspect_excel_data_model_com.ps1` | Windows desktop Excel | Read model tables, columns, relationships, measures, and workbook connections through Excel COM. |
| Model report | `tools/build_excel_bi_model_report.py` | Windows, Linux, macOS after JSON export | Combine Data Model JSON and OpenXML JSON into a readable Markdown report with measure usage and relationship graph. |
| Git Bash workflow | `tools/invoke_excel_bi_com.sh model-report` | Windows Git Bash/MSYS/Cygwin | Run Data Model export, OpenXML inspection, and Markdown report generation from one command. |

## Standard Workflow

1. On a new machine or when provider errors are reported, record the local provider environment.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/probe_excel_bi_providers.ps1 `
  -RunExcelComSmoke `
  -RunAdoWorkbookSmoke `
  -OutJson provider_probe.json
```

Check Excel COM availability, process bitness, MSOLAP ProgID registration, ADOMD COM activation, .NET ADOMD assembly availability, and generated-workbook ACE/ADO smoke status before attributing failures to workbook logic. The full release gate runs this provider baseline automatically in the Windows `full` profile and skips it in structural profile.

2. Run OpenXML inspection first.

```bash
python tools/inspect_excel_bi_workbook.py workbook.xlsx --out-json openxml.json --markdown
```

3. On Windows desktop Excel, export the actual Data Model metadata.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/inspect_excel_data_model_com.ps1 `
  -WorkbookPath workbook.xlsx `
  -IncludeColumns `
  -OutJson data_model.json
```

4. Build a human-readable model report.

```bash
python tools/build_excel_bi_model_report.py \
  --model-json data_model.json \
  --openxml-json openxml.json \
  --out-md model_report.md \
  --out-json model_summary.json
```

5. In Windows Git Bash, use the combined wrapper.

```bash
tools/invoke_excel_bi_com.sh provider-probe --excel-com --out-json provider_probe.json

tools/invoke_excel_bi_com.sh model-report \
  -w workbook.xlsx \
  --include-columns \
  --out-md model_report.md \
  --out-json model_summary.json
```

## What To Check Before Editing DAX

- Does each measure belong to the expected associated table?
- Does the measure formula match the expected reporting grain?
- Do relationships point from fact-like tables to lookup/dimension tables?
- Are relationship keys active and business-meaningful?
- Are CUBE formulas referencing the expected measures?
- Are any CUBE formula measure references missing from the model?
- Are any model measures unused by CUBE formulas but still business-critical for PivotTables?

## Warning Signs

- A measure uses row-level logic where a calculated column is needed.
- A report CUBE formula references `[Measures].[...]` that is not present in the model.
- `CUBEVALUE` formulas depend on hard-coded members that should be helper-cell driven.
- A relationship is inactive or points at an unexpected lookup table.
- Multiple model tables contain similar columns but only one is connected to the reporting layer.

## Validation Boundary

The Markdown report is structural. It can prove that formulas, measures, relationships, and references exist, but it cannot prove returned values are correct. After changing DAX or CUBE formulas, validate final numbers in desktop Excel.

Provider probes are environment diagnostics. They can explain why Data Model or ADOMD automation fails on a given machine, but they do not prove a DAX measure is correct.
