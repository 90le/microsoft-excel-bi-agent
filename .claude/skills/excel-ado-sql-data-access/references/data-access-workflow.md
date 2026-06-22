# ADO / SQL Data Access Workflow

Use this reference when Excel VBA or automation needs to read data through ADO, OLEDB, ADODB, ADOMD, or SQL-like queries.

## Source Decision Matrix

| Source | Preferred Access | Notes |
|---|---|---|
| Saved Excel worksheet range | ACE OLEDB + `SELECT * FROM [Sheet$]` | Good for closed workbook reads and simple filtering/grouping. |
| Excel ListObject/table | ACE OLEDB if exposed as named range; otherwise use sheet range | Test table visibility through schema before relying on table names. |
| CSV/Text folder | ACE Text driver or Power Query | SQL over CSV requires folder-level connection and schema assumptions. |
| Access `.accdb/.mdb` | ACE OLEDB | Use parameterized `ADODB.Command` where possible. |
| SQL Server | MSOLEDBSQL / ODBC / Power Query depending on workflow | Prefer native database tools for large data and query folding. |
| Power Query output | Query the loaded sheet/table, not the M query itself | Power Query is not a database; refresh/load first. |
| Power Pivot/Data Model | ADOMD/MSOLAP where installed, or CUBE formulas / Excel COM metadata | Do not treat the Data Model as a worksheet table. |

## Standard Workbook SQL Workflow

1. Probe the local provider environment when the machine is new or the user reports provider errors.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/probe_excel_bi_providers.ps1 `
  -RunExcelComSmoke `
  -RunAdoWorkbookSmoke `
  -OutJson provider_probe.json
```

This records ACE/MSOLAP/ADOMD registration, COM activation, process bitness, optional Excel COM smoke status, and optional ACE workbook SQL smoke status. The full release gate runs the same probe with both smoke switches and verifies the generic provider baseline when the local Windows runtime is available.

2. Inspect workbook structure.

```bash
python tools/inspect_excel_bi_workbook.py workbook.xlsx --markdown
```

3. Test provider and SQL with a read-only query.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_ado_sql_access.ps1 `
  -WorkbookPath workbook.xlsx `
  -SqlText 'SELECT * FROM [Data$]' `
  -IncludeSchema `
  -OutJson ado_query.json
```

4. In Windows Git Bash, use:

```bash
tools/invoke_excel_bi_com.sh provider-probe --excel-com --ado-workbook-smoke --out-json provider_probe.json

tools/invoke_excel_bi_com.sh ado-query \
  -w workbook.xlsx \
  --sql 'SELECT * FROM [Data$]' \
  --include-schema \
  --out-json ado_query.json
```

5. Only after the provider, schema, and row counts are validated, write or modify VBA code.

## Fixture Smoke Test

Create a known workbook and query it:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_ado_sql_access.ps1 `
  -WorkbookPath "$env:TEMP\ado_fixture.xlsx" `
  -CreateFixture `
  -SqlText 'SELECT Region, SUM(Amount) AS TotalAmount FROM [Data$] GROUP BY Region ORDER BY Region' `
  -OutJson "$env:TEMP\ado_fixture_query.json"
```

Expected result:

- `succeeded = true`
- `rowCount = 3`
- Fields include `Region` and `TotalAmount`

## ADOMD Endpoint Query Smoke

Use this only when there is a real ADOMD/MSOLAP endpoint connection string. The script can probe COM activation without an endpoint; the full release gate runs this probe in `full` profile and returns `skip` when the local ADOMD/MSOLAP runtime is unavailable:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_adomd_query.ps1 `
  -ProbeOnly `
  -OutJson adomd_probe.json
```

Run an MDX query against an explicit endpoint:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_adomd_query.ps1 `
  -ConnectionString 'Provider=MSOLAP;Data Source=server;Initial Catalog=model;' `
  -Mdx 'SELECT [Measures].[Sales] ON 0 FROM [Model]' `
  -MaxCells 100 `
  -OutJson adomd_query.json
```

In Windows Git Bash:

```bash
tools/invoke_excel_bi_com.sh adomd-query \
  --probe-only \
  --out-json adomd_probe.json

tools/invoke_excel_bi_com.sh adomd-query \
  --connection-string 'Provider=MSOLAP;Data Source=server;Initial Catalog=model;' \
  --mdx 'SELECT [Measures].[Sales] ON 0 FROM [Model]' \
  --max-cells 100 \
  --out-json adomd_query.json
```

The output redacts password-like connection-string keys and records axes, returned cells, truncation, and errors. It does not create an endpoint for Excel's in-process `ThisWorkbookDataModel`.

## VBA Implementation Pattern

Use late binding unless the workbook already has ADODB references.

```vb
Dim cn As Object
Dim rs As Object

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Open connectionString
rs.Open sqlText, cn, 0, 1

If Not rs.EOF Then
    targetRange.CopyFromRecordset rs
End If

If rs.State <> 0 Then rs.Close
If cn.State <> 0 Then cn.Close
Set rs = Nothing
Set cn = Nothing
```

## Common Failure Modes

| Symptom | Likely Cause | Fix |
|---|---|---|
| Provider cannot be found | ACE/MSOLEDBSQL/driver not installed or bitness mismatch | Install matching provider or use Power Query/OpenXML fallback. |
| No value given for one or more required parameters | Field/table name not found, invalid sheet name, or SQL dialect mismatch | Use `-IncludeSchema`, quote names, and test `SELECT *` first. |
| Empty rows or wrong types | Mixed column types and ACE inference | Use `IMEX=1`, clean source types, or load through Power Query. |
| File locked | Connection/recordset not closed | Always close in cleanup/finally blocks. |
| Data Model not queryable | Trying worksheet SQL against Power Pivot model | Use ADOMD/MSOLAP, CUBE formulas, or Excel COM model metadata. |
| Works on one machine but not another | Provider registration, Office bitness, or ADOMD component mismatch | Run `probe_excel_bi_providers.ps1` on both machines and compare JSON. |

## Validation Boundary

ADO SQL tests prove provider connectivity and saved-file query behavior. They do not prove that:

- Power Query has refreshed.
- The in-memory workbook state matches the saved file.
- Power Pivot/Data Model calculations are correct.
- ADOMD/MSOLAP providers are installed.

Provider probes prove local environment facts; they do not prove a specific workbook query is semantically correct.

ADOMD query tests prove a specific endpoint and MDX query can execute. They do not prove Excel worksheet `CUBEVALUE` formulas recalculate correctly inside a workbook.
