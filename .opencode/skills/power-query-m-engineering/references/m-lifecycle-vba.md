# Power Query Lifecycle With VBA And Excel COM

Read this file when automating Power Query creation, edits, deletion, refresh, wait-for-finish behavior, performance testing, or error diagnosis. Do not read it for static M code review only.

## Object Model

Power Query queries are exposed through `Workbook.Queries`. A query is a `WorkbookQuery` with a read/write `Formula` property and methods such as `Refresh` and `Delete`.

Typical lifecycle actions:

```vb
' Add
ThisWorkbook.Queries.Add Name:="QueryName", Formula:=mCode

' Read
Debug.Print ThisWorkbook.Queries("QueryName").Formula

' Update
ThisWorkbook.Queries("QueryName").Formula = mCode

' Delete
ThisWorkbook.Queries("QueryName").Delete

' Refresh one query
ThisWorkbook.Queries("QueryName").Refresh
```

When a query is loaded to a worksheet table, `WorkbookQuery.Refresh` can refresh query metadata without proving that the loaded `ListObject` data changed. After editing a query formula, refresh matching load targets through `ListObject.QueryTable.Refresh BackgroundQuery:=False`, then call `Application.CalculateUntilAsyncQueriesDone`, and save a refreshed copy when the updated table values must persist.

## Refresh All And Wait

Basic pattern:

```vb
Public Sub RefreshAllAndWait()
    Dim startedAt As Double
    startedAt = Timer

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull

    Debug.Print "Refresh seconds=" & Format$(Timer - startedAt, "0.00")
End Sub
```

Important details:

- `Workbook.RefreshAll` refreshes external ranges and PivotTables.
- Connections with background refresh enabled can continue after `RefreshAll` returns.
- `Application.CalculateUntilAsyncQueriesDone` waits for pending asynchronous OLEDB/OLAP queries.
- Some prompts, credentials, or source errors can still block or raise runtime errors.

## Load Query To Worksheet Table

To create a workbook query and load it to a worksheet table, use `Workbook.Queries.Add` for the M formula, then create a `ListObject` with a Power Query mashup OLE DB connection:

```vb
Set q = ThisWorkbook.Queries.Add("SmokeQuery", _
    "let Source = #table({""A"",""B""}, {{1,""x""},{2,""y""}}) in Source")

Set lo = Sheet1.ListObjects.Add( _
    SourceType:=0, _
    Source:=Array("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=SmokeQuery;Extended Properties="""""), _
    LinkSource:=True, _
    XlListObjectHasHeaders:=1, _
    Destination:=Sheet1.Range("A1"))

lo.QueryTable.CommandType = 2
lo.QueryTable.CommandText = Array("SELECT * FROM [SmokeQuery]")
lo.QueryTable.BackgroundQuery = False
lo.QueryTable.Refresh BackgroundQuery:=False
```

Use `scripts/create_power_query_fixture_excel_com.ps1` for a tested PowerShell/COM implementation of this pattern. The `Source` argument must be an array for external `ListObject` sources; passing a plain string can produce invalid-connection errors.

## Disable Background Refresh Pattern

When a macro must continue only after loaded tables refresh, disable background refresh where possible:

```vb
Private Sub SetBackgroundRefresh(ByVal enabled As Boolean)
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            On Error Resume Next
            lo.QueryTable.BackgroundQuery = enabled
            On Error GoTo 0
        Next lo
    Next ws
End Sub
```

Then:

```vb
SetBackgroundRefresh False
ThisWorkbook.RefreshAll
Application.CalculateUntilAsyncQueriesDone
```

Do not assume every Power Query load target exposes a `QueryTable`; Data Model only queries may not appear as worksheet query tables.

If refreshing one named query, locate worksheet `ListObject` objects whose `QueryTable.WorkbookConnection.Name`, connection text, or command text references the query. Refresh those `QueryTable` objects directly. Use `WorkbookQuery.Refresh` only as a fallback when the query has no worksheet load target.

## Run Next Step After Refresh

```vb
Public Sub RefreshThenRunReport()
    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    SetBackgroundRefresh False
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull

    RunReportAfterRefresh

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "Power Query refresh failed before report step:" & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

## Performance Test Pattern

```vb
Public Sub BenchmarkRefresh()
    Dim startedAt As Double
    startedAt = Timer

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    Debug.Print "Refresh completed at " & Now
    Debug.Print "Elapsed seconds: " & Format$(Timer - startedAt, "0.00")
End Sub
```

For repeatable benchmarks:

- Close other heavy workbooks.
- Run twice and compare warm vs cold source behavior.
- Record file count, row count, and source location.
- Disable screen updating and events.
- Avoid changing workbook formulas during the refresh benchmark.

## Error Diagnosis

When a refresh error appears:

1. Export all query formulas.
2. Identify which query failed in Excel's Queries & Connections pane.
3. Open that query in Power Query Editor and locate the first failing step.
4. Check credentials, source path, missing columns, type conversions, and join cardinality.
5. Add temporary diagnostic steps such as row counts or duplicate-key checks.
6. Refresh only the failing query before refreshing all.

## Common Popup Causes

| Popup or symptom | Likely cause | Response |
|---|---|---|
| Credentials prompt | Data source privacy/auth changed | Re-authenticate in Excel, do not mask with VBA |
| Formula.Firewall | Privacy levels or query combination | Review privacy settings and staged queries |
| Column not found | Source schema drift | Use schema normalization or missing-field handling |
| Type conversion error | Invalid value in source | Guard conversion with `try ... otherwise` |
| Macro continues before data loaded | Background refresh | Disable background refresh where possible and call wait method |

## Validation Rule

After refresh automation, validate the downstream output, not just that no VBA error occurred. Check row counts, key output cells, and timestamps.
