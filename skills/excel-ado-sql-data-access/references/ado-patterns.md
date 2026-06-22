# ADO Patterns

Read this file when writing VBA ADO/OLEDB/ADOMD code.

## Basic ADODB Pattern

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

rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
```

Use late binding unless the workbook already has ADODB references.

## Error Handling

Always close resources in cleanup blocks. Do not leave external workbook files locked.

## Parameterized Command Pattern

Use this pattern for providers that support parameters, such as Access or SQL Server. Excel worksheet SQL has limited parameter support; validate provider behavior first.

```vb
Dim cn As Object
Dim cmd As Object
Dim rs As Object

Set cn = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")

cn.Open connectionString

Set cmd.ActiveConnection = cn
cmd.CommandText = "SELECT * FROM Sales WHERE Region = ?"
cmd.CommandType = 1
cmd.Parameters.Append cmd.CreateParameter("Region", 200, 1, 255, regionValue)

Set rs = cmd.Execute
```

## Recordset To Array Pattern

Use arrays when the result will be transformed before writing to a sheet.

```vb
If Not rs.EOF Then
    data = rs.GetRows()
End If
```

`GetRows` returns a field-major array: `data(fieldIndex, rowIndex)`.

## Workbook SQL Guardrails

- Save and close source workbooks before querying them through ACE when possible.
- Use `[SheetName$]` for worksheet ranges.
- Start with `SELECT * FROM [SheetName$]` before adding filters or grouping.
- If SQL reports missing parameters, inspect schema; it usually means the field or range name was not found.
- Do not query Power Query M definitions directly through ADO; query the loaded worksheet/table or Data Model endpoint instead.
