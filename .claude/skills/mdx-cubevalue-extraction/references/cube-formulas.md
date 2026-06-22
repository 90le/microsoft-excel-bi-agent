# Excel CUBE Formula Patterns

Read this file when creating or debugging `CUBEVALUE`, `CUBEMEMBER`, `CUBESET`, or `CUBERANKEDMEMBER`.

## ADOMD Endpoint Query Boundary

Use `tools/test_excel_adomd_query.ps1` only when a real ADOMD/MSOLAP endpoint connection string is available.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_adomd_query.ps1 `
  -ProbeOnly `
  -OutJson adomd_probe.json
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_adomd_query.ps1 `
  -ConnectionString 'Provider=MSOLAP;Data Source=server;Initial Catalog=model;' `
  -Mdx 'SELECT [Measures].[Sales] ON 0 FROM [Model]' `
  -OutJson adomd_query.json
```

This validates an endpoint query. It does not directly evaluate Excel `CUBEVALUE` formulas or expose a workbook's in-process `ThisWorkbookDataModel`.

## Basic Measure

```excel
=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Sales]")
```

## Measure With Member

```excel
=CUBEVALUE(
  "ThisWorkbookDataModel",
  "[Measures].[Sales]",
  "[Date].[Year].&[2026]"
)
```

## Helper Cell Pattern

```excel
A1 = CUBEMEMBER("ThisWorkbookDataModel","[Date].[Year].&[2026]")
B1 = CUBEVALUE("ThisWorkbookDataModel","[Measures].[Sales]",A1)
```

Helper cells are easier to audit and reuse than long nested strings.

## Parameterized Member

```excel
=CUBEMEMBER(
  "ThisWorkbookDataModel",
  "[Period].[Period].&[" & $B$2 & "]"
)
```

Validate the resulting member unique name in Excel.
