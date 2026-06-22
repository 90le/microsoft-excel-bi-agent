# Connection String Templates

Read this file when connecting VBA/ADO to Excel-like files.

## Excel Workbook Through ACE OLEDB

```vb
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\file.xlsx;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";
```

For `.xlsm`:

```vb
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\file.xlsm;Extended Properties="Excel 12.0 Macro;HDR=YES;IMEX=1";
```

For legacy `.xls`:

```vb
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\file.xls;Extended Properties="Excel 8.0;HDR=YES;IMEX=1";
```

## Notes

- ACE provider must be installed.
- 32-bit/64-bit Office mismatch can break providers.
- Sheet names use `[Sheet1$]`; Excel tables may use named ranges.
- Run `tools/probe_excel_bi_providers.ps1 -RunExcelComSmoke -RunAdoWorkbookSmoke` when provider installation or process bitness is uncertain. The full release gate runs this baseline probe automatically on Windows runtime profiles.

## CSV/Text Folder Through ACE OLEDB

Connect to the folder, then query the file name:

```vb
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\folder;Extended Properties="text;HDR=YES;FMT=Delimited";
```

```sql
SELECT * FROM [data.csv]
```

For repeatable production workflows, Power Query is usually safer for CSV encoding, delimiter, and schema drift.

## Access Database

```vb
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\database.accdb;Persist Security Info=False;
```

Legacy `.mdb` files can use the same ACE provider when installed.

## SQL Server

Prefer the current Microsoft OLE DB Driver when installed:

```vb
Provider=MSOLEDBSQL;Data Source=SERVER;Initial Catalog=DATABASE;Integrated Security=SSPI;
```

For SQL authentication:

```vb
Provider=MSOLEDBSQL;Data Source=SERVER;Initial Catalog=DATABASE;User ID=USER;Password=PASSWORD;
```

## Power Pivot / Data Model

Worksheet SQL through ACE does not query the Power Pivot Data Model. Use one of:

- `CUBEVALUE` / `CUBEMEMBER` formulas in Excel.
- Excel COM model metadata tools in this pack.
- ADOMD/MSOLAP only where the provider and model endpoint are available.

Use provider probing before promising ADOMD/MSOLAP behavior:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/probe_excel_bi_providers.ps1 `
  -RunExcelComSmoke `
  -RunAdoWorkbookSmoke `
  -OutJson provider_probe.json
```

For a focused ADOMD COM activation check, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test_excel_adomd_query.ps1 `
  -ProbeOnly `
  -OutJson adomd_probe.json
```

`registryDetected` only means the provider ProgID is registered. It is not the same as a successful connection to a workbook Data Model.
