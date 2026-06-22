# Power Query M Troubleshooting

Read this file when a Power Query refresh, formula, row count, or type conversion fails.

## Common Errors

| Error | Likely cause | Fix |
|---|---|---|
| The column was not found | Source schema drift or renamed header | Use `MissingField.UseNull`, header detection, or schema normalization |
| We cannot convert value null to type Text/Number | Type conversion before null handling | Replace nulls or use guarded conversion |
| Expression.Error: key did not match any rows | Navigating workbook item by hard-coded sheet/table name | Inspect `Excel.Workbook` output and select by kind/name defensively |
| Row count increased unexpectedly | Join expanded duplicate lookup rows | Group lookup table before join |
| Output order changed | Group/join/combine operation reordered rows | Add and restore original order index |

## Refresh Debug Steps

1. Classify the refresh JSON or copied popup text:

```powershell
python .agents\skills\power-query-m-engineering\scripts\classify_power_query_refresh_errors.py `
  "tmp\refresh.json" `
  --out-json "tmp\refresh-diagnosis.json" `
  --out-md "tmp\refresh-diagnosis.md"
```

2. Isolate the first failing step.
3. Inspect row count and schema at that step.
4. Check source file filters and hidden/temp files.
5. Verify header rows and promoted headers.
6. Verify join key uniqueness.
7. Test with a small subset of files.

## Classification Buckets

| Code | Meaning | First response |
|---|---|---|
| `credentials-or-permissions` | Source login or permission issue | Re-authenticate in Excel Data Source Settings |
| `privacy-firewall` | Query combination blocked by privacy rules | Review privacy levels and staging queries |
| `missing-source` | File, folder, sheet, table, or external source missing | Verify source paths and source inventory |
| `query-not-found` | Requested WorkbookQuery name is absent | List workbook queries and use the exact name |
| `missing-column` | Source schema drift or stale hard-coded column list | Inspect headers and use `MissingField.UseNull` where valid |
| `missing-workbook-item` | Hard-coded Excel.Workbook navigation key failed | Inspect `Excel.Workbook` output before navigation |
| `type-conversion` | Dirty value cannot convert to target type | Guard conversions with `try ... otherwise` |
| `syntax-or-formula` | M syntax, identifier, or dependency issue | Run `lint_power_query_m.py` and inspect step commas/final `in` |
| `connector-provider` | Driver, provider, OLE DB, ODBC, or Mashup provider failure | Run provider probe and check Office/provider bitness |
| `timeout-or-background-refresh` | Slow source, async refresh, or blocked connection | Disable background refresh and refresh only the failing query |
| `row-count-or-cardinality` | Join/expand may multiply rows | Pre-aggregate lookup side before `Table.NestedJoin` |

## Type Conversion Pattern

Prefer explicit guarded conversions:

```powerquery
Table.TransformColumns(
    Source,
    {{"Amount", each try Number.From(_) otherwise null, type nullable number}}
)
```
