# Power Query M Patterns

Read this file when creating or editing M transformations. Do not read it for pure DAX, MDX, or VBA tasks.

## Duplicate-Safe Lookup Join

Pre-aggregate lookup rows before joining back to a main table:

```powerquery
LookupGrouped =
    Table.Group(
        Lookup,
        {"Key"},
        {{"LookupRows", each _, type table}}
    ),
Joined =
    Table.NestedJoin(Main, {"Key"}, LookupGrouped, {"Key"}, "Lookup", JoinKind.LeftOuter)
```

Use this when duplicate lookup keys could multiply the main table.

## Preserve Row Order Through Group/Join

Add an index before operations that may reorder rows:

```powerquery
WithOrder = Table.AddIndexColumn(Source, "__OriginalOrder", 0, 1, Int64.Type),
...
Restored = Table.Sort(Result, {{"__OriginalOrder", Order.Ascending}})
```

Remove helper order columns only in the final delivery step.

## Folder Ingestion Filters

Filter early:

```powerquery
Filtered =
    Table.SelectRows(
        Folder.Files(FolderPath),
        each [Attributes]?[Hidden]? <> true
            and not Text.StartsWith([Name], "~$")
            and List.Contains({".xlsx", ".xlsm", ".xls", ".csv"}, Text.Lower([Extension]))
    )
```

Before importing changed M source, run the static lint:

```powershell
python .agents\skills\power-query-m-engineering\scripts\lint_power_query_m.py "src\m" --out-json "tmp\m-lint.json"
```

Use `--warnings-as-errors` for delivery gates when risky join, expand, or folder-ingestion patterns should block import.

## Safe Column Selection

Use `MissingField.UseNull` when source schemas drift:

```powerquery
Table.SelectColumns(Source, {"A", "B", "C"}, MissingField.UseNull)
```

## Dynamic Expand Columns

When the nested table schema can change, derive expansion fields from a sample table instead of hard-coding a stale list:

```powerquery
ColumnsToExpand = Table.ColumnNames(Combined{0}[Content]),
Expanded = Table.ExpandTableColumn(Combined, "Content", ColumnsToExpand, ColumnsToExpand)
```

Guard the sample access when the source can be empty.

## Guard Empty Lists

For latest-period logic, protect `List.Max`:

```powerquery
Latest =
    if List.IsEmpty(Periods)
    then null
    else List.Max(Periods)
```

## Final Delivery Cleanup

Keep upstream source tables intact. Normalize blanks, rename output fields, and trim columns in the final output query.
