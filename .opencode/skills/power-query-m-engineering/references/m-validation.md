# Power Query M Validation

Read this file when validating a Power Query change or workbook deliverable.

## Minimum Validation Table

| Check | Why |
|---|---|
| Source file count | Detect missing/extra files and temp files |
| Source schema | Detect renamed or missing columns |
| Row count before and after joins | Catch row multiplication |
| Row count before and after grouping | Catch accidental aggregation |
| Key uniqueness | Confirm join assumptions |
| Final column names | Match delivery contract |
| Final data types | Avoid refresh/runtime type errors |
| Sort/order stability | Important for report-facing outputs |

## Join Validation

For a left join expected to preserve main rows:

```powerquery
MainCount = Table.RowCount(Main),
JoinedCount = Table.RowCount(ExpandedJoin)
```

The two counts should match unless the business rule intentionally duplicates rows.

## Duplicate Key Check

```powerquery
KeyCounts =
    Table.Group(
        Lookup,
        {"Key"},
        {{"Count", each Table.RowCount(_), Int64.Type}}
    ),
Duplicates = Table.SelectRows(KeyCounts, each [Count] > 1)
```

Use this before expanding lookup columns.

## Output Contract

Record:

```text
Query:
Source count:
Final row count:
Final columns:
Known intentional duplicates:
Refresh environment:
```

If refresh was not tested in Excel or Power BI, state that clearly.
