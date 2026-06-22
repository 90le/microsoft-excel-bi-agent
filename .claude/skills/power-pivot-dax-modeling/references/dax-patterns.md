# DAX Patterns

Read this file for common Excel Power Pivot measure patterns.

## Ratio Measure

```dax
Share :=
DIVIDE([Numerator], [Denominator])
```

Use `DIVIDE` instead of `/` when denominator can be zero or blank.

## Filter Removal

```dax
Share of Total :=
DIVIDE(
    [Sales],
    CALCULATE([Sales], ALL('DimProduct'))
)
```

Prefer `ALL` for Excel Power Pivot compatibility. Use `REMOVEFILTERS` only when the target is confirmed to support it, such as modern Power BI models.

Before using a DAX expression in an Excel Power Pivot workbook, run:

```bash
python .agents/skills/power-pivot-dax-modeling/scripts/lint_dax_compat.py measures.dax --profile excel --warn-division --out-json tmp/dax-lint.json
```

This static lint separates blocking compatibility errors from review warnings. `REMOVEFILTERS` is blocked for Excel-oriented work unless the target host is confirmed to support it; `SELECTEDVALUE` is reported as version-sensitive; `/` is reported with `--warn-division` so ratio measures can be reviewed for `DIVIDE`. The lint does not validate relationships, evaluate measures, or replace Excel runtime testing.

## Selected Value

```dax
Selected Period :=
SELECTEDVALUE('Period'[Period], "All")
```

For older Excel Power Pivot versions, use:

```dax
Selected Period :=
IF(
    HASONEVALUE('Period'[Period]),
    VALUES('Period'[Period]),
    "All"
)
```

## Iterator

```dax
Weighted Score :=
SUMX(
    'Fact',
    'Fact'[Weight] * 'Fact'[Score]
)
```

## Defensive Blank Handling

Use blanks intentionally. Do not turn all blanks into zero unless the report meaning requires it.
