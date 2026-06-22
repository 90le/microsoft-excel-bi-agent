# DAX Context Notes

Read this file when a DAX calculation involves filters, slicers, relationships, or unexpected totals.

## Row Context vs Filter Context

- Row context exists while evaluating a calculated column or iterator row.
- Filter context comes from PivotTables, slicers, report filters, CUBE formulas, and `CALCULATE`.
- Measures are evaluated in filter context.

## CALCULATE

`CALCULATE` evaluates an expression in a modified filter context.

Use it deliberately:

```dax
Sales LY :=
CALCULATE(
    [Sales],
    SAMEPERIODLASTYEAR('Date'[Date])
)
```

## Common Total Issue

If row values look right but grand total looks wrong, the measure likely needs an iterator:

```dax
Total Corrected :=
SUMX(
    VALUES('Dim'[Group]),
    [Measure Per Group]
)
```

## Relationship Checks

- Confirm one-to-many direction.
- Confirm active relationship.
- Avoid ambiguous many-to-many behavior unless explicitly modeled.
