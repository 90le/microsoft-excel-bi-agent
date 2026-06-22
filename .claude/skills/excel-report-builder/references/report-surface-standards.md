# Report Surface Standards

Use these standards for Excel report or dashboard sheets.

## Sheet Roles

- `Input`: user-editable parameters, clearly colored or validated.
- `Calc`: helper calculations, normally hidden after validation.
- `Output`: client-facing tables, charts, and commentary.
- `QA`: optional reviewer checks, removed or hidden before final delivery.

## Layout Defaults

- Freeze panes on wide tables.
- Use stable table/named-range references for charts and controls.
- Keep units and period labels visible.
- Use consistent number formats; avoid mixed text/number output in calculation cells.
- Avoid placing process notes on final-facing sheets.

## Validation Before Publishing

- Inspect workbook structure.
- Run formula quality report.
- Run workbook controls report.
- Run external dependency report if the workbook will leave the authoring environment.
- Reopen in Excel when charts, pivots, VBA, Power Query, or Data Model behavior matters.
