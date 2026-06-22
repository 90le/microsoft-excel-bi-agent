# Delivery Checklist

Use this checklist before telling the user a workbook is client-ready.

## Delivery Shapes

| Shape | Keep formulas | Keep VBA | Keep Power Query/Data Model | Required proof |
|---|---:|---:|---:|---|
| Pure `.xlsx` | No | No | No | Post-clean dependency report plus pure-deliverable verification |
| Live `.xlsm` | Maybe | Yes | Maybe | Runtime validation, macro/button binding evidence, dependency note |
| Internal review copy | Maybe | Maybe | Maybe | Clear owner-facing risk list and next validation command |

## Non-Destructive Rule

Always work from a copy. Do not remove formulas, links, queries, model parts, macros, or process sheets from the only source workbook.

## Publish Note Template

```text
Source workbook:
Delivery workbook:
Delivery shape:
Refresh/recalculate evidence:
Value-freeze status:
Removed items:
Intentionally preserved items:
Post-clean verification:
Known boundaries:
```

## Blockers

- External links or credential-like connection strings remain in a pure deliverable.
- Formula cells remain in a values-only workbook.
- Power Query or model metadata remains when the requested shape is pure `.xlsx`.
- VBA remains in a macro-free deliverable.
- Cleanup was performed on the source workbook instead of a copy.
