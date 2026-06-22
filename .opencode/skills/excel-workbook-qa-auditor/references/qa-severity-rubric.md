# QA Severity Rubric

Use this rubric to keep workbook audit findings consistent.

## Severity Levels

| Severity | Meaning | Examples |
|---|---|---|
| Blocker | Must resolve before delivery or refactor | Broken formula, unresolved macro button, credential-like connection string in deliverable, failed required refresh |
| High | Likely to break user trust or automation | Hard-coded local path, unexpected external workbook link, missing model measure, stale helper-cell dependency |
| Medium | Needs owner confirmation | Hidden or protected sheet, volatile formula, dynamic reference, mixed-source query lineage |
| Low | Improve maintainability | Naming, formatting, documentation, redundant helper ranges |
| Accepted | Known and approved | Intentional macro-enabled workbook, intentional live refresh, deliberate hidden config sheet |

## Finding Format

```text
Severity:
Surface:
Location:
Finding:
Why it matters:
Evidence command/report:
Recommended next action:
Runtime boundary:
```

## Audit Boundary

Static reports show structure and risk. They do not prove final numeric correctness, refresh credentials, Data Model semantics, or VBA runtime behavior.
