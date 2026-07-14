# Local Observed-Usage Evidence

Observed usage is optional local evidence for teams that want to count sanitized workflow outcomes over time. It is not telemetry: the package does not create a log, send data anywhere, or require a log for release validation. Record only after you sanitize the event yourself.

## Privacy Boundary

Keep local JSONL files under `.local-observed-usage/`; that directory and `observed-usage*.jsonl` are ignored by Git. Never commit them. Do not record customer workbook names or contents, screenshots, local/network paths, credentials, tokens, connection strings, personal data, or business-specific rules. Use opaque case IDs and generic skill/outcome labels instead.

## Schema

Each nonblank line is one JSON object with these required fields:

| Field | Meaning |
| --- | --- |
| `schemaVersion` | Exactly `1.0`. |
| `eventId` | Unique opaque event identifier. |
| `recordedAt` | UTC ISO-8601 timestamp ending in `Z`. |
| `caseId` | Sanitized, opaque case label. |
| `requestedSkill` / `selectedSkill` | Requested and selected package skill IDs. |
| `outcome` | Sanitized outcome label, such as `completed` or `needs-review`. |
| `durationMs` | Nonnegative integer duration in milliseconds. |
| `evidenceLevel` | `structural`, `runtime-capability`, or `workbook-behavior`. |
| `evidenceBoundary` | A short sanitized boundary label. `workbook-behavior` must use `local-user-supplied-sanitized`. |

Synthetic example only (do not copy real names, paths, or content into logs):

```json
{"schemaVersion":"1.0","eventId":"sample-001","recordedAt":"2026-01-15T09:30:00Z","caseId":"synthetic-pq-01","requestedSkill":"excel-bi-router","selectedSkill":"power-query-m-engineering","outcome":"completed","durationMs":42000,"evidenceLevel":"structural","evidenceBoundary":"synthetic-local-example"}
```

## Validate and Summarize

Validate local data before using it:

```powershell
python tools\validate_observed_usage.py .local-observed-usage\observed-usage.jsonl
```

Create an aggregate-only summary after validation:

```powershell
python tools\summarize_observed_usage.py .local-observed-usage\observed-usage.jsonl
```

The validator rejects path-like text, credentials, and spreadsheet artifact names anywhere in an event. That is a guardrail, not a substitute for user sanitization: review and sanitize every field before recording it. The summary reports counts and duration aggregates only; it does not prove Excel execution, workbook behavior, or customer outcomes.
