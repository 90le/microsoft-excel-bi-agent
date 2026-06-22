# Official Documentation Knowledge Base

Read this file when exact Power Query M, Excel-hosted Power Query, refresh, or performance semantics are needed. This is a routing index, not a copy of Microsoft documentation.

## Lookup Protocol

1. Search the local index first:

   ```bash
   python scripts/search_power_query_official_kb.py "Table.NestedJoin"
   python scripts/search_power_query_official_kb.py "RefreshAll"
   python scripts/search_power_query_official_kb.py "query folding"
   ```

2. Read the matching local reference file in this skill:

   - M code pattern: `references/m-patterns.md`
   - M style rewrite: `references/m-style-guide.md`
   - Validation and row-count safety: `references/m-validation.md`
   - Excel COM lifecycle and refresh waiting: `references/m-lifecycle-vba.md`
   - Error triage: `references/m-troubleshooting.md`

3. Open the official Microsoft Learn URL from `official-docs-index.json` when:

   - exact function parameters matter,
   - a host behavior is version-sensitive,
   - a refresh/popup/credential issue is being diagnosed,
   - a claim needs a customer-facing citation,
   - a local note conflicts with current Microsoft documentation.

4. If a URL moved, search Microsoft Learn with the `online_query` field from the JSON index. Keep the query restricted to official Microsoft domains unless the user explicitly asks for broader research.

## Fast Routing

| Need | Local first | Official lookup id |
|---|---|---|
| Unknown M function | `official-docs-index.json` | `m-function-reference` |
| Table joins, grouping, expansion, missing columns | `m-patterns.md`, `m-validation.md` | `table-functions` |
| Folder/file/Excel/CSV/API source loading | `m-patterns.md` | `accessing-data-functions`, `excel-workbook`, `csv-document`, `web-contents` |
| Query folding or slow refresh | `m-style-guide.md`, `m-lifecycle-vba.md` | `query-folding-basics`, `power-query-best-practices`, `query-diagnostics` |
| Refresh from VBA and wait for completion | `m-lifecycle-vba.md` | `workbook-refreshall`, `calculate-until-async-queries-done`, `querytable-refresh`, `querytable-backgroundquery` |
| Popup errors, source prompts, type inference | `m-troubleshooting.md` | `common-authoring-issues`, `querytable-refresh` |
| Connector-specific limits or credentials | `official-docs-index.json` | `power-query-connectors` |

## Online Search Templates

Use these when browsing is available:

```text
site:learn.microsoft.com/en-us/powerquery-m <FunctionName>
site:learn.microsoft.com/en-us/power-query <topic>
site:learn.microsoft.com/en-us/office/vba/api/excel <ExcelObjectOrMethod>
```

Examples:

```text
site:learn.microsoft.com/en-us/powerquery-m Table.TransformColumnTypes MissingField
site:learn.microsoft.com/en-us/power-query query diagnostics multiple evaluations
site:learn.microsoft.com/en-us/office/vba/api/excel CalculateUntilAsyncQueriesDone
```

## Local Search Templates

Use these from the skill folder or project root:

```bash
rg -n "Table.NestedJoin|Table.Group|Table.Buffer|MissingField" references scripts
rg -n "RefreshAll|CalculateUntilAsyncQueriesDone|BackgroundQuery|QueryTable" references scripts
rg -n "column not found|type conversion|Formula.Firewall|credentials" references scripts
python scripts/search_power_query_official_kb.py "Folder.Files"
python scripts/search_power_query_official_kb.py "BackgroundQuery" --json
```

## Documentation Hygiene

- Do not paste long Microsoft documentation into the skill. Store stable URLs, search keys, and operational interpretation.
- Prefer source-specific guidance over generic advice. Example: for Excel desktop automation, check Excel VBA object model pages, not only Power Query M pages.
- When producing a customer-facing explanation, distinguish three layers:
  - M formula semantics.
  - Power Query engine behavior.
  - Excel host automation behavior.
- When a workbook deliverable depends on refresh, verify in desktop Excel if possible. OpenXML inspection can prove structure, but it cannot prove credentials, refresh completion, or M engine execution.

## Source Families

### Power Query M

Use the M formula language and function reference for syntax, function parameters, table/list/record/text functions, and data source functions.

### Power Query Product Guidance

Use Power Query best practices, query folding, common authoring issues, multiple evaluations, and diagnostics pages for refresh behavior, performance, and troubleshooting.

### Excel VBA Host

Use Excel VBA documentation for `WorkbookQuery`, `Workbook.Queries`, `Queries.Add`, `WorkbookQuery.Formula`, `Workbook.RefreshAll`, `Application.CalculateUntilAsyncQueriesDone`, `QueryTable.Refresh`, and `QueryTable.BackgroundQuery`.

These pages define what can be automated through Excel COM and where VBA must wait, disable background refresh, or surface a prompt/error back to the user.
