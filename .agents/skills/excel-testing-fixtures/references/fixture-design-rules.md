# Fixture Design Rules

Use fixtures to prove generic tooling behavior without importing customer workbooks.

## Good Fixtures

- Small enough to inspect quickly.
- Deterministic across machines.
- Uses generic sheet names and generic measures.
- Exercises one clear behavior or risk category.
- Has an expected pass/fail result.
- Documents what it does not prove.

## Bad Fixtures

- Contains customer names, campaign names, local paths, credentials, or business rules.
- Requires a private data source.
- Depends on volatile dates or network availability.
- Mixes several unrelated risks so failures are hard to diagnose.

## Naming Pattern

```text
create_<surface>_fixture.py
build_<surface>_report.py
<surface> fixture smoke
```

## Evidence Boundary

A fixture can prove parser/report/release-gate behavior. It does not prove a specific customer workbook is correct.
