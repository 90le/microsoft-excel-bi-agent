# Excel Compatibility Core Design

## Objective

Upgrade Microsoft Excel BI Agent from a collection of specialist Excel skills into a capability-aware package that can distinguish structural evidence, local Office runtime evidence, and workbook-specific runtime proof across Windows, macOS, web, legacy Office, LTSC, and Microsoft 365 environments.

## Product boundary

This release keeps all published skill identifiers stable. `office-environment-diagnostics` becomes the compatibility authority without being renamed. It reports evidence; it does not claim that a workbook's business logic is correct. `excel-bi-router` uses platform and compatibility intent to route to that authority while preserving specialist routing for DAX, Power Query, VBA, MDX, and ADO requests.

## Architecture

The compatibility flow has three layers:

1. `probe_excel_capabilities.ps1` gathers Windows runtime evidence and reuses the existing provider probe.
2. `build_excel_compatibility_report.py` validates a probe contract, maps stable capability IDs to operations, and applies explicit user requirements.
3. Skills, task profiles, and documentation explain what was tested, what remains unknown, and which specialist runtime test is still required.

The structural release gate validates synthetic fixtures on every platform. The full Windows gate additionally exercises the live probe. Missing Excel or a provider is a compatibility result; malformed evidence, leaked Excel processes, or an invalid contract is a package failure.

## Stable capability IDs

- `excel.com.activation`
- `excel.workbook.roundtrip`
- `excel.vba.project-access`
- `excel.power-query.object-model`
- `excel.power-query.async-wait`
- `excel.data-model.object-model`
- `excel.pdf-export`
- `ado.com.activation`
- `ace.workbook-sql`
- `msolap.registration`
- `adomd.com.activation`

## Release hygiene

The same release fixes Windows subprocess decoding, prefers Git Bash over WSL `bash.exe` for Windows path handling, completes plugin interface URLs, reduces starter prompts from nine to three, and adds an allowlisted runtime-package builder. The repository remains the cross-agent source distribution; the Codex runtime package contains only the manifest, generated `skills/`, required tools/fixtures, license, and compact readme.

## Compatibility policy

- Active Microsoft 365/LTSC environments can receive full capability probes.
- Excel 2016/2019 are legacy-compatible and must receive lifecycle warnings, not forced upgrade instructions.
- Excel 2010/2013 are best-effort and may require legacy add-ins or fallback paths.
- macOS, web, Linux, WPS, and LibreOffice receive only the evidence level actually available.
- Offline and cloud-prohibited environments are first-class policy profiles.
- No structural check may be reported as proof of VBA execution, Power Query refresh, Data Model calculation, or business correctness.

## Success criteria

- Windows Chinese-path structural release gate no longer crashes while decoding subprocess output.
- The runtime package has an allowlist, manifest, SHA-256 file list, and no cross-agent mirror duplication.
- Exactly three starter prompts cover inspect, diagnose, and publish.
- Compatibility fixtures cover all-pass, Excel-blocked, partial evidence, and malformed evidence.
- Router fixtures distinguish platform compatibility requests from DAX compatibility requests.
- Source skills and all generated mirrors are synchronized.
- Structural validation passes before the branch is published.

