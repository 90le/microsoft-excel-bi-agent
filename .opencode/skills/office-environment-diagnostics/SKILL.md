---
name: office-environment-diagnostics
description: Diagnose local Microsoft Office and Excel automation environments for workbook engineering. Use when Codex must check Excel COM, VBA project access, Power Query refresh readiness, ACE/OLEDB, MSOLAP, ADODB, ADOMD, Office bitness, provider drift, PowerShell/Git Bash wrapper behavior, or Linux/macOS structural-only limits before running workbook automation.
---

# Office Environment Diagnostics

## Core Rule

Diagnose the machine before blaming the workbook. Capture provider and runtime facts, then decide which workbook operations can be validated on the current system.

## Workflow

1. **Probe Office and providers**
   - Run `tools/probe_excel_bi_providers.ps1` on Windows.
   - Add `-RunExcelComSmoke` when Excel automation must be proven.
   - Add `-RunAdoWorkbookSmoke` when ACE/OLEDB workbook SQL must be proven.
   - From Git Bash, use `tools/invoke_excel_bi_com.sh provider-probe`.

2. **Build a reviewer-facing report**
   - Run `tools/build_provider_environment_report.py` against the probe JSON.
   - Use `--baseline-json` and `--fail-on-drift` when comparing machines or checking provider drift.
   - Use `tools/create_provider_environment_fixture.py` only to test report logic without live Office.

3. **Validate specialized endpoints**
   - For ADO workbook SQL, run `tools/test_excel_ado_sql_access.ps1`.
   - For ADOMD COM activation, run `tools/test_excel_adomd_query.ps1 -ProbeOnly`.
   - For a real cube/model endpoint, require an explicit connection string and MDX query.

4. **State platform boundaries**
   - Windows desktop Excel can validate Excel COM, VBA, Power Query refresh, Data Model, and CUBE recalculation.
   - Linux/macOS can run structural OpenXML and package validators, but cannot prove Excel runtime behavior without Excel.
   - Git Bash on Windows is a shell wrapper path; it still depends on Windows Excel and providers.

## Required Evidence

- Probe command and result path.
- Excel/Office bitness when available.
- Provider availability and smoke status.
- Explicit list of operations that are safe to run on the current machine.

## Boundaries

- A provider probe does not prove a customer workbook's business logic.
- A synthetic provider fixture proves report behavior only, not local Office readiness.
- Do not claim missing providers are package failures without separating environment capability from workbook defects.

## References

- Read `references/environment-decision-tree.md` when diagnosing Excel COM, provider, bitness, Trust Center, or refresh-readiness failures.
