# Excel/VBA Workbook Workflow

## 1. Intake Checklist

- Confirm input file path and requested output path.
- Confirm the execution environment: Windows PowerShell, Windows Git Bash/MSYS/Cygwin, Linux, or macOS.
- Identify file type:
  - `.xlsx`: no VBA project in final output.
  - `.xlsm`: macro-enabled workbook.
  - `.xlsb`: binary macro-enabled workbook.
  - `.xls`: legacy format; prefer Excel COM for high-fidelity work.
- Confirm whether the workbook is open. If it is open in Excel, ask the user to close it when edits require saving the same file.
- Create a copy for edits unless the user explicitly requests in-place work.

## 2. Platform Selection

Choose the highest-fidelity path available:

| Environment | Use | Notes |
|---|---|---|
| Windows PowerShell + desktop Excel | `inspect_workbook.ps1`, `export_vba.ps1`, `import_vba.ps1` | Full VBA import/export and macro validation possible when VBProject access is trusted. |
| Windows Git Bash/MSYS/Cygwin + desktop Excel | `invoke_excel_com.sh` | Wrapper converts paths with `cygpath` when available and calls PowerShell. |
| Linux/macOS without desktop Excel | `inspect_openxml.py` plus workbook libraries | Good for structure/formula/data inspection; cannot compile or run VBA. |
| Linux/macOS with LibreOffice | LibreOffice only for compatible workbook operations | Do not treat LibreOffice macro behavior as equivalent to Excel VBA validation. |

Git Bash examples:

```bash
scripts/invoke_excel_com.sh inspect -w "/c/path/workbook.xlsm" -o "/c/path/inspect.json"
scripts/invoke_excel_com.sh export -w "/c/path/workbook.xlsm" -d "/c/path/src/vba"
scripts/invoke_excel_com.sh import -w "/c/path/source.xlsm" -s "/c/path/src/vba" -o "/c/path/output.xlsm"
```

Linux/macOS structural inspection:

```bash
python3 scripts/inspect_openxml.py "/path/workbook.xlsm" --out-json "/path/inspect.json"
```

VBA source lint before import:

```bash
python3 scripts/lint_vba_source.py "/path/src/vba" --strict-option-explicit --out-json "/path/tmp/vba-lint.json"
```

This static check can catch unclosed procedures, duplicate public standard-module entry macros, and missing `Option Explicit`. It does not compile VBA, resolve references, or prove macros run in Excel.

## 3. Inspection Checklist

Inspect these before editing:

- Worksheet names, tab order, visibility, used range, frozen panes, and protection.
- Named ranges and formulas in key output cells.
- External links, workbook connections, Power Query queries, and data model dependencies.
- Shapes, form controls, ActiveX controls, and `OnAction` macro names.
- VBA components, public entry macros, event handlers, and references.
- Hidden or very hidden support sheets.

Use the script:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/inspect_workbook.ps1 `
  -WorkbookPath "C:\path\workbook.xlsm" `
  -OutJson "C:\path\inspect.json"
```

If Excel COM is unavailable, run:

```bash
python3 scripts/inspect_openxml.py "/path/workbook.xlsm" --out-json "/path/inspect.json"
```

OpenXML inspection can identify sheets, defined names, formulas, links, connections, drawings, and `vbaProject.bin` presence, but it cannot inspect live VBE modules, compile VBA, run macros, refresh Power Query, or execute Solver.

## 4. Editing Strategy

Prefer the least invasive path that can be validated:

- For data/formula/format-only changes, edit the workbook with structured spreadsheet tools or Excel COM.
- For VBA changes, export modules, edit source files, import into a workbook copy, and save.
- Run `lint_vba_source.py` on edited `.bas`, `.cls`, and `.frm` files before import when source files are available.
- On Linux/macOS, keep `.xlsm` macros intact by preserving `vbaProject.bin` when using workbook libraries. Do not promise newly imported VBA unless a Windows Excel validation step is available.
- For event code in `ThisWorkbook` or sheet modules, update existing document modules rather than deleting them.
- For generated UI, prefer simple shapes with `OnAction` over ActiveX.
- For deliverables, remove scratch ranges or hide calculation sheets only after verification.

## 5. Validation Checklist

Minimum validation for a workbook deliverable:

- Reopen the output workbook successfully.
- Confirm edited VBA source passed `lint_vba_source.py` before import when source files were changed.
- Run one or more public macros that represent the workflow.
- Verify key inputs and outputs by cell address.
- Verify totals and constraints where applicable.
- Scan for formula errors: `#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!`, `#N/A`.
- Confirm macro buttons still target existing public macros.
- Confirm hidden/deleted sheets match the intended delivery shape.
- Confirm output file type is correct (`.xlsm` when macros exist).
- If the environment is Linux/macOS, explicitly mark macro execution, VBE compile, Solver, Power Query refresh, and button click behavior as not validated unless they were tested in Excel.

For complex work, create a short validation document with:

```text
Workbook: <path>
Inputs tested: <cells/values>
Macros run: <names>
Expected/actual key outputs: <cell/value pairs>
Known limitations: <if any>
```

## 6. Troubleshooting

### VBProject access denied

Automated VBA import/export requires this Excel setting:

```text
File > Options > Trust Center > Trust Center Settings >
Macro Settings > Trust access to the VBA project object model
```

If it is disabled, the COM scripts can still inspect sheets, formulas, links, and shapes, but not VBA modules.

### Workbook locked

If saving fails because the workbook is open:

- Do not overwrite or kill Excel processes blindly.
- Save to a new output filename.
- Ask the user to close the workbook if the exact file must be modified.

### Git Bash path problems

If PowerShell cannot find a file passed from Git Bash:

- Use `scripts/invoke_excel_com.sh` instead of calling `.ps1` scripts directly.
- Convert paths with `cygpath -w` manually if needed.
- Quote all paths; Excel files often contain spaces or non-ASCII characters.
- Avoid mixed path styles inside the same command.

### Linux/macOS limitation

If a task requires VBE import/export, Solver, Power Query refresh, cube formulas, ActiveX, or macro execution:

- State that a Windows desktop Excel validation step is required.
- Continue with OpenXML inspection, source-code drafting, workbook shape edits, or documentation if useful.
- Do not present structural inspection as proof that macros run correctly in Excel.

### Macro button does nothing

Check:

- `Shape.OnAction` points to a public macro name.
- Macro is in a standard module, not a private sheet module.
- Macro security allows macros.
- The workbook was saved as `.xlsm` or `.xlsb`.

### Compile error after import

Common causes:

- `Attribute VB_Name` lines were pasted into the VBE code pane.
- `Option Explicit` exposed undeclared variables.
- Missing object library references; switch to late binding if possible.
- Non-ASCII text was corrupted during import.
- A module, procedure, or variable name conflicts with an existing name.

### Formula or result mismatch

Check:

- Cell units and scale, especially yuan vs wan, percentages vs decimals.
- Absolute/relative references after fill-right/fill-down.
- Hidden helper cells or named ranges still pointing to old sheets.
- Manual calculation mode; force `Application.CalculateFullRebuild` when necessary.
- External links, Power Query refresh, or cube formulas that may not update offline.

## 7. Delivery Checklist

- The final workbook is in the requested format.
- Source workbook remains intact unless in-place modification was explicitly requested.
- VBA source files are retained for future maintenance.
- Validation notes identify what was tested.
- Validation notes identify the platform used and any platform-specific limitations.
- Temporary scratch files are not presented as final deliverables unless useful.
