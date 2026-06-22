# VBA Patterns For Workbook Engineering

## Standard Module Header

Use exported source files for durable VBA edits:

```vb
Attribute VB_Name = "modExample"
Option Explicit

Public Sub RunExample()
    ' Entry point used by buttons or Application.Run.
End Sub
```

When pasting directly into the VBE editor, omit the `Attribute VB_Name` line.

## Macro Entry Points

Button macros should be public procedures in a standard module:

```vb
Option Explicit

Public Sub RunReport()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' Work here.

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "Report did not complete: " & Err.Description, vbExclamation, "Workbook"
    Resume CleanExit
End Sub
```

## Event Guard Pattern

Use a module-level guard to prevent recursive `Worksheet_Change` or button-triggered recalculation:

```vb
Option Explicit

Private mRunning As Boolean

Public Sub RunCalculation()
    If mRunning Then Exit Sub
    On Error GoTo Fail

    mRunning = True
    Application.EnableEvents = False

    ' Work here.

CleanExit:
    Application.EnableEvents = True
    mRunning = False
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

## Safe Application State

Always restore Excel state:

```vb
Dim oldEvents As Boolean
Dim oldScreenUpdating As Boolean
Dim oldCalculation As XlCalculation

oldEvents = Application.EnableEvents
oldScreenUpdating = Application.ScreenUpdating
oldCalculation = Application.Calculation

On Error GoTo Fail
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Work here.

CleanExit:
Application.Calculation = oldCalculation
Application.ScreenUpdating = oldScreenUpdating
Application.EnableEvents = oldEvents
Exit Sub

Fail:
MsgBox Err.Description, vbExclamation
Resume CleanExit
```

## Late Binding Pattern

Use late binding when a workbook may run on machines without the same references:

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

Avoid requiring users to manually add references unless the workbook already depends on them.

## Cross-Platform VBA Pattern

If the workbook may run on Mac Excel as well as Windows Excel, avoid Windows-only APIs unless guarded:

```vb
#If Mac Then
    ' Mac-compatible path or behavior.
#Else
    ' Windows-specific API, COM, or file picker behavior.
#End If
```

Prefer `Application.PathSeparator`, workbook-relative paths, and late binding. Avoid hard-coded `C:\...` paths inside workbook VBA unless the workbook is explicitly Windows-only.

## Shape Button Binding

Create or update a simple shape button:

```vb
Dim shp As Shape
Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 300, 40, 120, 32)
shp.TextFrame2.TextRange.Text = "Run"
shp.OnAction = "RunReport"
```

Validation: read `Shape.OnAction` and confirm the macro exists.

## Solver Automation Notes

If Solver is required:

- Prefer `Application.Run "Solver.xlam!SolverReset"` and related calls instead of compile-time Solver references.
- Verify the Solver add-in exists and is installed.
- Always validate returned results against constraints after Solver runs.
- Do not trust a Solver status code alone; verify totals, bounds, and locked variables by cell value.

For workbooks where Solver creates unstable or hard-to-explain allocations, consider implementing a deterministic VBA algorithm and using Solver only as a benchmark.

## Workbook Cleanup Patterns

Convert formulas to values for a clean delivery copy:

```vb
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.UsedRange.Value = ws.UsedRange.Value
Next ws
```

Remove links and queries carefully in a copy:

```vb
Dim c As WorkbookConnection
For Each c In ThisWorkbook.Connections
    c.Delete
Next c
```

Power Query and data model cleanup can be workbook-specific. Inspect `Workbook.Queries`, `Workbook.Connections`, and links before deleting anything.

## Common VBA Errors

- **Variable not defined**: Declare the variable, fix spelling, or replace missing constants with local constants.
- **Sub or Function not defined**: Macro name mismatch, private procedure, missing module, or missing add-in.
- **Object required**: A worksheet/range/shape lookup returned `Nothing`.
- **Application-defined or object-defined error**: Usually bad range address, protected sheet, invalid shape/control, or workbook state.
- **Type mismatch**: Cell contains text/error/blank where numeric value was expected; validate inputs before calculation.

## Numeric Validation Pattern

Use explicit input validation and units:

```vb
Private Function ReadPositiveNumber(ByVal value As Variant, ByRef result As Double, ByVal label As String) As Boolean
    If IsError(value) Or Not IsNumeric(value) Then
        MsgBox label & " must be numeric.", vbExclamation
        Exit Function
    End If
    result = CDbl(value)
    If result <= 0 Then
        MsgBox label & " must be greater than zero.", vbExclamation
        Exit Function
    End If
    ReadPositiveNumber = True
End Function
```

For budget tools, write units into labels and variable names, such as `budgetWan`, `budgetYuan`, `unitWan`, or `unitYuan`.
