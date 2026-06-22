param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$OutDir,

    [Parameter(Mandatory = $true)]
    [string]$OutJson,

    [string]$OutMd = "",

    [switch]$CreateFixture
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    try {
        if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    } catch {
        # COM release errors should not invalidate the evidence report; the
        # finally block also stops the hidden Excel process started here.
    }
}

function Get-ExcelProcessId {
    param([object]$ExcelApplication)
    try {
        if (-not ("ExcelBiVisualQa.NativeMethods" -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace ExcelBiVisualQa {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
"@
        }
        $processId = 0
        [void][ExcelBiVisualQa.NativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
        return [int]$processId
    } catch {
        return $null
    }
}

function Stop-OwnExcelIfStillRunning {
    param([object]$ProcessId, [int[]]$PreExistingProcessIds)
    $candidateIds = @()
    if ($null -ne $ProcessId -and [int]$ProcessId -gt 0) { $candidateIds += [int]$ProcessId }
    foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if (($PreExistingProcessIds -notcontains [int]$process.Id) -and [string]::IsNullOrEmpty($process.MainWindowTitle)) {
            $candidateIds += [int]$process.Id
        }
    }
    foreach ($id in @($candidateIds | Select-Object -Unique)) {
        $process = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -ne $process) { Stop-Process -Id $id -Force }
    }
}

function Safe-FileName {
    param([string]$Value)
    $safe = $Value -replace '[\\/:*?"<>|]', '_'
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "sheet"
    }
    return $safe
}

function Write-MarkdownReport {
    param(
        [string]$Path,
        [object]$Report
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# Visual QA Render Evidence")
    $lines.Add("")
    $lines.Add("- source: Windows Excel COM PDF export")
    $lines.Add("- readiness: $($Report.summary.readiness)")
    $lines.Add("- workbook: $($Report.workbook.name)")
    $lines.Add("- exported sheets: $($Report.summary.exportedSheetCount)")
    $lines.Add("- exported files: $($Report.summary.exportedFileCount)")
    $lines.Add("- total bytes: $($Report.summary.totalBytes)")
    $lines.Add("")
    $lines.Add("| Sheet | File | Bytes | Print Area | Used Rows | Used Columns |")
    $lines.Add("|---|---|---:|---|---:|---:|")
    foreach ($item in $Report.exports) {
        $printArea = [string]$item.printArea
        if ([string]::IsNullOrWhiteSpace($printArea)) {
            $printArea = "(none)"
        }
        $lines.Add("| $($item.sheetName) | $($item.fileName) | $($item.bytes) | $printArea | $($item.usedRows) | $($item.usedColumns) |")
    }
    $lines.Add("")
    $lines.Add("## Boundaries")
    foreach ($boundary in $Report.boundaries) {
        $lines.Add("- $boundary")
    }
    Set-Content -LiteralPath $Path -Value ($lines -join [Environment]::NewLine) -Encoding UTF8
}

function Add-ReportSheet {
    param(
        [object]$Workbook,
        [string]$Name,
        [bool]$Blank,
        [bool]$Clipped
    )
    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $Name
    $sheet.Columns.Item(1).ColumnWidth = 18
    $sheet.Columns.Item(2).ColumnWidth = $(if ($Clipped) { 8 } else { 16 })
    $sheet.Columns.Item(3).ColumnWidth = 16
    $sheet.Columns.Item(4).ColumnWidth = 16
    $sheet.Columns.Item(5).ColumnWidth = 16
    $sheet.Columns.Item(6).ColumnWidth = 16

    if (-not $Blank) {
        $sheet.Range("A1").Value2 = "Visual QA Report Surface"
        $sheet.Range("A3").Value2 = "Metric"
        $sheet.Range("B3").Value2 = "Value"
        $sheet.Range("C3").Value2 = "Target"
        $sheet.Range("A4").Value2 = "Reach"
        $sheet.Range("B4").Value2 = 128
        $sheet.Range("C4").Value2 = 120
        $sheet.Range("A5").Value2 = "Awareness"
        $sheet.Range("B5").Value2 = 84
        $sheet.Range("C5").Value2 = 80
        $sheet.Range("A6").Value2 = "Conversion"
        $sheet.Range("B6").Value2 = 42
        $sheet.Range("C6").Value2 = 40
        $sheet.Range("A8").Value2 = "Status"
        $sheet.Range("B8").Value2 = $(if ($Clipped) { "This label is intentionally too long for the narrow column and should be visible in render evidence review." } else { "Ready" })
        $sheet.Range("A10").Value2 = "Total"
        $sheet.Range("B10").Formula = "=SUM(B4:B6)"
        $sheet.Range("A1:F1").Merge() | Out-Null
        $sheet.Range("A1").Font.Bold = $true
        $sheet.Range("A1").Font.Size = 16
        $sheet.Range("A3:C3").Font.Bold = $true
        $sheet.Range("A3:C10").Borders.LineStyle = 1
        if (-not $Clipped) {
            $chart = $sheet.Shapes.AddChart2(201, 51, 260, 70, 280, 160).Chart
            $chart.SetSourceData($sheet.Range("A4:B6"))
            $chart.HasTitle = $true
            $chart.ChartTitle.Text = "Report Metrics"
            Release-ComObject $chart
        }
        $sheet.PageSetup.PrintArea = '$A$1:$F$14'
    }

    $sheet.Activate() | Out-Null
    $sheet.Range("A4").Select() | Out-Null
    $Workbook.Application.ActiveWindow.FreezePanes = $true
    return $sheet
}

function New-VisualRenderFixtureWorkbook {
    param(
        [object]$Excel,
        [string]$Path
    )
    $parent = Split-Path -Parent $Path
    if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $workbook = $Excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }
    $dataSheet = $workbook.Worksheets.Item(1)
    $dataSheet.Name = "Data"
    $dataSheet.Range("A1").Value2 = "Channel"
    $dataSheet.Range("B1").Value2 = "Amount"
    $dataSheet.Range("C1").Value2 = "Score"
    $dataSheet.Range("A2").Value2 = "A"
    $dataSheet.Range("B2").Value2 = 100
    $dataSheet.Range("C2").Value2 = 1.2
    $dataSheet.Range("A3").Value2 = "B"
    $dataSheet.Range("B3").Value2 = 80
    $dataSheet.Range("C3").Value2 = 0.9

    $blank = Add-ReportSheet -Workbook $workbook -Name "Report_Blank" -Blank $true -Clipped $false
    $clipped = Add-ReportSheet -Workbook $workbook -Name "Report_Clipped" -Blank $false -Clipped $true
    $ok = Add-ReportSheet -Workbook $workbook -Name "Report_OK" -Blank $false -Clipped $false
    $dataSheet.Visible = 0
    $ok.Activate() | Out-Null
    $workbook.SaveAs($Path, 51)

    Release-ComObject $blank
    Release-ComObject $clipped
    Release-ComObject $ok
    Release-ComObject $dataSheet
    return $workbook
}

$resolvedWorkbook = [System.IO.Path]::GetFullPath($WorkbookPath)
if (-not $CreateFixture) {
    $resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
}
$resolvedOutDir = [System.IO.Path]::GetFullPath($OutDir)
$resolvedOutJson = [System.IO.Path]::GetFullPath($OutJson)
$resolvedOutMd = ""
if (-not [string]::IsNullOrWhiteSpace($OutMd)) {
    $resolvedOutMd = [System.IO.Path]::GetFullPath($OutMd)
}

New-Item -ItemType Directory -Path $resolvedOutDir -Force | Out-Null
New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($resolvedOutJson)) -Force | Out-Null
if ($resolvedOutMd) {
    New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($resolvedOutMd)) -Force | Out-Null
}

$excel = $null
$workbook = $null
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$exports = New-Object System.Collections.Generic.List[object]

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    if ($CreateFixture) {
        $workbook = New-VisualRenderFixtureWorkbook -Excel $excel -Path $resolvedWorkbook
    } else {
        $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $true)
    }
    $workbook.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()
    $excel.CalculateFull()

    foreach ($worksheet in $workbook.Worksheets) {
        $sheetName = [string]$worksheet.Name
        if ($worksheet.Visible -ne -1 -or $sheetName -notlike "Report*") {
            Release-ComObject $worksheet
            continue
        }

        $safeName = Safe-FileName $sheetName
        $pdfPath = Join-Path $resolvedOutDir "$safeName.pdf"
        $usedRange = $worksheet.UsedRange
        $usedRows = [int]$usedRange.Rows.Count
        $usedColumns = [int]$usedRange.Columns.Count
        $printArea = [string]$worksheet.PageSetup.PrintArea

        $worksheet.ExportAsFixedFormat(0, $pdfPath)
        $fileInfo = Get-Item -LiteralPath $pdfPath
        $exports.Add([ordered]@{
            sheetName = $sheetName
            fileName = $fileInfo.Name
            path = $fileInfo.FullName
            exists = $true
            bytes = [int64]$fileInfo.Length
            usedRows = $usedRows
            usedColumns = $usedColumns
            printArea = $printArea
        })

        Release-ComObject $usedRange
        Release-ComObject $worksheet
    }

    $failures = New-Object System.Collections.Generic.List[string]
    if ($exports.Count -lt 1) {
        $failures.Add("no Report* sheets were exported")
    }
    foreach ($item in $exports) {
        if (-not $item.exists -or $item.bytes -le 0) {
            $failures.Add("empty export: $($item.sheetName)")
        }
    }

    $totalBytes = 0
    foreach ($item in $exports) {
        $totalBytes += [int64]$item.bytes
    }
    $failureCount = [int]$failures.Count
    $readiness = "rendered"
    if ($failureCount -ne 0) {
        $readiness = "render-failed"
    }
    $exportItems = @($exports.ToArray())
    $failureItems = @($failures.ToArray())

    $report = [ordered]@{
        generatedAt = (Get-Date).ToUniversalTime().ToString("o")
        source = "Windows Excel COM PDF export"
        workbook = [ordered]@{
            name = [System.IO.Path]::GetFileName($resolvedWorkbook)
            createdFixture = [bool]$CreateFixture
            sheetCount = [int]$workbook.Worksheets.Count
        }
        output = [ordered]@{
            directory = $resolvedOutDir
            json = $resolvedOutJson
            markdown = $resolvedOutMd
        }
        summary = [ordered]@{
            readiness = $readiness
            exportedSheetCount = [int]$exports.Count
            exportedFileCount = [int]$exports.Count
            totalBytes = [int64]$totalBytes
            failureCount = $failureCount
        }
        exports = $exportItems
        failures = $failureItems
        boundaries = @(
            "This proves Excel could open the workbook and export visible Report* sheets to PDF on this Windows machine.",
            "It does not compare pixels, judge design quality, or prove a customer workbook is correct.",
            "Generated PDFs and reports are task-local evidence and must not be committed into the plugin package.",
            "Linux/macOS structural checks cannot claim this Excel COM rendering evidence."
        )
    }

    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedOutJson -Encoding UTF8
    if ($resolvedOutMd) {
        Write-MarkdownReport -Path $resolvedOutMd -Report $report
    }

    if ($failures.Count -gt 0) {
        Write-Error ("Visual QA render evidence failed: " + ($failures -join "; "))
    }

    Write-Host ("Visual QA render evidence rendered: sheets={0}, files={1}, bytes={2}" -f $exports.Count, $exports.Count, $totalBytes)
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnExcelIfStillRunning $excelProcessId $preExistingExcelProcessIds
}
