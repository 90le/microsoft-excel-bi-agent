param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$OutJson
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Get-ExcelProcessId {
    param([object]$ExcelApplication)
    try {
        if (-not ("ExcelVbaWorkbookEngineering.NativeMethods" -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace ExcelVbaWorkbookEngineering {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
"@
        }
        $processId = 0
        [void][ExcelVbaWorkbookEngineering.NativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
        return [int]$processId
    } catch {
        return $null
    }
}

function Stop-OwnExcelIfStillRunning {
    param(
        [object]$ProcessId,
        [int[]]$PreExistingProcessIds
    )
    $candidateIds = @()
    if ($null -ne $ProcessId -and [int]$ProcessId -gt 0) {
        $candidateIds += [int]$ProcessId
    }
    foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if (($PreExistingProcessIds -notcontains [int]$process.Id) -and [string]::IsNullOrEmpty($process.MainWindowTitle)) {
            $candidateIds += [int]$process.Id
        }
    }
    foreach ($id in @($candidateIds | Select-Object -Unique)) {
        $process = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -ne $process) {
            Stop-Process -Id $id -Force
        }
    }
}

function Try-Read {
    param(
        [scriptblock]$Script,
        [object]$Default = $null
    )
    try {
        & $Script
    } catch {
        $Default
    }
}

function Sheet-VisibilityName {
    param([int]$VisibleValue)
    switch ($VisibleValue) {
        -1 { "Visible" }
        0 { "Hidden" }
        2 { "VeryHidden" }
        default { "Unknown:$VisibleValue" }
    }
}

$resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$excel = $null
$workbook = $null
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })

$xlCellTypeFormulas = -4123
$xlExcelLinks = 1
$xlOLELinks = 2

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($resolvedPath, $null, $true)

    $sheets = @()
    foreach ($sheet in @($workbook.Worksheets)) {
        $used = Try-Read { $sheet.UsedRange }
        $formulaCount = 0
        if ($null -ne $used) {
            $formulaCount = Try-Read { $used.SpecialCells($xlCellTypeFormulas).Count } 0
        }

        $shapes = @()
        foreach ($shape in @($sheet.Shapes)) {
            $shapes += [ordered]@{
                name     = Try-Read { [string]$shape.Name } ""
                type     = Try-Read { [int]$shape.Type } $null
                onAction = Try-Read { [string]$shape.OnAction } ""
                text     = Try-Read { [string]$shape.TextFrame2.TextRange.Text } ""
            }
        }

        $sheets += [ordered]@{
            name         = [string]$sheet.Name
            codeName     = Try-Read { [string]$sheet.CodeName } ""
            visible      = Sheet-VisibilityName ([int]$sheet.Visible)
            usedAddress  = Try-Read { [string]$used.Address($false, $false) } ""
            rows         = Try-Read { [int]$used.Rows.Count } 0
            columns      = Try-Read { [int]$used.Columns.Count } 0
            formulaCount = [int]$formulaCount
            shapes       = $shapes
        }
    }

    $names = @()
    foreach ($name in @($workbook.Names)) {
        $names += [ordered]@{
            name       = Try-Read { [string]$name.Name } ""
            refersTo   = Try-Read { [string]$name.RefersTo } ""
            visible    = Try-Read { [bool]$name.Visible } $null
        }
    }

    $connections = @()
    foreach ($connection in @($workbook.Connections)) {
        $connections += [ordered]@{
            name        = Try-Read { [string]$connection.Name } ""
            description = Try-Read { [string]$connection.Description } ""
            type        = Try-Read { [int]$connection.Type } $null
        }
    }

    $queries = @()
    $queryAccessError = $null
    try {
        foreach ($query in @($workbook.Queries)) {
            $queries += [ordered]@{
                name    = Try-Read { [string]$query.Name } ""
                formula = Try-Read { [string]$query.Formula } ""
            }
        }
    } catch {
        $queryAccessError = $_.Exception.Message
    }

    $links = @()
    foreach ($linkType in @($xlExcelLinks, $xlOLELinks)) {
        $linkSources = Try-Read { $workbook.LinkSources($linkType) } $null
        if ($null -ne $linkSources) {
            foreach ($link in @($linkSources)) {
                $links += [ordered]@{
                    type = $linkType
                    path = [string]$link
                }
            }
        }
    }

    $vbaComponents = @()
    $vbaAccessError = $null
    try {
        foreach ($component in @($workbook.VBProject.VBComponents)) {
            $vbaComponents += [ordered]@{
                name        = [string]$component.Name
                type        = [int]$component.Type
                lineCount   = Try-Read { [int]$component.CodeModule.CountOfLines } 0
            }
        }
    } catch {
        $vbaAccessError = $_.Exception.Message
    }

    $result = [ordered]@{
        workbookPath     = $resolvedPath
        name             = [string]$workbook.Name
        fileFormat       = Try-Read { [int]$workbook.FileFormat } $null
        hasVBProject     = Try-Read { [bool]$workbook.HasVBProject } $null
        worksheets       = $sheets
        names            = $names
        links            = $links
        connections      = $connections
        queries          = $queries
        queryAccessError = $queryAccessError
        vbaComponents    = $vbaComponents
        vbaAccessError   = $vbaAccessError
    }

    $json = $result | ConvertTo-Json -Depth 12
    if ($OutJson) {
        $outParent = Split-Path -Parent $OutJson
        if ($outParent) {
            New-Item -ItemType Directory -Path $outParent -Force | Out-Null
        }
        Set-Content -LiteralPath $OutJson -Value $json -Encoding UTF8
    } else {
        $json
    }
} finally {
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
