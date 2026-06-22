param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [ValidateSet("List", "Add", "Update", "Delete")]
    [string]$Action,

    [string]$QueryName,

    [string]$FormulaPath,

    [string]$Description = "",

    [string]$OutputWorkbookPath,

    [string]$OutJson,

    [switch]$InPlace
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
        if (-not ("ExcelBiAgentPack.NativeMethods" -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace ExcelBiAgentPack {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
"@
        }
        $processId = 0
        [void][ExcelBiAgentPack.NativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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

function Try-Read {
    param([scriptblock]$Script, [object]$Default = $null)
    try { & $Script } catch { $Default }
}

function FileFormat-ForPath {
    param([string]$Path)
    switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        ".xlsx" { 51 }
        ".xlsm" { 52 }
        ".xlsb" { 50 }
        ".xls"  { 56 }
        default { throw "Unsupported output extension. Use .xlsx, .xlsm, .xlsb, or .xls." }
    }
}

function Get-QueriesSnapshot {
    param([object]$Workbook)
    $items = @()
    foreach ($query in @($Workbook.Queries)) {
        $items += [ordered]@{
            name        = [string]$query.Name
            description = Try-Read { [string]$query.Description } ""
            formula     = Try-Read { [string]$query.Formula } ""
        }
    }
    return $items
}

if ($Action -ne "List") {
    if (-not $InPlace -and [string]::IsNullOrWhiteSpace($OutputWorkbookPath)) {
        throw "Add/Update/Delete require -OutputWorkbookPath unless -InPlace is explicitly supplied."
    }
    if ([string]::IsNullOrWhiteSpace($QueryName)) {
        throw "$Action requires -QueryName."
    }
}

if (($Action -eq "Add" -or $Action -eq "Update") -and [string]::IsNullOrWhiteSpace($FormulaPath)) {
    throw "$Action requires -FormulaPath."
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedFormula = $null
if ($FormulaPath) { $resolvedFormula = (Resolve-Path -LiteralPath $FormulaPath).Path }
$resolvedOutput = $null
if ($OutputWorkbookPath) { $resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputWorkbookPath) }

$excel = $null
$workbook = $null
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$startedAt = Get-Date

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $readOnly = ($Action -eq "List")
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $readOnly)

    $before = Get-QueriesSnapshot $workbook
    $message = ""

    switch ($Action) {
        "List" {
            $message = "Listed workbook queries."
        }
        "Add" {
            foreach ($query in @($workbook.Queries)) {
                if ([string]$query.Name -eq $QueryName) { throw "Query already exists: $QueryName" }
            }
            $formula = Get-Content -Raw -LiteralPath $resolvedFormula
            $workbook.Queries.Add($QueryName, $formula, $Description) | Out-Null
            $message = "Added query: $QueryName"
        }
        "Update" {
            $target = $null
            foreach ($query in @($workbook.Queries)) {
                if ([string]$query.Name -eq $QueryName) { $target = $query; break }
            }
            if ($null -eq $target) { throw "Query not found: $QueryName" }
            $target.Formula = Get-Content -Raw -LiteralPath $resolvedFormula
            if ($PSBoundParameters.ContainsKey("Description")) {
                $target.Description = $Description
            }
            $message = "Updated query: $QueryName"
        }
        "Delete" {
            $target = $null
            foreach ($query in @($workbook.Queries)) {
                if ([string]$query.Name -eq $QueryName) { $target = $query; break }
            }
            if ($null -eq $target) { throw "Query not found: $QueryName" }
            $target.Delete()
            $message = "Deleted query: $QueryName"
        }
    }

    if ($Action -ne "List") {
        if ($InPlace) {
            $workbook.Save()
        } else {
            $parent = Split-Path -Parent $resolvedOutput
            if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
            $workbook.SaveAs($resolvedOutput, (FileFormat-ForPath $resolvedOutput))
        }
    }

    $after = Get-QueriesSnapshot $workbook
    $result = [ordered]@{
        workbookPath       = $resolvedWorkbook
        action             = $Action
        queryName          = $QueryName
        outputWorkbookPath = $resolvedOutput
        inPlace            = [bool]$InPlace
        startedAt          = $startedAt.ToString("o")
        completedAt        = (Get-Date).ToString("o")
        message            = $message
        before             = $before
        after              = $after
    }

    $json = $result | ConvertTo-Json -Depth 10
    if ($OutJson) {
        $outParent = Split-Path -Parent $OutJson
        if ($outParent) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
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
