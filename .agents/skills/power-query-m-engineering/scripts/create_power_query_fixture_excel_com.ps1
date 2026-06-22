param(
    [Parameter(Mandatory = $true)]
    [string]$OutputWorkbookPath,

    [string]$QueryName = "SmokeQuery",

    [string]$TableName = "SmokeQueryTable",

    [string]$FormulaPath,

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

function Read-MFormula {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return 'let Source = #table({"A","B"}, {{1,"x"},{2,"y"}}) in Source'
    }
    return Get-Content -LiteralPath (Resolve-Path -LiteralPath $Path).Path -Raw -Encoding UTF8
}

$resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputWorkbookPath)
$outputParent = Split-Path -Parent $resolvedOutput
if ($outputParent) { New-Item -ItemType Directory -Path $outputParent -Force | Out-Null }
if (Test-Path -LiteralPath $resolvedOutput) { Remove-Item -LiteralPath $resolvedOutput -Force }

$excel = $null
$workbook = $null
$listObject = $null
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$startedAt = Get-Date
$stopwatch = [Diagnostics.Stopwatch]::StartNew()

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "PQ_Load"

    $mFormula = Read-MFormula $FormulaPath
    [void]$workbook.Queries.Add($QueryName, $mFormula)

    $connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$QueryName;Extended Properties=`"`""
    $source = [object[]]@($connectionString)
    $listObject = $sheet.ListObjects.Add(0, $source, $true, 1, $sheet.Range("A1"))
    $listObject.Name = $TableName
    try { $listObject.DisplayName = $TableName } catch { }
    $listObject.QueryTable.CommandType = 2
    $listObject.QueryTable.CommandText = [object[]]@("SELECT * FROM [$QueryName]")
    $listObject.QueryTable.BackgroundQuery = $false
    [void]$listObject.QueryTable.Refresh($false)

    foreach ($connection in @($workbook.Connections)) {
        try {
            if ([string]$connection.OLEDBConnection.Connection -like "*Location=$QueryName*") {
                $connection.Name = "Query - $QueryName"
                $connection.Description = "Connection to the '$QueryName' query in the workbook."
            }
        } catch {
            # Some connection types do not expose OLEDBConnection.
        }
    }

    $workbook.SaveAs($resolvedOutput, (FileFormat-ForPath $resolvedOutput))
    $stopwatch.Stop()

    $result = [ordered]@{
        outputWorkbookPath = $resolvedOutput
        queryName = $QueryName
        tableName = $TableName
        formulaPath = $FormulaPath
        rows = if ($null -ne $listObject.DataBodyRange) { [int]$listObject.DataBodyRange.Rows.Count } else { 0 }
        columns = [int]$listObject.Range.Columns.Count
        connectionString = $connectionString
        startedAt = $startedAt.ToString("o")
        completedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
    }

    $json = $result | ConvertTo-Json -Depth 10
    if ($OutJson) {
        $outParent = Split-Path -Parent $OutJson
        if ($outParent) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
        Set-Content -LiteralPath $OutJson -Value $json -Encoding UTF8
    } else {
        $json
    }
} catch {
    $stopwatch.Stop()
    $failure = [ordered]@{
        outputWorkbookPath = $resolvedOutput
        queryName = $QueryName
        failedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        error = $_.Exception.Message
    }
    $json = $failure | ConvertTo-Json -Depth 10
    if ($OutJson) {
        $outParent = Split-Path -Parent $OutJson
        if ($outParent) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
        Set-Content -LiteralPath $OutJson -Value $json -Encoding UTF8
    } else {
        $json
    }
    throw
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
