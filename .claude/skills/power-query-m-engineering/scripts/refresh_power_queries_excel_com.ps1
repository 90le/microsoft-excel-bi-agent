param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$QueryName,

    [string]$OutputWorkbookPath,

    [string]$OutJson,

    [int]$TimeoutSeconds = 600,

    [switch]$DisableBackgroundRefresh,

    [switch]$CalculateFull
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

function Set-BackgroundRefresh {
    param([object]$Workbook, [bool]$Enabled)
    $changed = @()
    foreach ($sheet in @($Workbook.Worksheets)) {
        foreach ($listObject in @($sheet.ListObjects)) {
            try {
                $old = [bool]$listObject.QueryTable.BackgroundQuery
                $listObject.QueryTable.BackgroundQuery = $Enabled
                $changed += [ordered]@{
                    sheet = [string]$sheet.Name
                    listObject = [string]$listObject.Name
                    oldBackgroundQuery = $old
                    newBackgroundQuery = $Enabled
                }
            } catch {
                # Not every ListObject has a QueryTable.
            }
        }
    }
    return $changed
}

function Snapshot-Connections {
    param([object]$Workbook)
    $items = @()
    foreach ($connection in @($Workbook.Connections)) {
        $items += [ordered]@{
            name = Try-Read { [string]$connection.Name } ""
            type = Try-Read { [int]$connection.Type } $null
            refreshing = Try-Read { [bool]$connection.Refreshing } $null
            description = Try-Read { [string]$connection.Description } ""
        }
    }
    return $items
}

function Text-MatchesQueryName {
    param(
        [object]$Value,
        [string]$QueryName
    )
    if ([string]::IsNullOrWhiteSpace($QueryName) -or $null -eq $Value) {
        return $false
    }
    $text = (@($Value) | ForEach-Object { [string]$_ }) -join " "
    return ($text -like "*[$QueryName]*" -or $text -like "*Location=$QueryName*" -or $text -like "*'$QueryName'*" -or $text -eq $QueryName)
}

function Refresh-LoadedQueryTargets {
    param(
        [object]$Workbook,
        [string]$QueryName
    )
    $items = @()

    foreach ($sheet in @($Workbook.Worksheets)) {
        foreach ($listObject in @($sheet.ListObjects)) {
            try {
                $queryTable = $listObject.QueryTable
                $connectionName = Try-Read { [string]$queryTable.WorkbookConnection.Name } ""
                $connectionText = Try-Read { $queryTable.Connection } $null
                $commandText = Try-Read { $queryTable.CommandText } $null
                if (
                    $connectionName -eq "Query - $QueryName" -or
                    $connectionName -like "*$QueryName*" -or
                    (Text-MatchesQueryName $connectionText $QueryName) -or
                    (Text-MatchesQueryName $commandText $QueryName)
                ) {
                    [void]$queryTable.Refresh($false)
                    $items += [ordered]@{
                        type = "ListObject.QueryTable"
                        sheet = [string]$sheet.Name
                        listObject = [string]$listObject.Name
                        connection = $connectionName
                    }
                }
            } catch {
                # Not every ListObject is backed by a Power Query QueryTable.
            }
        }
    }

    return $items
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = $null
if ($OutputWorkbookPath) { $resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputWorkbookPath) }
$excel = $null
$workbook = $null
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$startedAt = Get-Date
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$errors = @()
$backgroundChanges = @()
$targetLoadRefreshes = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $false)

    if ($DisableBackgroundRefresh) {
        $backgroundChanges = Set-BackgroundRefresh $workbook $false
    }

    $beforeConnections = Snapshot-Connections $workbook

    if ([string]::IsNullOrWhiteSpace($QueryName)) {
        $workbook.RefreshAll()
    } else {
        $target = $null
        foreach ($query in @($workbook.Queries)) {
            if ([string]$query.Name -eq $QueryName) { $target = $query; break }
        }
        if ($null -eq $target) { throw "Query not found: $QueryName" }
        $targetLoadRefreshes = Refresh-LoadedQueryTargets $workbook $QueryName
        if (@($targetLoadRefreshes).Count -eq 0) {
            $target.Refresh()
            $targetLoadRefreshes += [ordered]@{
                type = "WorkbookQuery"
                name = $QueryName
            }
        }
    }

    $waitStarted = Get-Date
    try {
        $excel.CalculateUntilAsyncQueriesDone()
    } catch {
        $errors += [ordered]@{
            phase = "CalculateUntilAsyncQueriesDone"
            message = $_.Exception.Message
        }
    }

    while ($true) {
        $refreshing = @()
        foreach ($connection in @($workbook.Connections)) {
            $isRefreshing = Try-Read { [bool]$connection.Refreshing } $false
            if ($isRefreshing) { $refreshing += [string]$connection.Name }
        }
        if ($refreshing.Count -eq 0) { break }
        if ((New-TimeSpan -Start $waitStarted -End (Get-Date)).TotalSeconds -gt $TimeoutSeconds) {
            throw "Timed out waiting for refresh after $TimeoutSeconds seconds. Still refreshing: $($refreshing -join ', ')"
        }
        Start-Sleep -Milliseconds 500
        [void]$excel.CalculateUntilAsyncQueriesDone()
    }

    if ($CalculateFull) {
        $excel.CalculateFull()
    }

    if ($resolvedOutput) {
        $parent = Split-Path -Parent $resolvedOutput
        if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $workbook.SaveAs($resolvedOutput, (FileFormat-ForPath $resolvedOutput))
    }

    $stopwatch.Stop()
    $result = [ordered]@{
        workbookPath = $resolvedWorkbook
        outputWorkbookPath = $resolvedOutput
        queryName = $QueryName
        startedAt = $startedAt.ToString("o")
        completedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        disableBackgroundRefresh = [bool]$DisableBackgroundRefresh
        calculateFull = [bool]$CalculateFull
        backgroundChanges = $backgroundChanges
        targetLoadRefreshes = $targetLoadRefreshes
        beforeConnections = $beforeConnections
        afterConnections = Snapshot-Connections $workbook
        errors = $errors
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
        workbookPath = $resolvedWorkbook
        queryName = $QueryName
        startedAt = $startedAt.ToString("o")
        failedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        error = $_.Exception.Message
        errors = $errors
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
