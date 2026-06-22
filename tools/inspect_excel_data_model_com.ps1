param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$OutJson,

    [switch]$IncludeColumns
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

function Read-ModelTables {
    param([object]$Model, [bool]$WithColumns)
    $items = @()
    $count = Try-Read { [int]$Model.ModelTables.Count } 0
    for ($index = 1; $index -le $count; $index++) {
        $table = Try-Read { $Model.ModelTables.Item($index) } $null
        if ($null -eq $table) { continue }
        $columnCount = Try-Read { [int]$table.ModelTableColumns.Count } 0
        $record = [ordered]@{
            index = $index
            name = Try-Read { [string]$table.Name } ""
            sourceName = Try-Read { [string]$table.SourceName } ""
            recordCount = Try-Read { [int]$table.RecordCount } $null
            columnCount = $columnCount
        }
        if ($WithColumns) {
            $columns = @()
            for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
                $column = Try-Read { $table.ModelTableColumns.Item($columnIndex) } $null
                if ($null -eq $column) { continue }
                $columns += [ordered]@{
                    index = $columnIndex
                    name = Try-Read { [string]$column.Name } ""
                    sourceName = Try-Read { [string]$column.SourceName } ""
                    dataType = Try-Read { [string]$column.DataType } ""
                    visible = Try-Read { [bool]$column.Visible } $null
                    description = Try-Read { [string]$column.Description } ""
                }
            }
            $record.columns = $columns
        }
        $items += $record
    }
    return $items
}

function Read-ModelRelationships {
    param([object]$Model)
    $items = @()
    $count = Try-Read { [int]$Model.ModelRelationships.Count } 0
    for ($index = 1; $index -le $count; $index++) {
        $relationship = Try-Read { $Model.ModelRelationships.Item($index) } $null
        if ($null -eq $relationship) { continue }
        $items += [ordered]@{
            index = $index
            foreignKeyTable = Try-Read { [string]$relationship.ForeignKeyTable.Name } ""
            foreignKeyColumn = Try-Read { [string]$relationship.ForeignKeyColumn.Name } ""
            primaryKeyTable = Try-Read { [string]$relationship.PrimaryKeyTable.Name } ""
            primaryKeyColumn = Try-Read { [string]$relationship.PrimaryKeyColumn.Name } ""
            active = Try-Read { [bool]$relationship.Active } $null
        }
    }
    return $items
}

function Read-ModelMeasures {
    param([object]$Model)
    $items = @()
    $count = Try-Read { [int]$Model.ModelMeasures.Count } 0
    for ($index = 1; $index -le $count; $index++) {
        $measure = Try-Read { $Model.ModelMeasures.Item($index) } $null
        if ($null -eq $measure) { continue }
        $items += [ordered]@{
            index = $index
            name = Try-Read { [string]$measure.Name } ""
            associatedTable = Try-Read { [string]$measure.AssociatedTable.Name } ""
            formula = Try-Read { [string]$measure.Formula } ""
            description = Try-Read { [string]$measure.Description } ""
            formatInformation = Try-Read { [string]$measure.FormatInformation } ""
        }
    }
    return $items
}

function Read-Connections {
    param([object]$Workbook)
    $items = @()
    foreach ($connection in @($Workbook.Connections)) {
        $items += [ordered]@{
            name = Try-Read { [string]$connection.Name } ""
            type = Try-Read { [int]$connection.Type } $null
            description = Try-Read { [string]$connection.Description } ""
            refreshing = Try-Read { [bool]$connection.Refreshing } $null
            hasModelConnection = [bool](Try-Read { $null -ne $connection.ModelConnection } $false)
            oledbConnection = Try-Read { [string]$connection.OLEDBConnection.Connection } ""
        }
    }
    return $items
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$excel = $null
$workbook = $null
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

    $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $true)
    $model = Try-Read { $workbook.Model } $null
    $tables = @()
    $relationships = @()
    $measures = @()
    $modelAvailable = $null -ne $model
    if ($modelAvailable) {
        $tables = Read-ModelTables $model ([bool]$IncludeColumns)
        $relationships = Read-ModelRelationships $model
        $measures = Read-ModelMeasures $model
    }

    $stopwatch.Stop()
    $result = [ordered]@{
        workbookPath = $resolvedWorkbook
        readOnly = $true
        modelAvailable = $modelAvailable
        tableCount = $tables.Count
        relationshipCount = $relationships.Count
        measureCount = $measures.Count
        tables = $tables
        relationships = $relationships
        measures = $measures
        connections = Read-Connections $workbook
        startedAt = $startedAt.ToString("o")
        completedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        limitations = @(
            "This script requires Windows desktop Excel COM.",
            "It reads Excel's object model and does not refresh Power Query or recalculate formulas.",
            "Some Data Model metadata can vary by Excel version and installed components."
        )
    }

    $json = $result | ConvertTo-Json -Depth 20
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
