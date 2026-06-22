param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$SqlText = "SELECT * FROM [Data$]",

    [string]$OutJson,

    [switch]$CreateFixture,

    [switch]$IncludeSchema,

    [string]$Provider = "Microsoft.ACE.OLEDB.12.0"
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

function Get-ExcelExtendedProperties {
    param([string]$Path)
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        ".xls" { return "Excel 8.0;HDR=YES;IMEX=1" }
        ".xlsm" { return "Excel 12.0 Macro;HDR=YES;IMEX=1" }
        ".xlsb" { return "Excel 12.0;HDR=YES;IMEX=1" }
        default { return "Excel 12.0 Xml;HDR=YES;IMEX=1" }
    }
}

function New-AdoFixtureWorkbook {
    param([string]$Path)
    $parent = Split-Path -Parent $Path
    if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

    $preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
    $excel = $null
    $workbook = $null
    $worksheet = $null
    $excelProcessId = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = Get-ExcelProcessId $excel
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Data"

        $values = New-Object 'object[,]' 5, 4
        $values[0, 0] = "Region"; $values[0, 1] = "Category"; $values[0, 2] = "Amount"; $values[0, 3] = "Period"
        $values[1, 0] = "North";  $values[1, 1] = "A";        $values[1, 2] = 100;      $values[1, 3] = "2026Q1"
        $values[2, 0] = "North";  $values[2, 1] = "B";        $values[2, 2] = 150;      $values[2, 3] = "2026Q1"
        $values[3, 0] = "South";  $values[3, 1] = "A";        $values[3, 2] = 200;      $values[3, 3] = "2026Q2"
        $values[4, 0] = "East";   $values[4, 1] = "B";        $values[4, 2] = 50;       $values[4, 3] = "2026Q2"

        $worksheet.Range("A1:D5").Value2 = $values
        $range = $worksheet.Range("A1:D5")
        $listObject = $worksheet.ListObjects.Add(1, $range, $null, 1)
        $listObject.Name = "SalesTable"
        $worksheet.Columns.AutoFit() | Out-Null

        $workbook.SaveAs($Path, 51)
    } finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
            Release-ComObject $workbook
        }
        Release-ComObject $worksheet
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
}

function Convert-Recordset {
    param([object]$Recordset)
    $fields = @()
    for ($i = 0; $i -lt [int]$Recordset.Fields.Count; $i++) {
        $field = $Recordset.Fields.Item($i)
        $fields += [ordered]@{
            index = $i
            name = [string]$field.Name
            type = [int]$field.Type
            definedSize = [int]$field.DefinedSize
        }
    }

    $rows = @()
    while (-not [bool]$Recordset.EOF) {
        $row = [ordered]@{}
        foreach ($field in $fields) {
            $value = $Recordset.Fields.Item($field.name).Value
            if ($null -ne $value -and $value -is [datetime]) {
                $value = $value.ToString("o")
            }
            $row[$field.name] = $value
        }
        $rows += $row
        $Recordset.MoveNext()
    }

    return [ordered]@{
        fields = $fields
        rows = $rows
        rowCount = $rows.Count
    }
}

function Get-SchemaTables {
    param([object]$Connection)
    $schema = @()
    try {
        $recordset = $Connection.OpenSchema(20)
        while (-not [bool]$recordset.EOF) {
            $schema += [ordered]@{
                tableName = [string]$recordset.Fields.Item("TABLE_NAME").Value
                tableType = [string]$recordset.Fields.Item("TABLE_TYPE").Value
            }
            $recordset.MoveNext()
        }
        $recordset.Close()
        Release-ComObject $recordset
    } catch {
        $schema += [ordered]@{
            error = $_.Exception.Message
        }
    }
    return $schema
}

$startedAt = Get-Date
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$resolvedWorkbook = $WorkbookPath
$connection = $null
$recordset = $null

try {
    if ($CreateFixture) {
        New-AdoFixtureWorkbook $WorkbookPath
    }
    $resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $extendedProperties = Get-ExcelExtendedProperties $resolvedWorkbook
    $connectionString = "Provider=$Provider;Data Source=$resolvedWorkbook;Extended Properties=""$extendedProperties"";"

    $connection = New-Object -ComObject ADODB.Connection
    $connection.Open($connectionString)
    $recordset = New-Object -ComObject ADODB.Recordset
    $recordset.Open($SqlText, $connection, 0, 1)
    $data = Convert-Recordset $recordset
    $schemaTables = @()
    if ($IncludeSchema) {
        $schemaTables = @(Get-SchemaTables $connection)
    }

    $stopwatch.Stop()
    $result = [ordered]@{
        succeeded = $true
        workbookPath = $resolvedWorkbook
        provider = $Provider
        extendedProperties = $extendedProperties
        sqlText = $SqlText
        rowCount = $data.rowCount
        fields = $data.fields
        rows = $data.rows
        schemaTables = @($schemaTables)
        startedAt = $startedAt.ToString("o")
        completedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        limitations = @(
            "Requires a compatible ACE OLEDB provider for Excel workbook queries.",
            "Uses worksheet/table SQL over a saved workbook file; it does not query unrefreshed in-memory Power Query results.",
            "Power Pivot/Data Model queries require ADOMD or cube formulas and are not covered by worksheet SQL."
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
        succeeded = $false
        workbookPath = $resolvedWorkbook
        provider = $Provider
        sqlText = $SqlText
        error = $_.Exception.Message
        startedAt = $startedAt.ToString("o")
        failedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
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
    if ($null -ne $recordset) {
        try { if ([int]$recordset.State -ne 0) { $recordset.Close() } } catch {}
        Release-ComObject $recordset
    }
    if ($null -ne $connection) {
        try { if ([int]$connection.State -ne 0) { $connection.Close() } } catch {}
        Release-ComObject $connection
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
