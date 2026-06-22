param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$OutDir
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

function Safe-FileName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $chars = foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -contains $ch) { "_" } else { $ch }
    }
    $safe = -join $chars
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "query"
    }
    return $safe
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

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
New-Item -ItemType Directory -Path $resolvedOutDir -Force | Out-Null

$excel = $null
$workbook = $null
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $true)

    $queries = @()
    $queryIndex = 0
    foreach ($query in @($workbook.Queries)) {
        $queryIndex += 1
        $name = [string]$query.Name
        $formula = [string]$query.Formula
        $safeName = "{0:D3}_{1}.m" -f $queryIndex, (Safe-FileName $name)
        $path = Join-Path $resolvedOutDir $safeName
        Set-Content -LiteralPath $path -Value $formula -Encoding UTF8
        $queries += [ordered]@{
            index       = $queryIndex
            name        = $name
            description = Try-Read { [string]$query.Description } ""
            formulaFile = $path
            formula     = $formula
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

    $result = [ordered]@{
        workbookPath = $resolvedWorkbook
        outDir       = $resolvedOutDir
        queryCount   = $queries.Count
        queries      = $queries
        connections  = $connections
    }

    $jsonPath = Join-Path $resolvedOutDir "power_queries.json"
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    $result | ConvertTo-Json -Depth 8
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
