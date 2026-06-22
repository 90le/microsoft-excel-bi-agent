param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$SourceDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputWorkbookPath,

    [switch]$KeepExistingImportableComponents
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

function Read-VbName {
    param([string]$Path)
    $lines = Get-Content -LiteralPath $Path -TotalCount 20 -ErrorAction Stop
    foreach ($line in $lines) {
        if ($line -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            return $Matches[1]
        }
    }
    return [IO.Path]::GetFileNameWithoutExtension($Path)
}

function FileFormat-ForPath {
    param([string]$Path)
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        ".xlsm" { 52 }
        ".xlsb" { 50 }
        ".xls"  { 56 }
        default {
            throw "OutputWorkbookPath must end with .xlsm, .xlsb, or .xls for VBA import."
        }
    }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSourceDir = (Resolve-Path -LiteralPath $SourceDir).Path
$resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputWorkbookPath)

if ([string]::Equals($resolvedWorkbook, $resolvedOutput, [StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputWorkbookPath must be different from WorkbookPath. Work on a copy."
}

$excel = $null
$workbook = $null
$imported = @()
$skippedDocumentModules = @()
$removedComponents = @()
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $false)
    $components = $workbook.VBProject.VBComponents

    $documentModuleNames = @{}
    foreach ($component in @($components)) {
        if ([int]$component.Type -eq 100) {
            $documentModuleNames[[string]$component.Name] = $true
        }
    }

    if (-not $KeepExistingImportableComponents) {
        for ($i = $components.Count; $i -ge 1; $i--) {
            $component = $components.Item($i)
            if (@(1, 2, 3) -contains [int]$component.Type) {
                $removedComponents += [string]$component.Name
                $components.Remove($component)
            }
        }
    }

    $files = Get-ChildItem -LiteralPath $resolvedSourceDir -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm") } |
        Sort-Object Name

    foreach ($file in $files) {
        $vbName = Read-VbName $file.FullName
        if ($documentModuleNames.ContainsKey($vbName)) {
            $skippedDocumentModules += $file.FullName
            continue
        }

        $components.Import($file.FullName) | Out-Null
        $imported += $file.FullName
    }

    $outParent = Split-Path -Parent $resolvedOutput
    if ($outParent) {
        New-Item -ItemType Directory -Path $outParent -Force | Out-Null
    }

    $fileFormat = FileFormat-ForPath $resolvedOutput
    $workbook.SaveAs($resolvedOutput, $fileFormat)

    [ordered]@{
        sourceWorkbook         = $resolvedWorkbook
        sourceDir              = $resolvedSourceDir
        outputWorkbook         = $resolvedOutput
        removedComponents      = $removedComponents
        imported               = $imported
        skippedDocumentModules = $skippedDocumentModules
    } | ConvertTo-Json -Depth 6
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
