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

function Component-Extension {
    param([int]$ComponentType)
    switch ($ComponentType) {
        1 { ".bas" }
        2 { ".cls" }
        3 { ".frm" }
        100 { ".cls" }
        default { ".bas" }
    }
}

$resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
New-Item -ItemType Directory -Path $resolvedOutDir -Force | Out-Null

$excel = $null
$workbook = $null
$exported = @()
$excelProcessId = $null
$preExistingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })

try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($resolvedPath, $null, $true)

    foreach ($component in @($workbook.VBProject.VBComponents)) {
        $extension = Component-Extension ([int]$component.Type)
        $target = Join-Path $resolvedOutDir ([string]$component.Name + $extension)
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Force
        }
        $component.Export($target)
        $exported += $target
    }

    [ordered]@{
        workbookPath = $resolvedPath
        outDir       = $resolvedOutDir
        exported     = $exported
    } | ConvertTo-Json -Depth 5
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
