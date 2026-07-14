param(
    [Parameter(Mandatory = $true)]
    [string]$OutJson,

    [ValidateSet("inventory", "runtime")]
    [string]$Profile = "inventory",

    [string]$AdoSmokeProvider = "Microsoft.ACE.OLEDB.12.0"
)

$ErrorActionPreference = "Stop"
$ScriptRoot = Split-Path -Parent $PSCommandPath
$ProviderProbeScript = Join-Path $ScriptRoot "probe_excel_bi_providers.ps1"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Get-ExcelProcessIds {
    return @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
}

function Stop-NewExcelProcesses {
    param([int[]]$ExistingIds)
    $existing = @{}
    foreach ($processId in @($ExistingIds)) { $existing[$processId] = $true }
    foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if (-not $existing.ContainsKey([int]$process.Id)) {
            try {
                if (-not $process.HasExited) {
                    $process.CloseMainWindow() | Out-Null
                    $process.WaitForExit(1500) | Out-Null
                }
            } catch {}
            try {
                if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
            } catch {}
        }
    }
}

function Get-ErrorCategory {
    param([string]$Message)
    $text = $Message.ToLowerInvariant()
    if ($text.Contains("class not registered") -or $text.Contains("cannot create activex")) { return "class-not-registered" }
    if ($text.Contains("provider") -or $text.Contains("oledb")) { return "provider-unavailable" }
    if ($text.Contains("vbproject") -or $text.Contains("programmatic access") -or $text.Contains("access denied")) { return "access-denied" }
    return "probe-failed"
}

function New-Capability {
    param(
        [ValidateSet("pass", "fail", "skip", "error")]
        [string]$Status,
        [ValidateSet("registration", "activation", "smoke", "not-tested")]
        [string]$EvidenceLevel,
        [string]$Detail,
        [string]$ErrorMessage = "",
        [string]$ErrorCategory = ""
    )
    if ($ErrorMessage -and -not $ErrorCategory) { $ErrorCategory = Get-ErrorCategory $ErrorMessage }
    return [ordered]@{
        status = $Status
        evidenceLevel = $EvidenceLevel
        detail = $Detail
        errorCategory = $ErrorCategory
        error = $ErrorMessage
    }
}

function Find-ProviderRow {
    param([object]$ProviderReport, [string]$Name)
    foreach ($item in @($ProviderReport.providers)) {
        if ([string]$item.provider -eq $Name) { return $item }
    }
    return $null
}

function Find-ComRow {
    param([object]$ProviderReport, [string]$ProgId)
    foreach ($item in @($ProviderReport.comProgIds)) {
        if ([string]$item.progId -eq $ProgId) { return $item }
    }
    return $null
}

function Test-ExcelRegistration {
    $paths = @(
        "Registry::HKEY_CLASSES_ROOT\Excel.Application",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Excel.Application",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\Excel.Application"
    )
    return [bool](@($paths | Where-Object { Test-Path -LiteralPath $_ }).Count -gt 0)
}

$existingExcelIds = Get-ExcelProcessIds
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("excel_capability_probe_" + [Guid]::NewGuid().ToString("N"))
$providerJson = Join-Path $tempRoot "provider_probe.json"
$smokeWorkbook = Join-Path $tempRoot "provider_ado_smoke.xlsx"
$roundtripWorkbook = Join-Path $tempRoot "excel_capability_roundtrip.xlsx"
$pdfPath = Join-Path $tempRoot "excel_capability_render.pdf"
$errors = @()
$providerReport = $null
$capabilities = [ordered]@{}
$excelVersion = ""
$excelBuild = ""

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    if (-not (Test-Path -LiteralPath $ProviderProbeScript)) {
        throw "Cannot find probe_excel_bi_providers.ps1 next to this script."
    }

    $providerParams = @{
        OutJson = $providerJson
        ComProgId = @("ADODB.Connection", "ADODB.Recordset", "ADOMD.Catalog", "ADOMD.Cellset")
        AdoSmokeProvider = $AdoSmokeProvider
    }
    if ($Profile -eq "runtime") {
        $providerParams["RunExcelComSmoke"] = $true
        $providerParams["RunAdoWorkbookSmoke"] = $true
        $providerParams["SmokeWorkbookPath"] = $smokeWorkbook
    }

    try {
        & $ProviderProbeScript @providerParams
        $providerReport = Get-Content -LiteralPath $providerJson -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        $errors += "provider probe failed: $($_.Exception.Message)"
    } finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Stop-NewExcelProcesses -ExistingIds $existingExcelIds
    }

    $excelRegistered = Test-ExcelRegistration
    $adodb = if ($null -ne $providerReport) { Find-ComRow $providerReport "ADODB.Connection" } else { $null }
    $adomdCatalog = if ($null -ne $providerReport) { Find-ComRow $providerReport "ADOMD.Catalog" } else { $null }
    $adomdCellset = if ($null -ne $providerReport) { Find-ComRow $providerReport "ADOMD.Cellset" } else { $null }
    $msolapRegistered = $false
    if ($null -ne $providerReport) {
        foreach ($item in @($providerReport.providers)) {
            if ([string]$item.provider -like "MSOLAP*" -and [bool]$item.registryDetected) { $msolapRegistered = $true; break }
        }
    }

    if ($Profile -eq "runtime") {
        $excelSmoke = if ($null -ne $providerReport) { $providerReport.excelComSmoke } else { $null }
        if ($null -ne $excelSmoke -and [bool]$excelSmoke.succeeded) {
            $capabilities["excel.com.activation"] = New-Capability "pass" "smoke" "Excel.Application activated and returned version/build evidence."
            $excelVersion = [string]$excelSmoke.version
            $excelBuild = [string]$excelSmoke.build
        } else {
            $message = if ($null -ne $excelSmoke) { [string]$excelSmoke.error } else { "Excel COM smoke evidence is unavailable." }
            $capabilities["excel.com.activation"] = New-Capability "fail" "activation" "Excel.Application activation failed." $message
        }
    } else {
        $detail = if ($excelRegistered) { "Excel.Application registration detected; activation was not requested." } else { "Excel.Application registration was not detected." }
        $capabilities["excel.com.activation"] = New-Capability "skip" "registration" $detail "" "runtime-not-requested"
    }

    if ($null -ne $adodb -and [bool]$adodb.creatable) {
        $capabilities["ado.com.activation"] = New-Capability "pass" "activation" "ADODB.Connection COM activation succeeded."
    } elseif ($null -ne $adodb) {
        $capabilities["ado.com.activation"] = New-Capability "fail" "activation" "ADODB.Connection COM activation failed." ([string]$adodb.error)
    } else {
        $capabilities["ado.com.activation"] = New-Capability "error" "not-tested" "ADODB probe evidence is unavailable." "Provider probe did not return ADODB.Connection evidence."
    }

    if ($Profile -eq "runtime") {
        $adoSmoke = if ($null -ne $providerReport) { $providerReport.adoWorkbookSmoke } else { $null }
        if ($null -ne $adoSmoke -and [bool]$adoSmoke.succeeded) {
            $capabilities["ace.workbook-sql"] = New-Capability "pass" "smoke" "ACE workbook SQL succeeded against a generated worksheet fixture."
        } else {
            $message = if ($null -ne $adoSmoke) { [string]$adoSmoke.error } else { "ADO workbook smoke evidence is unavailable." }
            $capabilities["ace.workbook-sql"] = New-Capability "fail" "smoke" "ACE workbook SQL smoke failed." $message
        }
    } else {
        $capabilities["ace.workbook-sql"] = New-Capability "skip" "not-tested" "ACE workbook SQL runtime smoke was not requested." "" "runtime-not-requested"
    }

    $capabilities["msolap.registration"] = if ($msolapRegistered) {
        New-Capability "pass" "registration" "At least one MSOLAP ProgID registration was detected."
    } else {
        New-Capability "fail" "registration" "No MSOLAP ProgID registration was detected." "MSOLAP registration unavailable." "class-not-registered"
    }

    if ($null -ne $adomdCatalog -and $null -ne $adomdCellset -and [bool]$adomdCatalog.creatable -and [bool]$adomdCellset.creatable) {
        $capabilities["adomd.com.activation"] = New-Capability "pass" "activation" "ADOMD.Catalog and ADOMD.Cellset COM activation succeeded."
    } else {
        $messages = @()
        if ($null -ne $adomdCatalog -and $adomdCatalog.error) { $messages += [string]$adomdCatalog.error }
        if ($null -ne $adomdCellset -and $adomdCellset.error) { $messages += [string]$adomdCellset.error }
        $message = if ($messages.Count) { $messages -join " | " } else { "ADOMD COM evidence is unavailable." }
        $capabilities["adomd.com.activation"] = New-Capability "fail" "activation" "ADOMD COM activation is unavailable." $message
    }

    $runtimeOnlyIds = @(
        "excel.workbook.roundtrip",
        "excel.vba.project-access",
        "excel.power-query.object-model",
        "excel.power-query.async-wait",
        "excel.data-model.object-model",
        "excel.pdf-export"
    )
    if ($Profile -eq "inventory") {
        foreach ($capabilityId in $runtimeOnlyIds) {
            $capabilities[$capabilityId] = New-Capability "skip" "not-tested" "Runtime profile was not requested." "" "runtime-not-requested"
        }
    } elseif ($capabilities["excel.com.activation"].status -ne "pass") {
        foreach ($capabilityId in $runtimeOnlyIds) {
            $capabilities[$capabilityId] = New-Capability "skip" "not-tested" "Blocked because Excel COM activation did not pass." "" "blocked-by-dependency"
        }
    } else {
        $excel = $null
        $workbook = $null
        $worksheet = $null
        $model = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Cells.Item(1, 1).Value2 = "Excel capability fixture"
            $workbook.SaveAs($roundtripWorkbook, 51)
            $workbook.Close($false)
            Release-ComObject $worksheet
            $worksheet = $null
            Release-ComObject $workbook
            $workbook = $excel.Workbooks.Open($roundtripWorkbook)
            $worksheet = $workbook.Worksheets.Item(1)
            if ([string]$worksheet.Cells.Item(1, 1).Value2 -eq "Excel capability fixture") {
                $capabilities["excel.workbook.roundtrip"] = New-Capability "pass" "smoke" "A generated workbook was saved, reopened, and read successfully."
            } else {
                $capabilities["excel.workbook.roundtrip"] = New-Capability "fail" "smoke" "Workbook roundtrip returned unexpected cell content." "Unexpected fixture content."
            }

            try {
                $componentCount = [int]$workbook.VBProject.VBComponents.Count
                $capabilities["excel.vba.project-access"] = New-Capability "pass" "activation" "VBProject access succeeded; componentCount=$componentCount."
            } catch {
                $capabilities["excel.vba.project-access"] = New-Capability "fail" "activation" "VBProject access failed; check Trust Center policy." $_.Exception.Message
            }

            try {
                $queryCount = [int]$workbook.Queries.Count
                $capabilities["excel.power-query.object-model"] = New-Capability "pass" "activation" "Workbook.Queries object model access succeeded; queryCount=$queryCount."
            } catch {
                $capabilities["excel.power-query.object-model"] = New-Capability "fail" "activation" "Workbook.Queries object model access failed." $_.Exception.Message
            }

            try {
                [void]$excel.CalculateUntilAsyncQueriesDone()
                $capabilities["excel.power-query.async-wait"] = New-Capability "pass" "smoke" "CalculateUntilAsyncQueriesDone completed on the generated workbook."
            } catch {
                $capabilities["excel.power-query.async-wait"] = New-Capability "fail" "smoke" "CalculateUntilAsyncQueriesDone failed." $_.Exception.Message
            }

            try {
                $model = $workbook.Model
                if ($null -eq $model) { throw "Workbook.Model returned null." }
                $capabilities["excel.data-model.object-model"] = New-Capability "pass" "activation" "Workbook.Model object model access succeeded."
            } catch {
                $capabilities["excel.data-model.object-model"] = New-Capability "fail" "activation" "Workbook.Model object model access failed." $_.Exception.Message
            } finally {
                Release-ComObject $model
                $model = $null
            }

            try {
                $worksheet.ExportAsFixedFormat(0, $pdfPath)
                if (-not (Test-Path -LiteralPath $pdfPath)) { throw "PDF output was not created." }
                $capabilities["excel.pdf-export"] = New-Capability "pass" "smoke" "A worksheet PDF was exported from the generated workbook."
            } catch {
                $capabilities["excel.pdf-export"] = New-Capability "fail" "smoke" "Worksheet PDF export failed." $_.Exception.Message
            }
        } catch {
            $errors += "runtime workbook probe failed: $($_.Exception.Message)"
            foreach ($capabilityId in $runtimeOnlyIds) {
                if (-not $capabilities.Contains($capabilityId)) {
                    $capabilities[$capabilityId] = New-Capability "error" "not-tested" "Runtime workbook probe ended before this capability was tested." $_.Exception.Message
                }
            }
        } finally {
            if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
            Release-ComObject $model
            Release-ComObject $worksheet
            Release-ComObject $workbook
            if ($null -ne $excel) { try { $excel.Quit() } catch {} }
            Release-ComObject $excel
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Stop-NewExcelProcesses -ExistingIds $existingExcelIds
        }
    }
} catch {
    $errors += $_.Exception.Message
} finally {
    Stop-NewExcelProcesses -ExistingIds $existingExcelIds
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$allCapabilityIds = @(
    "excel.com.activation",
    "excel.workbook.roundtrip",
    "excel.vba.project-access",
    "excel.power-query.object-model",
    "excel.power-query.async-wait",
    "excel.data-model.object-model",
    "excel.pdf-export",
    "ado.com.activation",
    "ace.workbook-sql",
    "msolap.registration",
    "adomd.com.activation"
)
foreach ($capabilityId in $allCapabilityIds) {
    if (-not $capabilities.Contains($capabilityId)) {
        $capabilities[$capabilityId] = New-Capability "error" "not-tested" "Probe did not produce capability evidence." "Capability result missing." "invalid-probe-state"
    }
}

$result = [ordered]@{
    schemaVersion = "1.0"
    kind = "excel-capability-probe"
    generatedAt = (Get-Date).ToUniversalTime().ToString("o")
    probe = [ordered]@{
        profile = $Profile
        platform = "windows"
        syntheticFixture = ($Profile -eq "runtime")
    }
    environment = [ordered]@{
        osVersion = [string][Environment]::OSVersion.VersionString
        is64BitOperatingSystem = [bool][Environment]::Is64BitOperatingSystem
        is64BitProcess = [bool][Environment]::Is64BitProcess
        powershellVersion = [string]$PSVersionTable.PSVersion
        excelVersion = $excelVersion
        excelBuild = $excelBuild
    }
    capabilities = $capabilities
    boundaries = @(
        "Inventory profile records registration and non-Excel COM activation evidence; Excel runtime operations remain not tested.",
        "Runtime profile uses generated temporary workbooks and removes them before returning.",
        "COM or object-model activation does not prove a customer workbook's business logic.",
        "ADOMD/MSOLAP activation does not prove a real endpoint is queryable."
    )
    errors = @($errors)
}

$outPath = [IO.Path]::GetFullPath($OutJson)
$outParent = Split-Path -Parent $outPath
if ($outParent) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
$result | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $outPath -Encoding UTF8
