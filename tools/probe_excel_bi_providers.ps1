param(
    [string]$OutJson,

    [string[]]$Provider = @(
        "Microsoft.ACE.OLEDB.16.0",
        "Microsoft.ACE.OLEDB.12.0",
        "Microsoft.Jet.OLEDB.4.0",
        "MSOLAP",
        "MSOLAP.8",
        "MSOLAP.7",
        "MSOLAP.6",
        "MSOLAP.5",
        "MSOLAP.4",
        "MSDASQL",
        "MSOLEDBSQL",
        "SQLOLEDB"
    ),

    [string[]]$ComProgId = @(
        "Excel.Application",
        "ADODB.Connection",
        "ADODB.Recordset",
        "ADOMD.Catalog",
        "ADOMD.Cellset"
    ),

    [switch]$RunExcelComSmoke,

    [switch]$RunAdoWorkbookSmoke,

    [string]$SmokeWorkbookPath,

    [string]$AdoSmokeProvider = "Microsoft.ACE.OLEDB.12.0"
)

$ErrorActionPreference = "Stop"
$ScriptRoot = Split-Path -Parent $PSCommandPath

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Get-RegistryDefaultValue {
    param([string]$Path)
    try {
        $key = Get-Item -LiteralPath $Path -ErrorAction Stop
        return [string]$key.GetValue("")
    } catch {
        return ""
    }
}

function Get-ProgIdRegistryInfo {
    param([string]$ProgId)
    $paths = @(
        "Registry::HKEY_CLASSES_ROOT\$ProgId",
        "Registry::HKEY_CLASSES_ROOT\WOW6432Node\$ProgId",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$ProgId",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\$ProgId"
    )
    $hits = @()
    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            $clsid = Get-RegistryDefaultValue (Join-Path $path "CLSID")
            $hits += [ordered]@{
                path = $path
                defaultValue = Get-RegistryDefaultValue $path
                clsid = $clsid
            }
        }
    }
    return @($hits)
}

function Test-ComCreation {
    param([string]$ProgId)
    $obj = $null
    try {
        $obj = New-Object -ComObject $ProgId
        return [ordered]@{
            progId = $ProgId
            creatable = $true
            error = ""
        }
    } catch {
        return [ordered]@{
            progId = $ProgId
            creatable = $false
            error = $_.Exception.Message
        }
    } finally {
        Release-ComObject $obj
    }
}

function Test-DotNetAssemblyLoad {
    param([string]$AssemblyName)
    try {
        $assembly = [System.Reflection.Assembly]::Load($AssemblyName)
        return [ordered]@{
            assemblyName = $AssemblyName
            loadable = $true
            fullName = [string]$assembly.FullName
            error = ""
        }
    } catch {
        return [ordered]@{
            assemblyName = $AssemblyName
            loadable = $false
            fullName = ""
            error = $_.Exception.Message
        }
    }
}

function Test-ExcelCom {
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        return [ordered]@{
            succeeded = $true
            version = [string]$excel.Version
            build = [string]$excel.Build
            operatingSystem = [string]$excel.OperatingSystem
            hWnd = [int]$excel.Hwnd
            error = ""
        }
    } catch {
        return [ordered]@{
            succeeded = $false
            version = ""
            build = ""
            operatingSystem = ""
            hWnd = $null
            error = $_.Exception.Message
        }
    } finally {
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-AdoWorkbookSmoke {
    param(
        [string]$WorkbookPath,
        [string]$ProviderName
    )
    $adoScript = Join-Path $ScriptRoot "test_excel_ado_sql_access.ps1"
    if (-not (Test-Path -LiteralPath $adoScript)) {
        return [ordered]@{
            succeeded = $false
            provider = $ProviderName
            workbookPath = $WorkbookPath
            error = "Cannot find test_excel_ado_sql_access.ps1 next to this script."
        }
    }

    if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
        $WorkbookPath = Join-Path ([IO.Path]::GetTempPath()) "excel_bi_provider_probe_ado.xlsx"
    }
    $tempJson = [IO.Path]::ChangeExtension($WorkbookPath, ".ado_probe.json")
    $sql = "SELECT Category, SUM(Amount) AS TotalAmount FROM [Data$] GROUP BY Category ORDER BY Category"

    try {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $adoScript `
            -WorkbookPath $WorkbookPath `
            -CreateFixture `
            -SqlText $sql `
            -IncludeSchema `
            -Provider $ProviderName `
            -OutJson $tempJson | Out-Null
        $raw = Get-Content -LiteralPath $tempJson -Raw -Encoding UTF8
        $parsed = $raw | ConvertFrom-Json
        return [ordered]@{
            succeeded = [bool]$parsed.succeeded
            provider = $ProviderName
            workbookPath = [string]$parsed.workbookPath
            rowCount = [int]$parsed.rowCount
            fields = @($parsed.fields | ForEach-Object { [string]$_.name })
            schemaTableCount = @($parsed.schemaTables).Count
            error = ""
        }
    } catch {
        $errorMessage = $_.Exception.Message
        if (Test-Path -LiteralPath $tempJson) {
            try {
                $raw = Get-Content -LiteralPath $tempJson -Raw -Encoding UTF8
                $parsed = $raw | ConvertFrom-Json
                if ($parsed.error) { $errorMessage = [string]$parsed.error }
            } catch {}
        }
        return [ordered]@{
            succeeded = $false
            provider = $ProviderName
            workbookPath = $WorkbookPath
            rowCount = $null
            fields = @()
            schemaTableCount = 0
            error = $errorMessage
        }
    }
}

$startedAt = Get-Date
$providerResults = @()
foreach ($name in $Provider) {
    $registryHits = @(Get-ProgIdRegistryInfo $name)
    $providerResults += [ordered]@{
        provider = $name
        registryDetected = ($registryHits.Count -gt 0)
        registryHits = $registryHits
    }
}

$comResults = @()
foreach ($progId in $ComProgId) {
    $comResults += Test-ComCreation $progId
}

$excelComSmoke = $null
if ($RunExcelComSmoke) {
    $excelComSmoke = Test-ExcelCom
}

$adoWorkbookSmoke = $null
if ($RunAdoWorkbookSmoke) {
    $adoWorkbookSmoke = Invoke-AdoWorkbookSmoke -WorkbookPath $SmokeWorkbookPath -ProviderName $AdoSmokeProvider
}

$result = [ordered]@{
    startedAt = $startedAt.ToString("o")
    completedAt = (Get-Date).ToString("o")
    machine = [ordered]@{
        computerName = [string]$env:COMPUTERNAME
        osVersion = [string][Environment]::OSVersion.VersionString
        is64BitOperatingSystem = [bool][Environment]::Is64BitOperatingSystem
        is64BitProcess = [bool][Environment]::Is64BitProcess
        powershellVersion = [string]$PSVersionTable.PSVersion
    }
    providers = $providerResults
    comProgIds = $comResults
    dotNetAssemblies = @(
        Test-DotNetAssemblyLoad "Microsoft.AnalysisServices.AdomdClient"
    )
    excelComSmoke = $excelComSmoke
    adoWorkbookSmoke = $adoWorkbookSmoke
    interpretation = @(
        "registryDetected means a ProgID key exists; it does not prove a full connection can open.",
        "creatable COM ProgIDs prove COM activation only; connection strings and endpoints still need task-specific tests.",
        "RunAdoWorkbookSmoke proves ACE workbook SQL over a generated worksheet fixture for the selected provider.",
        "ADOMD/MSOLAP availability does not by itself prove a workbook Data Model endpoint is queryable."
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
