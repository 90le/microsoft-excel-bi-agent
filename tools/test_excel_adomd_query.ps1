param(
    [string]$ConnectionString,

    [string]$Mdx,

    [string]$MdxPath,

    [string]$OutJson,

    [int]$MaxCells = 500,

    [switch]$ProbeOnly
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Redact-ConnectionString {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $sensitiveKeys = @(
        "Password",
        "Pwd",
        "User ID",
        "UID",
        "EffectiveUserName",
        "Access Token"
    )
    $parts = @()
    foreach ($part in ($Text -split ";")) {
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        $eq = $part.IndexOf("=")
        if ($eq -lt 0) {
            $parts += $part
            continue
        }
        $key = $part.Substring(0, $eq).Trim()
        $value = $part.Substring($eq + 1)
        if ($sensitiveKeys -contains $key) {
            $parts += "$key=***"
        } else {
            $parts += "$key=$value"
        }
    }
    return ($parts -join ";")
}

function Read-MdxText {
    param([string]$InlineMdx, [string]$Path)
    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        return Get-Content -LiteralPath (Resolve-Path -LiteralPath $Path).Path -Raw -Encoding UTF8
    }
    return $InlineMdx
}

function Test-ComProgId {
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

function Get-SafeProperty {
    param([scriptblock]$Script, [object]$Default = $null)
    try { & $Script } catch { $Default }
}

function Convert-CellValue {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToString("o") }
    return $Value
}

function Read-Axes {
    param([object]$Cellset)
    $axes = @()
    $axisCount = Get-SafeProperty { [int]$Cellset.Axes.Count } 0
    for ($axisIndex = 0; $axisIndex -lt $axisCount; $axisIndex++) {
        $axis = Get-SafeProperty { $Cellset.Axes.Item($axisIndex) } $null
        if ($null -eq $axis) { continue }
        $positions = @()
        $positionCount = Get-SafeProperty { [int]$axis.Positions.Count } 0
        for ($positionIndex = 0; $positionIndex -lt $positionCount; $positionIndex++) {
            $position = Get-SafeProperty { $axis.Positions.Item($positionIndex) } $null
            if ($null -eq $position) { continue }
            $members = @()
            $memberCount = Get-SafeProperty { [int]$position.Members.Count } 0
            for ($memberIndex = 0; $memberIndex -lt $memberCount; $memberIndex++) {
                $member = Get-SafeProperty { $position.Members.Item($memberIndex) } $null
                if ($null -eq $member) { continue }
                $members += [ordered]@{
                    index = $memberIndex
                    name = Get-SafeProperty { [string]$member.Name } ""
                    caption = Get-SafeProperty { [string]$member.Caption } ""
                    uniqueName = Get-SafeProperty { [string]$member.UniqueName } ""
                    levelName = Get-SafeProperty { [string]$member.LevelName } ""
                    hierarchy = Get-SafeProperty { [string]$member.Hierarchy } ""
                }
            }
            $positions += [ordered]@{
                index = $positionIndex
                members = $members
            }
        }
        $axes += [ordered]@{
            index = $axisIndex
            name = Get-SafeProperty { [string]$axis.Name } ""
            positionCount = $positionCount
            positions = $positions
        }
    }
    return $axes
}

function Read-Cells {
    param([object]$Cellset, [int]$Limit)
    $items = @()
    $cellCount = Get-SafeProperty { [int]$Cellset.Cells.Count } 0
    $readCount = [Math]::Min($cellCount, [Math]::Max(0, $Limit))
    for ($cellIndex = 0; $cellIndex -lt $readCount; $cellIndex++) {
        $cell = Get-SafeProperty { $Cellset.Cells.Item($cellIndex) } $null
        if ($null -eq $cell) { continue }
        $items += [ordered]@{
            ordinal = $cellIndex
            value = Convert-CellValue (Get-SafeProperty { $cell.Value } $null)
            formattedValue = Get-SafeProperty { [string]$cell.FormattedValue } ""
            cellOrdinal = Get-SafeProperty { [int]$cell.Ordinal } $cellIndex
        }
    }
    return [ordered]@{
        totalCellCount = $cellCount
        returnedCellCount = $items.Count
        truncated = ($cellCount -gt $items.Count)
        cells = $items
    }
}

function Invoke-AdomdQuery {
    param([string]$ConnString, [string]$Query, [int]$CellLimit)
    $connection = $null
    $cellset = $null
    try {
        $connection = New-Object -ComObject ADODB.Connection
        $connection.Open($ConnString)

        $cellset = New-Object -ComObject ADOMD.Cellset
        $cellset.ActiveConnection = $connection
        $cellset.Source = $Query
        $cellset.Open()

        $cells = Read-Cells $cellset $CellLimit
        return [ordered]@{
            succeeded = $true
            axes = Read-Axes $cellset
            totalCellCount = $cells.totalCellCount
            returnedCellCount = $cells.returnedCellCount
            truncated = $cells.truncated
            cells = $cells.cells
            error = ""
        }
    } catch {
        return [ordered]@{
            succeeded = $false
            axes = @()
            totalCellCount = 0
            returnedCellCount = 0
            truncated = $false
            cells = @()
            error = $_.Exception.Message
        }
    } finally {
        if ($null -ne $cellset) {
            try { $cellset.Close() } catch {}
            Release-ComObject $cellset
        }
        if ($null -ne $connection) {
            try { if ([int]$connection.State -ne 0) { $connection.Close() } } catch {}
            Release-ComObject $connection
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$startedAt = Get-Date
$stopwatch = [Diagnostics.Stopwatch]::StartNew()

$probe = [ordered]@{
    adodbConnection = Test-ComProgId "ADODB.Connection"
    adomdCatalog = Test-ComProgId "ADOMD.Catalog"
    adomdCellset = Test-ComProgId "ADOMD.Cellset"
}

if ($ProbeOnly) {
    $stopwatch.Stop()
    $result = [ordered]@{
        mode = "ProbeOnly"
        succeeded = [bool]($probe.adodbConnection.creatable -and $probe.adomdCellset.creatable)
        probe = $probe
        startedAt = $startedAt.ToString("o")
        completedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        limitations = @(
            "ProbeOnly confirms COM activation only.",
            "A real ADOMD query still requires a valid MSOLAP/ADOMD endpoint connection string and MDX query."
        )
    }
} else {
    $queryText = Read-MdxText $Mdx $MdxPath
    if ([string]::IsNullOrWhiteSpace($ConnectionString)) {
        throw "ConnectionString is required unless -ProbeOnly is supplied."
    }
    if ([string]::IsNullOrWhiteSpace($queryText)) {
        throw "Mdx or MdxPath is required unless -ProbeOnly is supplied."
    }

    $queryResult = Invoke-AdomdQuery $ConnectionString $queryText $MaxCells
    $stopwatch.Stop()
    $result = [ordered]@{
        mode = "Query"
        succeeded = [bool]$queryResult.succeeded
        connectionString = Redact-ConnectionString $ConnectionString
        mdx = $queryText
        maxCells = $MaxCells
        probe = $probe
        query = $queryResult
        startedAt = $startedAt.ToString("o")
        completedAt = (Get-Date).ToString("o")
        elapsedSeconds = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
        limitations = @(
            "This script validates an ADOMD/MSOLAP endpoint query, not worksheet SQL.",
            "It does not create or expose an Excel in-process ThisWorkbookDataModel endpoint.",
            "Connection strings can contain sensitive information; password-like keys are redacted in output."
        )
    }
}

$json = $result | ConvertTo-Json -Depth 30
if ($OutJson) {
    $outParent = Split-Path -Parent $OutJson
    if ($outParent) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
    Set-Content -LiteralPath $OutJson -Value $json -Encoding UTF8
} else {
    $json
}
