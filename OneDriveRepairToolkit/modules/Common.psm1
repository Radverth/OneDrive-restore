#Requires -Version 7.0
<#
.SYNOPSIS
    Shared plumbing for the OneDrive Sync Repair & Reconciliation Toolkit.

.DESCRIPTION
    Everything that more than one stage needs lives here: toolkit context and
    paths, the config JSON, the rolling run log, console prompt helpers, the
    certificate-based Graph connection, a throttling-aware Graph request
    wrapper, drive enumeration, and the sync-conflict filename patterns that
    both the duplicate scanner (stage 3) and the comparison (stage 4) rely on.

    Keeping these here is what lets Start-OneDriveRepair.ps1 stay a thin
    dispatcher with no duplicated logic across the stage modules.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Toolkit context and paths
# ---------------------------------------------------------------------------

function Get-ToolkitContext {
    <#
    .SYNOPSIS
        Returns (creating on first call) the process-wide toolkit context.
    .DESCRIPTION
        Held in a global so that re-importing the stage modules during a run
        does not reset the log path, cached config or connection state.
    #>
    [CmdletBinding()]
    param()

    $existing = Get-Variable -Name 'OneDriveRepairContext' -Scope Global -ErrorAction SilentlyContinue
    if ($existing -and $existing.Value) {
        return $existing.Value
    }

    $root = Split-Path -Parent $PSScriptRoot

    $context = [ordered]@{
        Root           = $root
        ConfigDir      = Join-Path $root 'config'
        ConfigPath     = Join-Path $root 'config' 'toolkit-config.json'
        LogDir         = Join-Path $root 'logs'
        ReportDir      = Join-Path $root 'reports'
        LogPath        = $null
        Config         = $null
        GraphConnected = $false
        GraphTenantId  = $null
        GraphAppId     = $null
    }

    Set-Variable -Name 'OneDriveRepairContext' -Scope Global -Value $context
    return $context
}

function Initialize-ToolkitEnvironment {
    <#
    .SYNOPSIS
        Creates the config/logs/reports folders and opens the log for this run.
    #>
    [CmdletBinding()]
    param()

    $ctx = Get-ToolkitContext
    foreach ($dir in @($ctx.ConfigDir, $ctx.LogDir, $ctx.ReportDir)) {
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    if (-not $ctx.LogPath) {
        $ctx.LogPath = Join-Path $ctx.LogDir ('toolkit-{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
    }

    return $ctx
}

function Get-ToolkitReportPath {
    <#
    .SYNOPSIS
        Builds a timestamped path under reports/ for a stage output file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaseName,
        [string]$Extension = 'csv'
    )

    $ctx = Initialize-ToolkitEnvironment
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    return Join-Path $ctx.ReportDir ('{0}-{1}.{2}' -f $BaseName, $stamp, $Extension.TrimStart('.'))
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Write-ToolkitLog {
    <#
    .SYNOPSIS
        Writes a timestamped line to the rolling daily log and (usually) the console.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')][string]$Level = 'INFO',
        [switch]$NoConsole
    )

    $ctx = Initialize-ToolkitEnvironment
    $line = '{0} [{1,-7}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    try {
        Add-Content -LiteralPath $ctx.LogPath -Value $line -Encoding utf8
    }
    catch {
        # Never let a logging failure take down a long-running stage.
        Write-Warning ('Could not write to log file {0}: {1}' -f $ctx.LogPath, $_.Exception.Message)
    }

    if ($NoConsole) { return }

    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'DEBUG'   { 'DarkGray' }
        default   { 'Gray' }
    }
    Write-Host $Message -ForegroundColor $color
}

function Remove-OldToolkitLog {
    <#
    .SYNOPSIS
        Prunes log files older than the retention window (default 30 days).
    #>
    [CmdletBinding()]
    param([int]$RetentionDays = 30)

    $ctx = Initialize-ToolkitEnvironment
    $cutoff = (Get-Date).AddDays(-1 * [math]::Abs($RetentionDays))

    Get-ChildItem -LiteralPath $ctx.LogDir -Filter 'toolkit-*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop }
            catch { Write-Verbose ('Could not prune log {0}: {1}' -f $_.FullName, $_.Exception.Message) }
        }
}

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

function New-ToolkitConfigObject {
    [CmdletBinding()]
    param()

    return [ordered]@{
        TenantId                 = ''
        AppId                    = ''
        CertificateThumbprint    = ''
        CertificateSubject       = ''
        CertificatePath          = ''
        TargetUserId             = ''
        DownloadPath             = ''
        LocalBackupPath          = ''
        RestorePointUtc          = ''
        LastManifestPath         = ''
        LastDuplicateReportPath  = ''
        LastComparisonReportPath = ''
        LastRecycleBinReportPath = ''
        RecycleBinDownloadPath   = ''
        LastRecycleBinDownloadPath = ''
        DuplicateArchivePath     = ''
        LastDuplicateArchivePath = ''
        LastReconcileReportPath  = ''
        LogRetentionDays         = 30
        UpdatedUtc               = ''
    }
}

function Get-ToolkitConfig {
    <#
    .SYNOPSIS
        Loads config/toolkit-config.json, merging in defaults for missing keys.
    #>
    [CmdletBinding()]
    param([switch]$Refresh)

    $ctx = Initialize-ToolkitEnvironment
    if ($ctx.Config -and -not $Refresh) {
        return $ctx.Config
    }

    $config = New-ToolkitConfigObject

    if (Test-Path -LiteralPath $ctx.ConfigPath) {
        try {
            $raw = Get-Content -LiteralPath $ctx.ConfigPath -Raw -Encoding utf8
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $loaded = $raw | ConvertFrom-Json -ErrorAction Stop
                foreach ($property in $loaded.PSObject.Properties) {
                    $config[$property.Name] = $property.Value
                }
            }
        }
        catch {
            Write-ToolkitLog ("Config file at {0} could not be parsed ({1}); starting from defaults." -f $ctx.ConfigPath, $_.Exception.Message) -Level WARN
        }
    }

    $ctx.Config = $config
    return $config
}

function Save-ToolkitConfig {
    <#
    .SYNOPSIS
        Persists the in-memory config back to config/toolkit-config.json.
    #>
    [CmdletBinding()]
    param($Config)

    $ctx = Initialize-ToolkitEnvironment
    if (-not $Config) { $Config = Get-ToolkitConfig }

    $Config['UpdatedUtc'] = (Get-Date).ToUniversalTime().ToString('o')
    $ctx.Config = $Config

    try {
        ($Config | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $ctx.ConfigPath -Encoding utf8
        Write-ToolkitLog ('Configuration saved to {0}' -f $ctx.ConfigPath) -Level DEBUG -NoConsole
    }
    catch {
        Write-ToolkitLog ('Failed to save configuration: {0}' -f $_.Exception.Message) -Level ERROR
    }

    return $Config
}

function Set-ToolkitConfigValue {
    <#
    .SYNOPSIS
        Sets a single config key and writes the file back out.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][AllowNull()]$Value
    )

    $config = Get-ToolkitConfig
    $config[$Name] = $Value
    return (Save-ToolkitConfig -Config $config)
}

function Get-ToolkitConfigValue {
    <#
    .SYNOPSIS
        Reads a config key, returning a default when unset.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        $Default = ''
    )

    $config = Get-ToolkitConfig
    if ($config.Contains($Name) -and -not [string]::IsNullOrWhiteSpace([string]$config[$Name])) {
        return $config[$Name]
    }
    return $Default
}

# ---------------------------------------------------------------------------
# Console helpers
# ---------------------------------------------------------------------------

function Get-ToolkitReadiness {
    <#
    .SYNOPSIS
        Works out which stages are ready to run, from what the config already holds.
    .DESCRIPTION
        Drives the "needs option N first" hints on the menu, so an operator can see
        what is and is not usable yet without running a stage to find out.
    #>
    [CmdletBinding()]
    param($Config)

    if (-not $Config) { $Config = Get-ToolkitConfig }

    function Test-ConfiguredPath {
        param([string]$Value)
        return (-not [string]::IsNullOrWhiteSpace($Value)) -and (Test-Path -LiteralPath $Value)
    }

    return [pscustomobject]@{
        Registered         = (-not [string]::IsNullOrWhiteSpace([string]$Config['AppId'])) -and
                             (-not [string]::IsNullOrWhiteSpace([string]$Config['CertificateThumbprint']))
        TargetUserSet      = -not [string]::IsNullOrWhiteSpace([string]$Config['TargetUserId'])
        RestorePointSet    = -not [string]::IsNullOrWhiteSpace([string]$Config['RestorePointUtc'])
        HasCloudCopy       = Test-ConfiguredPath -Value ([string]$Config['DownloadPath'])
        HasManifest        = Test-ConfiguredPath -Value ([string]$Config['LastManifestPath'])
        HasLocalBackup     = Test-ConfiguredPath -Value ([string]$Config['LocalBackupPath'])
        HasDuplicateReport = Test-ConfiguredPath -Value ([string]$Config['LastDuplicateReportPath'])
        HasComparison      = Test-ConfiguredPath -Value ([string]$Config['LastComparisonReportPath'])
    }
}

function Write-ToolkitHeader {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Title)

    $bar = '=' * 74
    Write-Host ''
    Write-Host $bar -ForegroundColor DarkCyan
    Write-Host ('  {0}' -f $Title) -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor DarkCyan
}

function Read-ToolkitValue {
    <#
    .SYNOPSIS
        Prompts for a value, offering a remembered default the user can accept.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = '',
        [switch]$AllowEmpty
    )

    while ($true) {
        $suffix = if ([string]::IsNullOrWhiteSpace($Default)) { '' } else { " [$Default]" }
        $answer = Read-Host ('{0}{1}' -f $Prompt, $suffix)

        if ([string]::IsNullOrWhiteSpace($answer)) {
            if (-not [string]::IsNullOrWhiteSpace($Default)) { return $Default }
            if ($AllowEmpty) { return '' }
            Write-Host 'A value is required.' -ForegroundColor Yellow
            continue
        }

        return $answer.Trim()
    }
}

function Confirm-ToolkitAction {
    <#
    .SYNOPSIS
        Yes/no confirmation. Defaults to "no" so destructive steps need an explicit yes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [switch]$DefaultYes
    )

    $hint = if ($DefaultYes) { '[Y/n]' } else { '[y/N]' }
    while ($true) {
        $answer = (Read-Host ('{0} {1}' -f $Prompt, $hint)).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($answer)) { return [bool]$DefaultYes }
        if ($answer -in @('y', 'yes')) { return $true }
        if ($answer -in @('n', 'no')) { return $false }
        Write-Host "Please answer 'y' or 'n'." -ForegroundColor Yellow
    }
}

function Select-ToolkitBatch {
    <#
    .SYNOPSIS
        Optionally narrows a work list to a small trial batch.

    .DESCRIPTION
        Lets an operator run one or two files through a long or irreversible stage,
        check the result, and only then commit to the rest. Returns the whole list
        unchanged when they ask for all of it.

    .OUTPUTS
        The selected items, always as an array.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$Items,
        [string]$Noun = 'file',
        [int]$Limit = 0,
        [switch]$NoPrompt
    )

    $all = @($Items)
    if ($all.Count -eq 0) { return ,@() }

    if ($Limit -le 0 -and -not $NoPrompt) {
        Write-Host ''
        Write-Host ('  TRIAL RUN: you can do one or two {0}s first, check the result, then' -f $Noun) -ForegroundColor Cyan
        Write-Host '  run this option again to do the rest. Nothing is skipped permanently.' -ForegroundColor Gray
        $answer = Read-ToolkitValue -Prompt ('How many {0}s this run? (a number, or ALL for all {1})' -f $Noun, $all.Count) -Default 'ALL'

        if ($answer -notmatch '^(?i)all$') {
            $parsed = 0
            if ([int]::TryParse($answer, [ref]$parsed) -and $parsed -gt 0) {
                $Limit = $parsed
            }
            else {
                Write-Host ('  "{0}" is not a number - doing all {1}.' -f $answer, $all.Count) -ForegroundColor Yellow
            }
        }
    }

    if ($Limit -gt 0 -and $Limit -lt $all.Count) {
        Write-ToolkitLog ('Trial run: taking the first {0} of {1} {2}(s). Re-run for the rest.' -f $Limit, $all.Count, $Noun) -Level WARN
        return ,@($all | Select-Object -First $Limit)
    }

    return ,$all
}

function Read-ToolkitDirectory {
    <#
    .SYNOPSIS
        Prompts for a folder path, validating it and optionally offering to create it.
    .OUTPUTS
        The resolved absolute path, or $null if the user backed out.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = '',
        [switch]$MustExist,
        [switch]$CreateIfMissing
    )

    while ($true) {
        $path = Read-ToolkitValue -Prompt $Prompt -Default $Default
        if ([string]::IsNullOrWhiteSpace($path)) { return $null }

        $path = [Environment]::ExpandEnvironmentVariables($path.Trim('"').Trim())

        if (Test-Path -LiteralPath $path -PathType Container) {
            return (Resolve-Path -LiteralPath $path).ProviderPath
        }

        if (Test-Path -LiteralPath $path -PathType Leaf) {
            Write-Host ('{0} is a file, not a folder.' -f $path) -ForegroundColor Yellow
            continue
        }

        if ($MustExist -and -not $CreateIfMissing) {
            Write-Host ('Folder not found: {0}' -f $path) -ForegroundColor Yellow
            continue
        }

        if ($CreateIfMissing -and (Confirm-ToolkitAction -Prompt ('{0} does not exist. Create it?' -f $path) -DefaultYes)) {
            try {
                New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
                Write-ToolkitLog ('Created folder {0}' -f $path) -Level INFO
                return (Resolve-Path -LiteralPath $path).ProviderPath
            }
            catch {
                Write-Host ('Could not create {0}: {1}' -f $path, $_.Exception.Message) -ForegroundColor Red
                continue
            }
        }

        Write-Host ('Folder not found: {0}' -f $path) -ForegroundColor Yellow
    }
}

function Read-ToolkitDateTime {
    <#
    .SYNOPSIS
        Prompts for a date/time and returns it as UTC.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = ''
    )

    while ($true) {
        $answer = Read-ToolkitValue -Prompt $Prompt -Default $Default
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($answer, [cultureinfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            if ($parsed.Kind -eq [System.DateTimeKind]::Unspecified) {
                $parsed = [datetime]::SpecifyKind($parsed, [System.DateTimeKind]::Local)
            }
            return $parsed.ToUniversalTime()
        }
        Write-Host "Could not read that as a date/time. Try e.g. '2026-05-01 14:30'." -ForegroundColor Yellow
    }
}

function Format-ByteSize {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int64]$Bytes)

    if ($Bytes -lt 0) { return '0 B' }
    $units = @('B', 'KB', 'MB', 'GB', 'TB', 'PB')
    $value = [double]$Bytes
    $index = 0
    while ($value -ge 1024 -and $index -lt ($units.Count - 1)) {
        $value = $value / 1024
        $index++
    }
    if ($index -eq 0) { return ('{0:N0} {1}' -f $value, $units[$index]) }
    return ('{0:N2} {1}' -f $value, $units[$index])
}

function Get-FreeDiskSpace {
    <#
    .SYNOPSIS
        Free bytes on the volume holding a path, or $null if it cannot be determined.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        $full = [System.IO.Path]::GetFullPath($Path)
        $root = [System.IO.Path]::GetPathRoot($full)
        $drive = [System.IO.DriveInfo]::new($root)
        if ($drive.IsReady) { return [int64]$drive.AvailableFreeSpace }
    }
    catch {
        Write-Verbose ('Free space lookup failed for {0}: {1}' -f $Path, $_.Exception.Message)
    }
    return $null
}

# ---------------------------------------------------------------------------
# Path helpers
# ---------------------------------------------------------------------------

function ConvertTo-NormalizedRelativePath {
    <#
    .SYNOPSIS
        Normalises a relative path so local and cloud sides compare cleanly.
    .DESCRIPTION
        Forward slashes, no leading/trailing separator. Comparisons elsewhere are
        done case-insensitively, matching Windows and OneDrive semantics.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return ($Path -replace '\\', '/').Trim('/')
}

function Join-RelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Parent,
        [Parameter(Mandatory)][string]$Child
    )

    $parent = ConvertTo-NormalizedRelativePath -Path $Parent
    if ([string]::IsNullOrWhiteSpace($parent)) { return $Child }
    return ('{0}/{1}' -f $parent, $Child)
}

function Get-SafeLocalPath {
    <#
    .SYNOPSIS
        Turns a drive-relative path into a full local path, refusing traversal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$RelativePath
    )

    $relative = (ConvertTo-NormalizedRelativePath -Path $RelativePath) -replace '/', [System.IO.Path]::DirectorySeparatorChar
    $combined = [System.IO.Path]::GetFullPath((Join-Path $Root $relative))
    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar)

    if (-not $combined.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw ('Refusing to write outside the destination root: {0}' -f $RelativePath)
    }
    return $combined
}

# ---------------------------------------------------------------------------
# Hashing
# ---------------------------------------------------------------------------

function Get-FileSha256 {
    <#
    .SYNOPSIS
        SHA256 of a local file, or $null if it cannot be read (locked, gone, denied).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch {
        Write-ToolkitLog ('Could not hash {0}: {1}' -f $Path, $_.Exception.Message) -Level WARN -NoConsole
        return $null
    }
}

# ---------------------------------------------------------------------------
# Prerequisites
# ---------------------------------------------------------------------------

function Test-ToolkitPrerequisite {
    <#
    .SYNOPSIS
        Checks PowerShell version and the Graph modules each stage needs.
    .OUTPUTS
        [bool] - $true when everything required is present.
    #>
    [CmdletBinding()]
    param([string[]]$RequiredModule = @('Microsoft.Graph.Authentication'))

    $ok = $true

    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-ToolkitLog ('PowerShell 7+ is required; this session is {0}.' -f $PSVersionTable.PSVersion) -Level ERROR
        $ok = $false
    }

    foreach ($name in $RequiredModule) {
        if (-not (Get-Module -ListAvailable -Name $name)) {
            Write-ToolkitLog ("Required module '{0}' is not installed. Run: Install-Module {0} -Scope CurrentUser" -f $name) -Level ERROR
            $ok = $false
        }
    }

    return $ok
}

# ---------------------------------------------------------------------------
# Graph connection
# ---------------------------------------------------------------------------

function Connect-ToolkitGraph {
    <#
    .SYNOPSIS
        App-only certificate connection to Microsoft Graph using saved config.
    .DESCRIPTION
        This is the auth pattern every stage from 2 onwards uses. Stage 1
        (app registration) connects interactively on its own instead.
    .OUTPUTS
        [bool] - $true when connected.
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [switch]$Force
    )

    if (-not (Test-ToolkitPrerequisite)) { return $false }

    $config = Get-ToolkitConfig
    if ([string]::IsNullOrWhiteSpace($TenantId)) { $TenantId = [string]$config['TenantId'] }
    if ([string]::IsNullOrWhiteSpace($AppId)) { $AppId = [string]$config['AppId'] }
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $CertificateThumbprint = [string]$config['CertificateThumbprint'] }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($TenantId)) { $missing += 'TenantId' }
    if ([string]::IsNullOrWhiteSpace($AppId)) { $missing += 'AppId' }
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $missing += 'CertificateThumbprint' }

    if ($missing.Count -gt 0) {
        Write-ToolkitLog ('Missing configuration: {0}. Run menu option 1 (app registration) first.' -f ($missing -join ', ')) -Level ERROR
        return $false
    }

    $ctx = Get-ToolkitContext
    if ($ctx.GraphConnected -and -not $Force -and $ctx.GraphAppId -eq $AppId -and $ctx.GraphTenantId -eq $TenantId) {
        try {
            $current = Get-MgContext
            if ($current -and $current.ClientId -eq $AppId) { return $true }
        }
        catch {
            Write-Verbose ('Graph context check failed: {0}' -f $_.Exception.Message)
        }
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Write-ToolkitLog ('Connecting to Graph as app {0} in tenant {1}...' -f $AppId, $TenantId) -Level INFO
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop

        $ctx.GraphConnected = $true
        $ctx.GraphAppId = $AppId
        $ctx.GraphTenantId = $TenantId
        Write-ToolkitLog 'Connected to Microsoft Graph (application permissions).' -Level SUCCESS
        return $true
    }
    catch {
        $ctx.GraphConnected = $false
        Write-ToolkitLog ('Graph connection failed: {0}' -f $_.Exception.Message) -Level ERROR
        Write-ToolkitLog 'Check that the certificate is installed in this account''s store and that admin consent has been granted.' -Level WARN
        return $false
    }
}

function Get-GraphErrorDetail {
    <#
    .SYNOPSIS
        Pulls status code and Retry-After out of a Graph error, whatever shape it arrives in.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $status = 0
    $retryAfter = $null
    $message = $ErrorRecord.Exception.Message

    $response = $null
    if ($ErrorRecord.Exception.PSObject.Properties['Response']) {
        $response = $ErrorRecord.Exception.Response
    }

    if ($response) {
        if ($response.PSObject.Properties['StatusCode']) {
            try { $status = [int]$response.StatusCode } catch { $status = 0 }
        }
        if ($response.PSObject.Properties['Headers'] -and $response.Headers) {
            try {
                $values = $null
                if ($response.Headers.TryGetValues('Retry-After', [ref]$values) -and $values) {
                    $retryAfter = [int]($values | Select-Object -First 1)
                }
            }
            catch {
                Write-Verbose 'No usable Retry-After header on the response.'
            }
        }
    }

    if ($status -eq 0 -and $ErrorRecord.Exception.PSObject.Properties['StatusCode']) {
        try { $status = [int]$ErrorRecord.Exception.StatusCode } catch { $status = 0 }
    }

    # Last resort: the status often appears in the message text.
    if ($status -eq 0 -and $message -match '\b(429|4\d\d|5\d\d)\b') {
        $status = [int]$Matches[1]
    }

    return [pscustomobject]@{
        StatusCode = $status
        RetryAfter = $retryAfter
        Message    = $message
    }
}

function Invoke-ToolkitGraphRequest {
    <#
    .SYNOPSIS
        Invoke-MgGraphRequest with throttling-aware retry.
    .DESCRIPTION
        Retries 429 (honouring Retry-After), 503/504 and transient socket errors
        with exponential backoff. Anything else is thrown to the caller.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method = 'GET',
        $Body,
        [string]$ContentType = 'application/json',
        [int]$MaxRetry = 6,
        [switch]$AllowNotFound
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Method      = $Method
                Uri         = $Uri
                OutputType  = 'Hashtable'
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) {
                $params['Body'] = $Body
                $params['ContentType'] = $ContentType
            }
            return (Invoke-MgGraphRequest @params)
        }
        catch {
            $detail = Get-GraphErrorDetail -ErrorRecord $_

            if ($AllowNotFound -and $detail.StatusCode -eq 404) { return $null }

            $retryable = $detail.StatusCode -in @(429, 500, 502, 503, 504) -or $detail.StatusCode -eq 0

            if (-not $retryable -or $attempt -gt $MaxRetry) {
                Write-ToolkitLog ('Graph {0} {1} failed (HTTP {2}): {3}' -f $Method, $Uri, $detail.StatusCode, $detail.Message) -Level ERROR -NoConsole
                throw
            }

            $delay = if ($detail.RetryAfter -and $detail.RetryAfter -gt 0) {
                [math]::Min($detail.RetryAfter, 300)
            }
            else {
                [math]::Min([math]::Pow(2, $attempt), 60)
            }

            Write-ToolkitLog ('Graph returned HTTP {0}; retry {1}/{2} in {3}s.' -f $detail.StatusCode, $attempt, $MaxRetry, $delay) -Level WARN -NoConsole
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-ToolkitWebRequest {
    <#
    .SYNOPSIS
        Retry wrapper for pre-authenticated URLs (download URLs, upload sessions).
    .DESCRIPTION
        These URLs carry their own auth token, so no Graph headers are attached.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'PUT', 'POST', 'DELETE')][string]$Method = 'GET',
        [string]$OutFile,
        $Body,
        [hashtable]$Headers,
        [int]$MaxRetry = 5,
        [int]$TimeoutSec = 900
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Uri         = $Uri
                Method      = $Method
                TimeoutSec  = $TimeoutSec
                ErrorAction = 'Stop'
            }
            if ($OutFile) { $params['OutFile'] = $OutFile }
            if ($null -ne $Body) { $params['Body'] = $Body }
            if ($Headers) { $params['Headers'] = $Headers }

            $previous = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest's own bar fights ours
            try {
                return (Invoke-WebRequest @params)
            }
            finally {
                $ProgressPreference = $previous
            }
        }
        catch {
            $detail = Get-GraphErrorDetail -ErrorRecord $_
            $retryable = $detail.StatusCode -in @(429, 500, 502, 503, 504) -or $detail.StatusCode -eq 0

            if (-not $retryable -or $attempt -gt $MaxRetry) { throw }

            $delay = if ($detail.RetryAfter -and $detail.RetryAfter -gt 0) {
                [math]::Min($detail.RetryAfter, 300)
            }
            else {
                [math]::Min([math]::Pow(2, $attempt), 60)
            }

            Write-ToolkitLog ('Transfer returned HTTP {0}; retry {1}/{2} in {3}s.' -f $detail.StatusCode, $attempt, $MaxRetry, $delay) -Level WARN -NoConsole
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-GraphValue {
    <#
    .SYNOPSIS
        Safe key read from a Graph hashtable response.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Item,
        [Parameter(Mandatory)][string]$Key,
        $Default = $null
    )

    if ($null -eq $Item) { return $Default }
    if ($Item -is [System.Collections.IDictionary]) {
        if ($Item.Contains($Key)) { return $Item[$Key] }
        return $Default
    }
    $property = $Item.PSObject.Properties[$Key]
    if ($property) { return $property.Value }
    return $Default
}

# ---------------------------------------------------------------------------
# Drive resolution and enumeration
# ---------------------------------------------------------------------------

function Resolve-OneDriveUser {
    <#
    .SYNOPSIS
        Resolves a UPN or object ID to the user and their OneDrive.
    .OUTPUTS
        PSCustomObject with Id, DisplayName, UserPrincipalName, DriveId, quota fields.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)

    $encoded = [uri]::EscapeDataString($UserId)

    $user = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/users/{0}?$select=id,displayName,userPrincipalName' -f $encoded)
    if (-not $user) { throw ('User not found: {0}' -f $UserId) }

    $drive = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive' -f $encoded)
    if (-not $drive) { throw ('No OneDrive found for {0}. Has the drive been provisioned?' -f $UserId) }

    $quota = Get-GraphValue -Item $drive -Key 'quota'

    return [pscustomobject]@{
        Id                = [string](Get-GraphValue -Item $user -Key 'id')
        DisplayName       = [string](Get-GraphValue -Item $user -Key 'displayName')
        UserPrincipalName = [string](Get-GraphValue -Item $user -Key 'userPrincipalName')
        DriveId           = [string](Get-GraphValue -Item $drive -Key 'id')
        QuotaTotal        = [int64](Get-GraphValue -Item $quota -Key 'total' -Default 0)
        QuotaUsed         = [int64](Get-GraphValue -Item $quota -Key 'used' -Default 0)
        QuotaRemaining    = [int64](Get-GraphValue -Item $quota -Key 'remaining' -Default 0)
    }
}

function ConvertFrom-GraphDriveItem {
    <#
    .SYNOPSIS
        Projects a Graph driveItem down to the slim shape the stages work with.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ParentPath
    )

    $name = [string](Get-GraphValue -Item $Item -Key 'name')
    $file = Get-GraphValue -Item $Item -Key 'file'
    $folder = Get-GraphValue -Item $Item -Key 'folder'
    $fsInfo = Get-GraphValue -Item $Item -Key 'fileSystemInfo'
    $hashes = if ($file) { Get-GraphValue -Item $file -Key 'hashes' } else { $null }

    return [pscustomobject]@{
        Id             = [string](Get-GraphValue -Item $Item -Key 'id')
        Name           = $name
        RelativePath   = Join-RelativePath -Parent $ParentPath -Child $name
        ParentPath     = ConvertTo-NormalizedRelativePath -Path $ParentPath
        IsFolder       = [bool]$folder
        Size           = [int64](Get-GraphValue -Item $Item -Key 'size' -Default 0)
        ChildCount     = if ($folder) { [int](Get-GraphValue -Item $folder -Key 'childCount' -Default 0) } else { 0 }
        CreatedUtc     = [string](Get-GraphValue -Item $fsInfo -Key 'createdDateTime')
        ModifiedUtc    = [string](Get-GraphValue -Item $fsInfo -Key 'lastModifiedDateTime')
        QuickXorHash   = [string](Get-GraphValue -Item $hashes -Key 'quickXorHash')
        Sha256Hash     = [string](Get-GraphValue -Item $hashes -Key 'sha256Hash')
        DownloadUrl    = [string](Get-GraphValue -Item $Item -Key '@microsoft.graph.downloadUrl')
    }
}

function Get-OneDriveItemInventory {
    <#
    .SYNOPSIS
        Walks a user's whole OneDrive, following @odata.nextLink at every level.
    .DESCRIPTION
        Breadth-first so folders are always emitted before their contents - the
        download stage relies on that to create directories in order. Shared by
        stage 2 (download) and stage 3 (duplicate scan).
    .OUTPUTS
        A stream of slim drive item objects (see ConvertFrom-GraphDriveItem).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [switch]$FilesOnly,
        [string]$Activity = 'Enumerating OneDrive'
    )

    $encoded = [uri]::EscapeDataString($UserId)
    $queue = [System.Collections.Generic.Queue[psobject]]::new()
    $queue.Enqueue([pscustomobject]@{ Uri = ('https://graph.microsoft.com/v1.0/users/{0}/drive/root/children?$top=200' -f $encoded); Path = '' })

    $fileCount = 0
    $folderCount = 0

    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        $uri = $node.Uri

        while ($uri) {
            $response = Invoke-ToolkitGraphRequest -Uri $uri
            $values = Get-GraphValue -Item $response -Key 'value' -Default @()

            foreach ($raw in $values) {
                $item = ConvertFrom-GraphDriveItem -Item $raw -ParentPath $node.Path

                if ($item.IsFolder) {
                    $folderCount++
                    $queue.Enqueue([pscustomobject]@{
                        Uri  = ('https://graph.microsoft.com/v1.0/users/{0}/drive/items/{1}/children?$top=200' -f $encoded, $item.Id)
                        Path = $item.RelativePath
                    })
                    if (-not $FilesOnly) { $item }
                }
                else {
                    $fileCount++
                    $item
                }
            }

            if (($fileCount + $folderCount) % 100 -eq 0) {
                Write-Progress -Activity $Activity -Status ('{0} files / {1} folders found, {2} folders queued' -f $fileCount, $folderCount, $queue.Count)
            }

            $uri = [string](Get-GraphValue -Item $response -Key '@odata.nextLink' -Default '')
        }
    }

    Write-Progress -Activity $Activity -Completed
    Write-ToolkitLog ('Enumeration complete: {0} files in {1} folders.' -f $fileCount, $folderCount) -Level INFO -NoConsole
}

function Get-OneDriveItemById {
    <#
    .SYNOPSIS
        Re-reads one drive item - used to refresh an expired download URL.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$ItemId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ParentPath
    )

    $encoded = [uri]::EscapeDataString($UserId)
    $raw = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive/items/{1}' -f $encoded, $ItemId) -AllowNotFound
    if (-not $raw) { return $null }
    return (ConvertFrom-GraphDriveItem -Item $raw -ParentPath $ParentPath)
}

# ---------------------------------------------------------------------------
# Sync-conflict filename patterns
# ---------------------------------------------------------------------------

function Get-ConflictNameInfo {
    <#
    .SYNOPSIS
        Tests a filename against the common OneDrive/Windows conflict-copy patterns.

    .DESCRIPTION
        Two tiers of confidence, because precision matters here - a false positive
        means a legitimate file silently never gets re-uploaded:

          High   - '(1)', '- Copy', 'conflicted copy'. Distinctive enough to act on
                   by themselves.
          Medium - the 'name-MACHINE' / 'name-user-MACHINE' shape OneDrive uses for
                   sync conflicts. This also matches ordinary hyphenated filenames
                   ('annual-report.docx'), so callers must only treat these as
                   conflicts when a file with the base name actually exists
                   alongside them - see RequiresOriginal.

    .OUTPUTS
        PSCustomObject: IsCandidate, BaseName, PatternType, Confidence, RequiresOriginal, Marker
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)

    $result = [pscustomobject]@{
        IsCandidate      = $false
        BaseName         = $Name
        PatternType      = 'None'
        Confidence       = 'None'
        RequiresOriginal = $false
        Marker           = ''
    }

    if ([string]::IsNullOrWhiteSpace($Name)) { return $result }

    $extension = [System.IO.Path]::GetExtension($Name)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    if ([string]::IsNullOrWhiteSpace($stem)) { return $result }

    $patterns = @(
        @{ Type = 'ConflictedCopy'; Confidence = 'High';   Regex = '^(?<base>.+?)[\s_-]*\((?<marker>[^)]*conflicted copy[^)]*)\)\s*$' }
        # The dashed form only: "report - Copy.docx", "report-Copy.docx",
        # "report - Copy (2).docx", "report - Copy 2.docx". A separator is required.
        # The space-only form ("report copy.docx") is deliberately NOT matched: the
        # dash is what makes the name unambiguous, and without it plenty of ordinary
        # documents ("Certified copy.pdf") would be caught by mistake.
        @{ Type = 'CopySuffix';     Confidence = 'High';   Regex = '^(?<base>.+?)\s*[-_]\s*(?<marker>Copy(?:\s*\(?\d{1,4}\)?)?)\s*$' }
        @{ Type = 'NumberedCopy';   Confidence = 'High';   Regex = '^(?<base>.+?)\s*\((?<marker>\d{1,4})\)\s*$' }
        @{ Type = 'UserMachineSuffix'; Confidence = 'Medium'; Regex = '^(?<base>.+?)-(?<marker>[A-Za-z0-9._'']+-[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*)$' }
        @{ Type = 'MachineSuffix';  Confidence = 'Medium'; Regex = '^(?<base>.+?)-(?<marker>[A-Za-z0-9][A-Za-z0-9]*(?:-[A-Za-z0-9]+)*)$' }
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($stem, $pattern.Regex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $match.Success) { continue }

        $base = $match.Groups['base'].Value.TrimEnd()
        if ([string]::IsNullOrWhiteSpace($base)) { continue }

        $result.IsCandidate = $true
        $result.BaseName = '{0}{1}' -f $base, $extension
        $result.PatternType = $pattern.Type
        $result.Confidence = $pattern.Confidence
        $result.RequiresOriginal = ($pattern.Confidence -eq 'Medium')
        $result.Marker = $match.Groups['marker'].Value
        return $result
    }

    return $result
}

function Test-MachineLikeToken {
    <#
    .SYNOPSIS
        Heuristic: does this suffix token look like a computer name?
    .DESCRIPTION
        Used only to rank medium-confidence matches in the scanner report;
        it never on its own decides that a file is a conflict copy.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token)) { return $false }
    if ($Token -match '^(?i)(desktop|laptop|surface|win|pc|book|imac|mac)') { return $true }
    if ($Token -cmatch '^[A-Z0-9][A-Z0-9-]{2,}$') { return $true }
    if ($Token -match '\d' -and $Token.Length -ge 5) { return $true }
    return $false
}

Export-ModuleMember -Function @(
    'Get-ToolkitContext'
    'Initialize-ToolkitEnvironment'
    'Get-ToolkitReportPath'
    'Write-ToolkitLog'
    'Remove-OldToolkitLog'
    'Get-ToolkitConfig'
    'Save-ToolkitConfig'
    'Set-ToolkitConfigValue'
    'Get-ToolkitConfigValue'
    'Get-ToolkitReadiness'
    'Write-ToolkitHeader'
    'Read-ToolkitValue'
    'Confirm-ToolkitAction'
    'Select-ToolkitBatch'
    'Read-ToolkitDirectory'
    'Read-ToolkitDateTime'
    'Format-ByteSize'
    'Get-FreeDiskSpace'
    'ConvertTo-NormalizedRelativePath'
    'Join-RelativePath'
    'Get-SafeLocalPath'
    'Get-FileSha256'
    'Test-ToolkitPrerequisite'
    'Connect-ToolkitGraph'
    'Get-GraphErrorDetail'
    'Invoke-ToolkitGraphRequest'
    'Invoke-ToolkitWebRequest'
    'Get-GraphValue'
    'Resolve-OneDriveUser'
    'ConvertFrom-GraphDriveItem'
    'Get-OneDriveItemInventory'
    'Get-OneDriveItemById'
    'Get-ConflictNameInfo'
    'Test-MachineLikeToken'
)
