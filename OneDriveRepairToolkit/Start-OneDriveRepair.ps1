#Requires -Version 7.0
<#
.SYNOPSIS
    OneDrive Sync Repair & Reconciliation Toolkit - text menu entry point.

.DESCRIPTION
    Wraps the four automated stages of the repair. The whole-drive rollback
    itself is done manually in the SharePoint/OneDrive admin centre; this
    toolkit picks up from the state that rollback leaves behind:

      2. Download a fresh cloud copy of the rolled-back drive (the trustworthy
         comparison source - neither device's sync folder can be trusted).
      3. Scan for sync-conflict duplicates, and optionally clean them up.
      4. Compare that download against the local PC backup to find the
         legitimate work done since the restore point.
      5. Upload just those canonical files back, without recreating duplicates.
      6. Inventory the recycle bin to see what was deleted and what is worth
         recovering by hand (optional, run at any point).
      7. Download the recycle bin's contents, by restoring each item, downloading
         it, and putting it back (optional; needs the SharePoint permission).
      8. Archive the duplicate copies to a local folder, verify each one, then
         optionally delete the verified ones (optional, run at any point).

    This script is a thin dispatcher - all the logic lives in modules/.

.PARAMETER Stage
    Runs a single stage and exits, for scripted or repeat runs. Omit for the menu.

.PARAMETER ConfigPath
    Overrides the default config location (config/toolkit-config.json).

.EXAMPLE
    ./Start-OneDriveRepair.ps1

.EXAMPLE
    ./Start-OneDriveRepair.ps1 -Stage 2
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 9)][int]$Stage,
    [string]$ConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$moduleRoot = Join-Path $PSScriptRoot 'modules'
foreach ($module in @('Common', 'AppRegistration', 'DownloadOneDrive', 'DuplicateScanner', 'CompareDrives', 'ReconcileUpload', 'RecycleBin')) {
    Import-Module (Join-Path $moduleRoot ('{0}.psm1' -f $module)) -Force -DisableNameChecking
}

$context = Initialize-ToolkitEnvironment
if ($ConfigPath) { $context.ConfigPath = $ConfigPath }

Remove-OldToolkitLog -RetentionDays ([int](Get-ToolkitConfigValue -Name 'LogRetentionDays' -Default 30))
Write-ToolkitLog ('=== Toolkit started (PowerShell {0} on {1}) ===' -f $PSVersionTable.PSVersion, [System.Environment]::OSVersion.VersionString) -Level INFO -NoConsole

function Show-ToolkitStatus {
    <#
    .SYNOPSIS
        Prints what the toolkit already knows, so the operator can see where they are.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ToolkitConfig

    $registered = -not [string]::IsNullOrWhiteSpace([string]$config['AppId']) -and
                  -not [string]::IsNullOrWhiteSpace([string]$config['CertificateThumbprint'])

    $restorePoint = 'not set'
    if (-not [string]::IsNullOrWhiteSpace([string]$config['RestorePointUtc'])) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse([string]$config['RestorePointUtc'], [ref]$parsed)) {
            $restorePoint = '{0:yyyy-MM-dd HH:mm} UTC' -f $parsed.ToUniversalTime()
        }
    }

    Write-Host ''
    Write-Host '  Current configuration' -ForegroundColor DarkCyan
    Write-Host ('    App registration : {0}' -f $(if ($registered) { 'configured (app ' + $config['AppId'] + ')' } else { 'not configured - run option 1' })) `
        -ForegroundColor $(if ($registered) { 'Gray' } else { 'Yellow' })
    Write-Host ('    Target user      : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$config['TargetUserId'])) { 'not set' } else { $config['TargetUserId'] }))
    Write-Host ('    Restore point    : {0}' -f $restorePoint)
    Write-Host ('    Cloud copy       : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$config['DownloadPath'])) { 'not set' } else { $config['DownloadPath'] }))
    Write-Host ('    Local backup     : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$config['LocalBackupPath'])) { 'not set' } else { $config['LocalBackupPath'] }))
    Write-Host ('    Log file         : {0}' -f (Get-ToolkitContext).LogPath)
}

function Show-ToolkitMenu {
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host '================================================================' -ForegroundColor DarkCyan
    Write-Host '   OneDrive Sync Repair & Reconciliation Toolkit' -ForegroundColor Cyan
    Write-Host '================================================================' -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host '   1. Register Azure AD App & Generate Certificate (one-time)'
    Write-Host '   2. Download Cloud Copy of User''s OneDrive (post-rollback)'
    Write-Host '   3. Scan OneDrive for Duplicate/Conflict Files'
    Write-Host '   4. Compare Downloaded OneDrive Copy vs. Local PC Backup'
    Write-Host '   5. Reconcile - Upload Missing/Newer Canonical Files'
    Write-Host '   6. Inventory the OneDrive Recycle Bin'
    Write-Host '   7. Download Recycle Bin Contents (restores, downloads, puts back)'
    Write-Host '   8. Archive Duplicate Copies - Download, then Optionally Delete'
    Write-Host '   9. View Last Run Log'
    Write-Host '  10. Exit'
}

function Show-LastRunLog {
    <#
    .SYNOPSIS
        Menu option 6 - tail the current day's log.
    #>
    [CmdletBinding()]
    param([int]$Lines = 60)

    $ctx = Get-ToolkitContext
    Write-ToolkitHeader 'Last run log'

    $logFile = $ctx.LogPath
    if (-not (Test-Path -LiteralPath $logFile)) {
        $logFile = Get-ChildItem -LiteralPath $ctx.LogDir -Filter 'toolkit-*.log' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    }

    if (-not $logFile -or -not (Test-Path -LiteralPath $logFile)) {
        Write-Host '  No log files yet.' -ForegroundColor Yellow
        return
    }

    Write-Host ('  {0}' -f $logFile) -ForegroundColor DarkGray
    Write-Host ''

    $requested = Read-ToolkitValue -Prompt 'How many lines from the end?' -Default ([string]$Lines)
    $count = $Lines
    [void][int]::TryParse($requested, [ref]$count)
    if ($count -le 0) { $count = $Lines }

    Get-Content -LiteralPath $logFile -Tail $count | ForEach-Object {
        $colour = if ($_ -match '\[ERROR' ) { 'Red' }
                  elseif ($_ -match '\[WARN' ) { 'Yellow' }
                  elseif ($_ -match '\[SUCCESS') { 'Green' }
                  else { 'Gray' }
        Write-Host $_ -ForegroundColor $colour
    }
}

function Invoke-ToolkitStage {
    <#
    .SYNOPSIS
        Dispatches one menu choice to its module function.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$Choice)

    switch ($Choice) {
        1 { Invoke-AppRegistrationSetup | Out-Null }
        2 { Invoke-OneDriveDownload | Out-Null }
        3 { Invoke-DuplicateScan | Out-Null }
        4 { Invoke-DriveComparison | Out-Null }
        5 { Invoke-Reconciliation | Out-Null }
        6 { Invoke-RecycleBinInventory | Out-Null }
        7 { Invoke-RecycleBinDownload | Out-Null }
        8 { Invoke-DuplicateArchiveFromReport | Out-Null }
        9 { Show-LastRunLog }
        default { Write-Host '  Not a valid choice.' -ForegroundColor Yellow }
    }
}

# --- Single-stage mode -------------------------------------------------------

if ($PSBoundParameters.ContainsKey('Stage')) {
    try {
        Invoke-ToolkitStage -Choice $Stage
    }
    catch {
        Write-ToolkitLog ('Stage {0} failed: {1}' -f $Stage, $_.Exception.Message) -Level ERROR
        Write-ToolkitLog $_.ScriptStackTrace -Level DEBUG -NoConsole
        exit 1
    }
    exit 0
}

# --- Menu loop ---------------------------------------------------------------

while ($true) {
    Show-ToolkitMenu
    Show-ToolkitStatus
    Write-Host ''

    $choice = (Read-Host '   Choose an option [1-10]').Trim()

    if ($choice -in @('10', 'q', 'Q', 'exit')) {
        Write-ToolkitLog '=== Toolkit exited ===' -Level INFO -NoConsole
        Write-Host ''
        Write-Host '  Goodbye.' -ForegroundColor Cyan
        break
    }

    $number = 0
    if (-not [int]::TryParse($choice, [ref]$number) -or $number -lt 1 -or $number -gt 9) {
        Write-Host '  Please choose a number between 1 and 10.' -ForegroundColor Yellow
        continue
    }

    try {
        Write-ToolkitLog ('--- Menu option {0} selected ---' -f $number) -Level INFO -NoConsole
        Invoke-ToolkitStage -Choice $number
    }
    catch {
        Write-ToolkitLog ('Option {0} failed: {1}' -f $number, $_.Exception.Message) -Level ERROR
        Write-ToolkitLog $_.ScriptStackTrace -Level DEBUG -NoConsole
    }

    Write-Host ''
    Read-Host '   Press Enter to return to the menu' | Out-Null
}
