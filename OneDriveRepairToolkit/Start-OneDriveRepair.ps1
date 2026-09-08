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

    $restorePoint = 'not set yet'
    if (-not [string]::IsNullOrWhiteSpace([string]$config['RestorePointUtc'])) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse([string]$config['RestorePointUtc'], [ref]$parsed)) {
            $restorePoint = '{0:yyyy-MM-dd HH:mm} UTC' -f $parsed.ToUniversalTime()
        }
    }

    function Format-Setting {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return 'not set yet' }
        return $Value
    }

    Write-Host ''
    Write-Host ('-' * 78) -ForegroundColor DarkGray
    Write-Host '  Remembered from last time (every option lets you change these):' -ForegroundColor DarkCyan
    Write-Host ('    User being repaired : {0}' -f (Format-Setting ([string]$config['TargetUserId'])))
    Write-Host ('    Rolled back to      : {0}' -f $restorePoint)
    Write-Host ('    Cloud copy saved in : {0}' -f (Format-Setting ([string]$config['DownloadPath'])))
    Write-Host ('    PC backup folder    : {0}' -f (Format-Setting ([string]$config['LocalBackupPath'])))
    Write-Host ('    Log for this run    : {0}' -f (Get-ToolkitContext).LogPath) -ForegroundColor DarkGray
}

function Write-MenuItem {
    <#
    .SYNOPSIS
        Renders one menu entry: number, title, effect tag, description, and any blocker.
    .DESCRIPTION
        The tag is the important part. An operator should never have to guess whether
        an option only reads, writes to this PC, or can change the user's OneDrive.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Number,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][ValidateSet('safe', 'local', 'onedrive', 'both', 'setup')][string]$Effect,
        [Parameter(Mandatory)][string]$Description,
        [string]$Note,
        [string]$Blocker
    )

    $tag = switch ($Effect) {
        'safe'     { '[reads only]' }
        'local'    { '[writes to this PC]' }
        'onedrive' { '[CHANGES ONEDRIVE]' }
        'both'     { '[this PC + CAN CHANGE ONEDRIVE]' }
        'setup'    { '[one-time setup]' }
    }
    $tagColour = switch ($Effect) {
        'onedrive' { 'Red' }
        'both'     { 'Yellow' }
        default    { 'DarkGray' }
    }

    # Title left, tag right-aligned to column 78.
    $left = '  {0,2}  {1}' -f $Number, $Title
    $pad = [math]::Max(1, 78 - $left.Length - $tag.Length)
    Write-Host $left -NoNewline -ForegroundColor White
    Write-Host (' ' * $pad) -NoNewline
    Write-Host $tag -ForegroundColor $tagColour

    Write-Host ('       {0}' -f $Description) -ForegroundColor Gray
    if ($Note) { Write-Host ('       {0}' -f $Note) -ForegroundColor DarkGray }
    if ($Blocker) { Write-Host ('       -> {0}' -f $Blocker) -ForegroundColor Yellow }
}

function Show-ToolkitMenu {
    [CmdletBinding()]
    param()

    $ready = Get-ToolkitReadiness
    $config = Get-ToolkitConfig

    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host '  OneDrive Sync Repair & Reconciliation Toolkit' -ForegroundColor Cyan
    Write-Host '  The drive rollback itself is done by hand in the OneDrive admin centre.' -ForegroundColor DarkGray
    Write-Host '  This toolkit does everything that comes after it.' -ForegroundColor DarkGray
    Write-Host ('=' * 78) -ForegroundColor DarkCyan

    Write-Host ''
    Write-Host '  SETUP' -ForegroundColor Cyan
    Write-MenuItem -Number '1' -Effect 'setup' `
        -Title 'Register Azure AD app and certificate' `
        -Description 'Creates the app-only login that options 2-8 use. Run once, as an admin.' `
        -Note $(if ($ready.Registered) { 'Already done: app {0}' -f $config['AppId'] } else { 'Not done yet - start here.' })

    Write-Host ''
    Write-Host '  STEP 1 - SEE WHAT IS ACTUALLY IN THE CLOUD' -ForegroundColor Cyan
    Write-MenuItem -Number '2' -Effect 'local' `
        -Title 'Download a copy of the user''s OneDrive' `
        -Description 'Copies the whole drive to a folder you choose, and lists what it found.' `
        -Note 'Run this after the rollback. Nothing on OneDrive is changed.' `
        -Blocker $(if (-not $ready.Registered) { 'Needs option 1 first.' })

    Write-MenuItem -Number '3' -Effect 'both' `
        -Title 'Find duplicate and sync-conflict files' `
        -Description 'Finds copies like "report (1).docx" and lists them for you.' `
        -Note 'The scan changes nothing. Deleting is a separate, confirmed step.' `
        -Blocker $(if (-not $ready.Registered) { 'Needs option 1 first.' })

    Write-Host ''
    Write-Host '  STEP 2 - WORK OUT WHAT THE ROLLBACK LOST' -ForegroundColor Cyan
    Write-MenuItem -Number '4' -Effect 'local' `
        -Title 'Compare the cloud copy against the PC backup' `
        -Description 'Finds the real work the rollback removed, by comparing the two.' `
        -Note 'Reads both folders and writes a report. Nothing is uploaded.' `
        -Blocker $(if (-not $ready.HasCloudCopy) { 'Needs option 2 first - no downloaded copy on record.' })

    Write-MenuItem -Number '5' -Effect 'onedrive' `
        -Title 'Put the recovered files back' `
        -Description 'Uploads ONLY the files option 4 confirmed as genuine missing work.' `
        -Note 'Offers a dry run first. Never uploads the whole backup.' `
        -Blocker $(if (-not $ready.HasComparison) { 'Needs option 4 first - no comparison report on record.' })

    Write-Host ''
    Write-Host '  RECOVERING DELETED FILES AND KEEPING BACKUPS' -ForegroundColor Cyan
    Write-MenuItem -Number '6' -Effect 'safe' `
        -Title 'List what is in the recycle bin' `
        -Description 'Shows what was deleted and flags what looks worth getting back.' `
        -Note 'Read-only. Changes nothing, anywhere.' `
        -Blocker $(if (-not $ready.Registered) { 'Needs option 1 first.' })

    Write-MenuItem -Number '7' -Effect 'both' `
        -Title 'Download the files in the recycle bin' `
        -Description 'Saves deleted files to a folder on this PC.' `
        -Note 'Restores each file, downloads it, deletes it again - one at a time.' `
        -Blocker $(if (-not $ready.Registered) { 'Needs option 1 first.' })

    Write-MenuItem -Number '8' -Effect 'both' `
        -Title 'Back up the duplicate copies to a folder' `
        -Description 'Downloads every copy found by option 3 - e.g. onto a USB stick.' `
        -Note 'Checks each arrived intact, then offers to delete only those.' `
        -Blocker $(if (-not $ready.HasDuplicateReport) { 'Needs option 3 first - no scan report on record.' })

    Write-Host ''
    Write-Host '  OTHER' -ForegroundColor Cyan
    Write-MenuItem -Number '9' -Effect 'safe' `
        -Title 'View the log from this run' `
        -Description 'Shows what the toolkit has done, newest last.'
    Write-Host '  10  Exit' -ForegroundColor White
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
