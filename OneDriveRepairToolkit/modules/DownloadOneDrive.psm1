#Requires -Version 7.0
<#
.SYNOPSIS
    Stage 2 - download a fresh cloud copy of the user's OneDrive.

.DESCRIPTION
    Run after the admin-centre rollback. The device with sync problems cannot be
    trusted as a comparison source, so this pulls the drive straight from Graph:
    that copy is the ground truth for "what is actually in the cloud right now"
    and forms the cloud-side half of the stage 4 diff.

    Re-runnable at any time for a fresh snapshot (for example after cleaning up
    duplicates in stage 3).
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking

function ConvertTo-UtcDateTime {
    <#
    .SYNOPSIS
        Parses a Graph ISO-8601 timestamp, returning $null when unusable.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    if ([datetime]::TryParse($Value, [cultureinfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Set-LocalFileTimestamp {
    <#
    .SYNOPSIS
        Stamps the downloaded file with the drive item's original times.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()][AllowEmptyString()][string]$CreatedUtc,
        [AllowNull()][AllowEmptyString()][string]$ModifiedUtc
    )

    try {
        $created = ConvertTo-UtcDateTime -Value $CreatedUtc
        if ($created) { [System.IO.File]::SetCreationTimeUtc($Path, $created) }

        $modified = ConvertTo-UtcDateTime -Value $ModifiedUtc
        if ($modified) { [System.IO.File]::SetLastWriteTimeUtc($Path, $modified) }
    }
    catch {
        Write-ToolkitLog ('Could not stamp timestamps on {0}: {1}' -f $Path, $_.Exception.Message) -Level WARN -NoConsole
    }
}

function Test-LocalCopyCurrent {
    <#
    .SYNOPSIS
        True when an existing local file already matches the cloud item.
    .DESCRIPTION
        Size plus last-write time to the second. Lets an interrupted download be
        resumed without re-fetching everything.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Item
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }

    $info = Get-Item -LiteralPath $Path
    if ($info.Length -ne $Item.Size) { return $false }

    $modified = ConvertTo-UtcDateTime -Value $Item.ModifiedUtc
    if (-not $modified) { return $false }

    return ([math]::Abs(($info.LastWriteTimeUtc - $modified).TotalSeconds) -le 2)
}

function Invoke-DriveItemDownload {
    <#
    .SYNOPSIS
        Downloads one drive item, refreshing the download URL if it has expired.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][string]$TargetPath,
        [Parameter(Mandatory)][string]$UserId
    )

    $directory = Split-Path -Parent $TargetPath
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $partial = '{0}.partial' -f $TargetPath
    $url = $Item.DownloadUrl

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        if ([string]::IsNullOrWhiteSpace($url)) {
            # Pre-authenticated download URLs expire after about an hour, which a
            # long enumeration can outlive. Re-read the item for a fresh one.
            $refreshed = Get-OneDriveItemById -UserId $UserId -ItemId $Item.Id -ParentPath $Item.ParentPath
            if (-not $refreshed -or [string]::IsNullOrWhiteSpace($refreshed.DownloadUrl)) {
                throw ('No download URL available for {0}.' -f $Item.RelativePath)
            }
            $url = $refreshed.DownloadUrl
        }

        try {
            Invoke-ToolkitWebRequest -Uri $url -Method GET -OutFile $partial | Out-Null
            Move-Item -LiteralPath $partial -Destination $TargetPath -Force
            Set-LocalFileTimestamp -Path $TargetPath -CreatedUtc $Item.CreatedUtc -ModifiedUtc $Item.ModifiedUtc
            return
        }
        catch {
            if (Test-Path -LiteralPath $partial) {
                Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
            }
            if ($attempt -ge 2) { throw }
            Write-ToolkitLog ('Download of {0} failed ({1}); refreshing the download URL and retrying.' -f $Item.RelativePath, $_.Exception.Message) -Level WARN -NoConsole
            $url = $null
        }
    }
}

function Invoke-OneDriveDownload {
    <#
    .SYNOPSIS
        Menu option 2 - download the post-rollback cloud copy and write a manifest.
    #>
    [CmdletBinding()]
    param(
        [string]$UserId,
        [string]$Destination,
        [switch]$NoPrompt
    )

    Write-ToolkitHeader 'Stage 2 - Download Cloud Copy of OneDrive (post-rollback)'

    if (-not (Connect-ToolkitGraph)) { return $null }

    $config = Get-ToolkitConfig

    if (-not $UserId) {
        $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
    }

    try {
        $user = Resolve-OneDriveUser -UserId $UserId
    }
    catch {
        Write-ToolkitLog ('Could not resolve the user or their drive: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    Write-Host ''
    Write-Host ('  User        : {0} <{1}>' -f $user.DisplayName, $user.UserPrincipalName)
    Write-Host ('  Drive ID    : {0}' -f $user.DriveId)
    Write-Host ('  Quota used  : {0} of {1}' -f (Format-ByteSize $user.QuotaUsed), (Format-ByteSize $user.QuotaTotal))
    Write-Host ''

    if (-not $Destination) {
        $Destination = Read-ToolkitDirectory `
            -Prompt 'Local destination folder for the download' `
            -Default ([string]$config['DownloadPath']) `
            -CreateIfMissing
    }
    if (-not $Destination) {
        Write-ToolkitLog 'No destination chosen; download cancelled.' -Level WARN
        return $null
    }

    # Each snapshot lands in its own folder so an earlier one is never overwritten.
    if (-not $NoPrompt -and (Confirm-ToolkitAction -Prompt 'Create a timestamped subfolder for this snapshot?' -DefaultYes)) {
        $leaf = '{0}-{1}' -f ($user.UserPrincipalName -replace '[^A-Za-z0-9._-]', '_'), (Get-Date -Format 'yyyyMMdd-HHmmss')
        $Destination = Join-Path $Destination $leaf
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    }

    $freeSpace = Get-FreeDiskSpace -Path $Destination
    if ($null -ne $freeSpace) {
        Write-Host ('  Free space  : {0} at {1}' -f (Format-ByteSize $freeSpace), $Destination)
        if ($user.QuotaUsed -gt 0 -and $freeSpace -lt $user.QuotaUsed) {
            Write-ToolkitLog ('The drive holds {0} but only {1} is free at the destination.' -f (Format-ByteSize $user.QuotaUsed), (Format-ByteSize $freeSpace)) -Level WARN
            if (-not (Confirm-ToolkitAction -Prompt 'Continue anyway?')) {
                Write-ToolkitLog 'Download cancelled - not enough free disk space.' -Level WARN
                return $null
            }
        }
    }
    else {
        Write-ToolkitLog 'Free disk space could not be determined for the destination; continuing.' -Level WARN
    }

    $config['TargetUserId'] = $UserId
    $config['DownloadPath'] = $Destination
    Save-ToolkitConfig -Config $config | Out-Null

    Write-ToolkitLog ('Downloading {0} to {1}' -f $user.UserPrincipalName, $Destination) -Level INFO

    # Pass 1: enumerate, so pass 2 can show real progress against a known total.
    Write-Host ''
    Write-ToolkitLog 'Enumerating the drive (this can take a while on large drives)...' -Level INFO
    $inventory = [System.Collections.Generic.List[psobject]]::new()
    try {
        Get-OneDriveItemInventory -UserId $UserId -Activity 'Stage 2: enumerating OneDrive' |
            ForEach-Object { $inventory.Add($_) }
    }
    catch {
        Write-ToolkitLog ('Enumeration failed: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    $folders = @($inventory | Where-Object { $_.IsFolder })
    $files = @($inventory | Where-Object { -not $_.IsFolder })
    $totalBytes = ($files | Measure-Object -Property Size -Sum).Sum
    if (-not $totalBytes) { $totalBytes = 0 }

    Write-ToolkitLog ('Found {0} files ({1}) in {2} folders.' -f $files.Count, (Format-ByteSize $totalBytes), $folders.Count) -Level INFO

    # Recreate the folder structure first, including folders that hold no files.
    foreach ($folder in $folders) {
        try {
            $path = Get-SafeLocalPath -Root $Destination -RelativePath $folder.RelativePath
            if (-not (Test-Path -LiteralPath $path)) {
                New-Item -ItemType Directory -Path $path -Force | Out-Null
            }
        }
        catch {
            Write-ToolkitLog ('Could not create folder {0}: {1}' -f $folder.RelativePath, $_.Exception.Message) -Level WARN
        }
    }

    $manifest = [System.Collections.Generic.List[psobject]]::new()
    $downloaded = 0
    $skipped = 0
    $failed = 0
    $bytesDone = [int64]0
    $index = 0

    foreach ($file in $files) {
        $index++
        $percent = if ($files.Count -gt 0) { [int](($index / $files.Count) * 100) } else { 100 }
        Write-Progress -Activity 'Stage 2: downloading OneDrive' `
            -Status ('{0}/{1} - {2} of {3} - {4}' -f $index, $files.Count, (Format-ByteSize $bytesDone), (Format-ByteSize $totalBytes), $file.RelativePath) `
            -PercentComplete ([math]::Min($percent, 100))

        $status = 'Downloaded'
        $localPath = $null
        $hash = $null

        try {
            $localPath = Get-SafeLocalPath -Root $Destination -RelativePath $file.RelativePath

            if (Test-LocalCopyCurrent -Path $localPath -Item $file) {
                $status = 'AlreadyCurrent'
                $skipped++
            }
            else {
                Invoke-DriveItemDownload -Item $file -TargetPath $localPath -UserId $UserId
                $downloaded++
            }

            $hash = Get-FileSha256 -Path $localPath
            $bytesDone += $file.Size
        }
        catch {
            $status = 'Failed'
            $failed++
            Write-ToolkitLog ('FAILED {0}: {1}' -f $file.RelativePath, $_.Exception.Message) -Level ERROR -NoConsole
        }

        $manifest.Add([pscustomobject]@{
            RelativePath = $file.RelativePath
            Name         = $file.Name
            SizeBytes    = $file.Size
            Sha256       = $hash
            CreatedUtc   = $file.CreatedUtc
            ModifiedUtc  = $file.ModifiedUtc
            ItemId       = $file.Id
            QuickXorHash = $file.QuickXorHash
            LocalPath    = $localPath
            Status       = $status
        })
    }

    Write-Progress -Activity 'Stage 2: downloading OneDrive' -Completed

    $manifestCsv = Get-ToolkitReportPath -BaseName 'cloud-manifest' -Extension 'csv'
    $manifestJson = [System.IO.Path]::ChangeExtension($manifestCsv, 'json')

    try {
        $manifest | Export-Csv -LiteralPath $manifestCsv -NoTypeInformation -Encoding utf8
        $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestJson -Encoding utf8
        # A second copy inside the snapshot keeps it self-describing if it is moved.
        $manifest | Export-Csv -LiteralPath (Join-Path $Destination '_cloud-manifest.csv') -NoTypeInformation -Encoding utf8
    }
    catch {
        Write-ToolkitLog ('Could not write the manifest: {0}' -f $_.Exception.Message) -Level ERROR
    }

    Set-ToolkitConfigValue -Name 'LastManifestPath' -Value $manifestCsv | Out-Null

    Write-Host ''
    Write-ToolkitHeader 'Download summary'
    Write-Host ('  Downloaded      : {0}' -f $downloaded)
    Write-Host ('  Already current : {0}' -f $skipped)
    Write-Host ('  Failed          : {0}' -f $failed) -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    Write-Host ('  Folders created : {0}' -f $folders.Count)
    Write-Host ('  Data            : {0}' -f (Format-ByteSize $bytesDone))
    Write-Host ('  Snapshot folder : {0}' -f $Destination)
    Write-Host ('  Manifest        : {0}' -f $manifestCsv)

    Write-ToolkitLog ('Stage 2 complete: {0} downloaded, {1} already current, {2} failed.' -f $downloaded, $skipped, $failed) `
        -Level $(if ($failed -gt 0) { 'WARN' } else { 'SUCCESS' })

    if ($failed -gt 0) {
        Write-ToolkitLog 'Re-run stage 2 to retry the failed files - files already downloaded are skipped.' -Level WARN
    }

    return [pscustomobject]@{
        Destination  = $Destination
        ManifestPath = $manifestCsv
        FileCount    = $files.Count
        Downloaded   = $downloaded
        Skipped      = $skipped
        Failed       = $failed
    }
}

Export-ModuleMember -Function @(
    'Invoke-OneDriveDownload'
    'Invoke-DriveItemDownload'
    'ConvertTo-UtcDateTime'
    'Set-LocalFileTimestamp'
    'Test-LocalCopyCurrent'
)
