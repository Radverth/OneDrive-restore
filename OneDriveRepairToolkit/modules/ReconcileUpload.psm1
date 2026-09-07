#Requires -Version 7.0
<#
.SYNOPSIS
    Stage 5 - upload the canonical version of each recovered file back to OneDrive.

.DESCRIPTION
    Uploads only what stage 4 classified as legitimate recovered work, plus any
    same-name-different-content conflicts the admin explicitly resolves in favour
    of the local copy. Nothing else is touched, so the duplicate problem the
    rollback was meant to fix does not come straight back.

    Every upload uses conflictBehavior=replace rather than rename: a rename would
    create exactly the 'file (1).docx' copies this whole exercise is undoing.
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking

# Graph's cutoff for a simple content PUT; anything larger needs an upload session.
$script:SimpleUploadLimit = 4MB

# Must be a multiple of 320 KiB (327680 bytes). 10 MB = 32 x 320 KiB.
$script:ChunkSize = 10485760

function ConvertTo-DrivePathUrl {
    <#
    .SYNOPSIS
        Escapes each segment of a drive-relative path for path-addressed Graph URLs.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RelativePath)

    $normalized = ConvertTo-NormalizedRelativePath -Path $RelativePath
    return (($normalized -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
}

function New-DriveFolderPath {
    <#
    .SYNOPSIS
        Makes sure every folder in a relative path exists on the drive.
    .DESCRIPTION
        Path-addressed uploads fail when the parent folder is missing, so the
        chain is created top-down. Already-created paths are cached for the run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$FolderPath,
        [Parameter(Mandatory)]$Cache
    )

    $normalized = ConvertTo-NormalizedRelativePath -Path $FolderPath
    if ([string]::IsNullOrWhiteSpace($normalized)) { return }
    if ($Cache.Contains($normalized)) { return }

    $encodedUser = [uri]::EscapeDataString($UserId)
    $segments = $normalized -split '/'
    $currentPath = ''

    foreach ($segment in $segments) {
        $parentPath = $currentPath
        $currentPath = if ($currentPath) { '{0}/{1}' -f $currentPath, $segment } else { $segment }

        if ($Cache.Contains($currentPath)) { continue }

        $probeUri = 'https://graph.microsoft.com/v1.0/users/{0}/drive/root:/{1}' -f $encodedUser, (ConvertTo-DrivePathUrl -RelativePath $currentPath)
        $existing = Invoke-ToolkitGraphRequest -Uri $probeUri -AllowNotFound

        if ($existing) {
            [void]$Cache.Add($currentPath)
            continue
        }

        $parentUri = if ($parentPath) {
            'https://graph.microsoft.com/v1.0/users/{0}/drive/root:/{1}:/children' -f $encodedUser, (ConvertTo-DrivePathUrl -RelativePath $parentPath)
        }
        else {
            'https://graph.microsoft.com/v1.0/users/{0}/drive/root/children' -f $encodedUser
        }

        $body = @{
            name                                = $segment
            folder                              = @{}
            '@microsoft.graph.conflictBehavior' = 'replace'
        }

        Invoke-ToolkitGraphRequest -Method POST -Uri $parentUri -Body ($body | ConvertTo-Json -Depth 4) | Out-Null
        [void]$Cache.Add($currentPath)
        Write-ToolkitLog ('Created folder {0}' -f $currentPath) -Level INFO -NoConsole
    }
}

function Invoke-SimpleFileUpload {
    <#
    .SYNOPSIS
        Uploads a file under 4 MB with a single content PUT.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$LocalPath,
        [Parameter(Mandatory)][string]$RelativePath,
        [int]$MaxRetry = 4
    )

    $uri = 'https://graph.microsoft.com/v1.0/users/{0}/drive/root:/{1}:/content?@microsoft.graph.conflictBehavior=replace' -f
        [uri]::EscapeDataString($UserId), (ConvertTo-DrivePathUrl -RelativePath $RelativePath)

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return (Invoke-MgGraphRequest -Method PUT -Uri $uri -InputFilePath $LocalPath -ContentType 'application/octet-stream' -OutputType Hashtable -ErrorAction Stop)
        }
        catch {
            $detail = Get-GraphErrorDetail -ErrorRecord $_
            $retryable = $detail.StatusCode -in @(429, 500, 502, 503, 504) -or $detail.StatusCode -eq 0
            if (-not $retryable -or $attempt -gt $MaxRetry) { throw }

            $delay = if ($detail.RetryAfter -and $detail.RetryAfter -gt 0) { [math]::Min($detail.RetryAfter, 300) }
                     else { [math]::Min([math]::Pow(2, $attempt), 60) }
            Write-ToolkitLog ('Upload of {0} returned HTTP {1}; retry {2}/{3} in {4}s.' -f $RelativePath, $detail.StatusCode, $attempt, $MaxRetry, $delay) -Level WARN -NoConsole
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-ChunkedFileUpload {
    <#
    .SYNOPSIS
        Uploads a file of 4 MB or more through a resumable upload session.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$LocalPath,
        [Parameter(Mandatory)][string]$RelativePath,
        [string]$ProgressActivity = 'Stage 5: uploading'
    )

    $sessionUri = 'https://graph.microsoft.com/v1.0/users/{0}/drive/root:/{1}:/createUploadSession' -f
        [uri]::EscapeDataString($UserId), (ConvertTo-DrivePathUrl -RelativePath $RelativePath)

    $sessionBody = @{
        item = @{
            '@microsoft.graph.conflictBehavior' = 'replace'
            name                                = [System.IO.Path]::GetFileName($RelativePath)
        }
    }

    $session = Invoke-ToolkitGraphRequest -Method POST -Uri $sessionUri -Body ($sessionBody | ConvertTo-Json -Depth 5)
    $uploadUrl = [string](Get-GraphValue -Item $session -Key 'uploadUrl')
    if ([string]::IsNullOrWhiteSpace($uploadUrl)) {
        throw ('Graph did not return an upload URL for {0}.' -f $RelativePath)
    }

    $fileInfo = Get-Item -LiteralPath $LocalPath
    $totalLength = $fileInfo.Length
    $stream = [System.IO.File]::OpenRead($LocalPath)
    $response = $null

    try {
        $buffer = [byte[]]::new($script:ChunkSize)
        $position = [int64]0

        while ($position -lt $totalLength) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }

            # A range slice would produce Object[] (and allocate ten million
            # entries for a full chunk), so copy into a real byte[] instead.
            $chunk = $buffer
            if ($read -ne $buffer.Length) {
                $chunk = [byte[]]::new($read)
                [System.Array]::Copy($buffer, 0, $chunk, 0, $read)
            }
            $rangeEnd = $position + $read - 1

            $headers = @{
                'Content-Range'  = ('bytes {0}-{1}/{2}' -f $position, $rangeEnd, $totalLength)
                'Content-Length' = $read
            }

            Write-Progress -Activity $ProgressActivity -Id 2 `
                -Status ('{0} - {1} of {2}' -f (Split-Path -Leaf $RelativePath), (Format-ByteSize ($position + $read)), (Format-ByteSize $totalLength)) `
                -PercentComplete ([int]((($position + $read) / [math]::Max($totalLength, 1)) * 100))

            $response = Invoke-ToolkitWebRequest -Uri $uploadUrl -Method PUT -Body $chunk -Headers $headers
            $position += $read
        }
    }
    catch {
        # Cancel the half-finished session so it does not linger on the drive.
        try { Invoke-ToolkitWebRequest -Uri $uploadUrl -Method DELETE -MaxRetry 1 | Out-Null }
        catch { Write-Verbose 'Upload session could not be cancelled; it will expire on its own.' }
        throw
    }
    finally {
        $stream.Dispose()
        Write-Progress -Activity $ProgressActivity -Id 2 -Completed
    }

    if ($response -and $response.Content) {
        try { return ($response.Content | ConvertFrom-Json -AsHashtable) }
        catch { Write-Verbose 'Final upload response was not JSON.' }
    }
    return $null
}

function Set-DriveItemTimestamp {
    <#
    .SYNOPSIS
        Restores the original modified time on an uploaded item.
    .DESCRIPTION
        createdDateTime is sent too, best-effort: OneDrive frequently keeps the
        upload time for newly created items regardless, which is a platform
        behaviour and not something the client can force.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$ItemId,
        [Parameter(Mandatory)][datetime]$ModifiedUtc,
        [AllowNull()][Nullable[datetime]]$CreatedUtc
    )

    $fileSystemInfo = @{ lastModifiedDateTime = $ModifiedUtc.ToUniversalTime().ToString('o') }
    if ($CreatedUtc) {
        $fileSystemInfo['createdDateTime'] = $CreatedUtc.Value.ToUniversalTime().ToString('o')
    }

    $body = @{ fileSystemInfo = $fileSystemInfo } | ConvertTo-Json -Depth 4
    $uri = 'https://graph.microsoft.com/v1.0/users/{0}/drive/items/{1}' -f [uri]::EscapeDataString($UserId), $ItemId

    Invoke-ToolkitGraphRequest -Method PATCH -Uri $uri -Body $body | Out-Null
}

function Invoke-CanonicalFileUpload {
    <#
    .SYNOPSIS
        Uploads one file and restores its timestamps, picking the right upload mode.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$LocalPath,
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)]$FolderCache
    )

    if (-not (Test-Path -LiteralPath $LocalPath -PathType Leaf)) {
        throw ('Local file no longer exists: {0}' -f $LocalPath)
    }

    $fileInfo = Get-Item -LiteralPath $LocalPath
    $parent = ConvertTo-NormalizedRelativePath -Path (Split-Path -Parent $RelativePath)
    New-DriveFolderPath -UserId $UserId -FolderPath $parent -Cache $FolderCache

    $uploaded = if ($fileInfo.Length -lt $script:SimpleUploadLimit) {
        Invoke-SimpleFileUpload -UserId $UserId -LocalPath $LocalPath -RelativePath $RelativePath
    }
    else {
        Invoke-ChunkedFileUpload -UserId $UserId -LocalPath $LocalPath -RelativePath $RelativePath
    }

    $itemId = [string](Get-GraphValue -Item $uploaded -Key 'id')
    $timestampSet = $false

    if ($itemId) {
        try {
            Set-DriveItemTimestamp -UserId $UserId -ItemId $itemId -ModifiedUtc $fileInfo.LastWriteTimeUtc -CreatedUtc $fileInfo.CreationTimeUtc
            $timestampSet = $true
        }
        catch {
            Write-ToolkitLog ('Uploaded {0} but could not restore its timestamps: {1}' -f $RelativePath, $_.Exception.Message) -Level WARN -NoConsole
        }
    }

    return [pscustomobject]@{
        ItemId       = $itemId
        SizeBytes    = $fileInfo.Length
        Mode         = if ($fileInfo.Length -lt $script:SimpleUploadLimit) { 'Simple' } else { 'Chunked' }
        TimestampSet = $timestampSet
    }
}

function Resolve-ContentConflict {
    <#
    .SYNOPSIS
        Walks the admin through each same-name-different-content conflict.
    .DESCRIPTION
        Never guessed at: choosing wrong silently loses data on one side or the
        other, so each one is an explicit decision (with an apply-to-all escape
        hatch for admins who have already made up their mind).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Rows)

    $decisions = [System.Collections.Generic.List[psobject]]::new()
    if ($Rows.Count -eq 0) { return $decisions }

    Write-Host ''
    Write-ToolkitHeader ('Manual review - {0} same-name conflict(s)' -f $Rows.Count)
    Write-Host '  L = keep the local backup copy (upload it, replacing the cloud file)'
    Write-Host '  C = keep the cloud copy (do nothing)'
    Write-Host '  S = skip and decide later'
    Write-Host '  A = apply the next choice to all remaining conflicts'
    Write-Host '  Q = stop reviewing; everything left is skipped'
    Write-Host ''

    $applyToAll = $null
    $index = 0

    foreach ($row in $Rows) {
        $index++
        $choice = $applyToAll

        if (-not $choice) {
            Write-Host ''
            Write-Host ('[{0}/{1}] {2}' -f $index, $Rows.Count, $row.RelativePath) -ForegroundColor Cyan
            Write-Host ('    local : {0,-12} modified {1}' -f (Format-ByteSize ([int64]$row.LocalSizeBytes)), $row.LocalModifiedUtc)
            $cloudSize = if ([string]::IsNullOrWhiteSpace([string]$row.CloudSizeBytes)) { 0 } else { [int64]$row.CloudSizeBytes }
            Write-Host ('    cloud : {0,-12} modified {1}' -f (Format-ByteSize $cloudSize), $row.CloudModifiedUtc)
            Write-Host ('    why   : {0}' -f $row.Reason) -ForegroundColor DarkGray

            $answer = (Read-Host '    Keep [L]ocal / [C]loud / [S]kip / [A]ll / [Q]uit').Trim().ToUpperInvariant()

            if ($answer -eq 'Q') {
                Write-ToolkitLog 'Review stopped; the remaining conflicts are left untouched.' -Level WARN
                break
            }
            if ($answer -eq 'A') {
                $answer = (Read-Host '    Apply which choice to all remaining? [L]ocal / [C]loud / [S]kip').Trim().ToUpperInvariant()
                if ($answer -in @('L', 'C', 'S')) { $applyToAll = $answer } else { $answer = 'S' }
                $choice = if ($applyToAll) { $applyToAll } else { 'S' }
            }
            elseif ($answer -in @('L', 'C', 'S')) {
                $choice = $answer
            }
            else {
                Write-Host '    Not a valid choice; skipping this one.' -ForegroundColor Yellow
                $choice = 'S'
            }
        }

        $decisions.Add([pscustomobject]@{
            Row      = $row
            Decision = switch ($choice) { 'L' { 'UploadLocal' } 'C' { 'KeepCloud' } default { 'Skipped' } }
        })
    }

    return $decisions
}

function Invoke-Reconciliation {
    <#
    .SYNOPSIS
        Menu option 5 - upload the recovered files and record what changed.
    #>
    [CmdletBinding()]
    param(
        [string]$UserId,
        [string]$ComparisonReportPath,
        [switch]$WhatIfOnly
    )

    Write-ToolkitHeader 'Stage 5 - Reconcile: Upload Canonical Files Back to OneDrive'

    $config = Get-ToolkitConfig

    if (-not $ComparisonReportPath) {
        $ComparisonReportPath = Read-ToolkitValue -Prompt 'Stage 4 comparison report (CSV)' -Default ([string]$config['LastComparisonReportPath'])
    }
    if (-not (Test-Path -LiteralPath $ComparisonReportPath)) {
        Write-ToolkitLog ('Comparison report not found: {0}. Run stage 4 first.' -f $ComparisonReportPath) -Level ERROR
        return $null
    }

    $rows = @(Import-Csv -LiteralPath $ComparisonReportPath)
    $uploadRows = @($rows | Where-Object { $_.Category -eq 'UploadCandidate' })
    $conflictRows = @($rows | Where-Object { $_.Category -eq 'ContentMismatch' })
    $uploadBytes = [int64]0
    foreach ($row in $uploadRows) {
        # CSV values arrive as strings; convert explicitly rather than letting
        # Measure-Object silently skip anything it cannot coerce.
        $size = [int64]0
        if ([int64]::TryParse([string]$row.LocalSizeBytes, [ref]$size)) { $uploadBytes += $size }
    }

    Write-Host ''
    Write-Host ('  Report            : {0}' -f $ComparisonReportPath)
    Write-Host ('  Ready to upload   : {0} file(s), {1}' -f $uploadRows.Count, (Format-ByteSize $uploadBytes)) -ForegroundColor Green
    Write-Host ('  Needing a decision: {0} conflict(s)' -f $conflictRows.Count) -ForegroundColor Yellow
    Write-Host ('  Skipped as noise  : {0} duplicate-pattern file(s)' -f @($rows | Where-Object { $_.Category -eq 'DuplicatePatternSkip' }).Count)
    Write-Host ''

    if ($uploadRows.Count -eq 0 -and $conflictRows.Count -eq 0) {
        Write-ToolkitLog 'Nothing to reconcile - the cloud copy already matches the backup.' -Level SUCCESS
        return $null
    }

    if (-not $UserId) {
        $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
    }

    if (-not $WhatIfOnly) {
        $WhatIfOnly = -not (Confirm-ToolkitAction -Prompt 'Upload for real? Answer no for a dry run that changes nothing.')
    }

    if ($WhatIfOnly) {
        Write-ToolkitLog 'Dry run: no files will be uploaded.' -Level WARN
    }
    elseif (-not (Connect-ToolkitGraph)) {
        return $null
    }

    if (-not $WhatIfOnly) {
        Set-ToolkitConfigValue -Name 'TargetUserId' -Value $UserId | Out-Null
    }

    $work = [System.Collections.Generic.List[psobject]]::new()
    foreach ($row in $uploadRows) {
        $work.Add([pscustomobject]@{ Row = $row; Decision = 'UploadLocal'; Source = 'UploadCandidate' })
    }

    if ($conflictRows.Count -gt 0 -and (Confirm-ToolkitAction -Prompt ('Review the {0} same-name conflict(s) now?' -f $conflictRows.Count) -DefaultYes)) {
        foreach ($decision in (Resolve-ContentConflict -Rows $conflictRows)) {
            $work.Add([pscustomobject]@{ Row = $decision.Row; Decision = $decision.Decision; Source = 'ContentMismatch' })
        }
    }

    $toUpload = @($work | Where-Object { $_.Decision -eq 'UploadLocal' })

    Write-Host ''
    Write-ToolkitLog ('{0} file(s) will be uploaded.' -f $toUpload.Count) -Level INFO

    if (-not $WhatIfOnly -and $toUpload.Count -gt 0) {
        if (-not (Confirm-ToolkitAction -Prompt ('Proceed with uploading {0} file(s) to {1}?' -f $toUpload.Count, $UserId))) {
            Write-ToolkitLog 'Reconciliation cancelled before any upload.' -Level WARN
            return $null
        }
    }

    $folderCache = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $results = [System.Collections.Generic.List[psobject]]::new()
    $uploaded = 0
    $failed = 0
    $skipped = 0
    $bytesDone = [int64]0
    $index = 0

    foreach ($item in $work) {
        $row = $item.Row
        $index++

        Write-Progress -Activity 'Stage 5: reconciling' -Id 1 `
            -Status ('{0}/{1} - {2}' -f $index, $work.Count, $row.RelativePath) `
            -PercentComplete ([int](($index / [math]::Max($work.Count, 1)) * 100))

        if ($item.Decision -ne 'UploadLocal') {
            $skipped++
            $results.Add([pscustomobject]@{
                RelativePath = $row.RelativePath
                Category     = $row.Category
                Decision     = $item.Decision
                Status       = if ($item.Decision -eq 'KeepCloud') { 'KeptCloudCopy' } else { 'Skipped' }
                SizeBytes    = $row.LocalSizeBytes
                UploadMode   = ''
                ItemId       = ''
                Detail       = $row.Reason
            })
            continue
        }

        if ($WhatIfOnly) {
            $results.Add([pscustomobject]@{
                RelativePath = $row.RelativePath
                Category     = $row.Category
                Decision     = $item.Decision
                Status       = 'WouldUpload'
                SizeBytes    = $row.LocalSizeBytes
                UploadMode   = if ([int64]$row.LocalSizeBytes -lt $script:SimpleUploadLimit) { 'Simple' } else { 'Chunked' }
                ItemId       = ''
                Detail       = 'Dry run'
            })
            continue
        }

        try {
            $result = Invoke-CanonicalFileUpload -UserId $UserId -LocalPath $row.LocalPath -RelativePath $row.RelativePath -FolderCache $folderCache
            $uploaded++
            $bytesDone += $result.SizeBytes

            $results.Add([pscustomobject]@{
                RelativePath = $row.RelativePath
                Category     = $row.Category
                Decision     = $item.Decision
                Status       = 'Uploaded'
                SizeBytes    = $result.SizeBytes
                UploadMode   = $result.Mode
                ItemId       = $result.ItemId
                Detail       = if ($result.TimestampSet) { 'Timestamps restored' } else { 'Uploaded; timestamps not restored' }
            })

            Write-ToolkitLog ('Uploaded {0} ({1}, {2})' -f $row.RelativePath, (Format-ByteSize $result.SizeBytes), $result.Mode) -Level INFO -NoConsole
        }
        catch {
            $failed++
            $results.Add([pscustomobject]@{
                RelativePath = $row.RelativePath
                Category     = $row.Category
                Decision     = $item.Decision
                Status       = 'Failed'
                SizeBytes    = $row.LocalSizeBytes
                UploadMode   = ''
                ItemId       = ''
                Detail       = $_.Exception.Message
            })
            Write-ToolkitLog ('FAILED {0}: {1}' -f $row.RelativePath, $_.Exception.Message) -Level ERROR -NoConsole
        }
    }

    Write-Progress -Activity 'Stage 5: reconciling' -Id 1 -Completed

    $summaryPath = Get-ToolkitReportPath -BaseName $(if ($WhatIfOnly) { 'reconcile-dryrun' } else { 'reconcile-summary' }) -Extension 'csv'
    $results | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
    Set-ToolkitConfigValue -Name 'LastReconcileReportPath' -Value $summaryPath | Out-Null

    Write-Host ''
    Write-ToolkitHeader $(if ($WhatIfOnly) { 'Dry run summary' } else { 'Reconciliation summary' })
    if ($WhatIfOnly) {
        Write-Host ('  Would upload : {0}' -f @($results | Where-Object { $_.Status -eq 'WouldUpload' }).Count) -ForegroundColor Green
    }
    else {
        Write-Host ('  Uploaded     : {0} ({1})' -f $uploaded, (Format-ByteSize $bytesDone)) -ForegroundColor Green
        Write-Host ('  Failed       : {0}' -f $failed) -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    }
    Write-Host ('  Skipped      : {0}' -f $skipped)
    Write-Host ('  Summary CSV  : {0}' -f $summaryPath)

    Write-ToolkitLog ('Stage 5 complete: {0} uploaded, {1} skipped, {2} failed.' -f $uploaded, $skipped, $failed) `
        -Level $(if ($failed -gt 0) { 'WARN' } else { 'SUCCESS' })

    if (-not $WhatIfOnly -and $uploaded -gt 0) {
        Write-ToolkitLog 'Re-run stage 2 for a fresh snapshot if you want to verify the drive now matches the backup.' -Level INFO
    }

    return [pscustomobject]@{
        SummaryPath = $summaryPath
        Uploaded    = $uploaded
        Skipped     = $skipped
        Failed      = $failed
        DryRun      = [bool]$WhatIfOnly
    }
}

Export-ModuleMember -Function @(
    'Invoke-Reconciliation'
    'Invoke-CanonicalFileUpload'
    'Invoke-SimpleFileUpload'
    'Invoke-ChunkedFileUpload'
    'New-DriveFolderPath'
    'ConvertTo-DrivePathUrl'
    'Set-DriveItemTimestamp'
    'Resolve-ContentConflict'
)
