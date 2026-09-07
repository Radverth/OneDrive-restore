#Requires -Version 7.0
<#
.SYNOPSIS
    Stage 4 - compare the downloaded cloud copy against the local PC backup.

.DESCRIPTION
    Finds the legitimate work done between the restore point and now: what the
    local backup has that the post-rollback cloud copy does not, separated from
    the duplicate noise that caused the original problem.

    The cloud side is deliberately the stage 2 download rather than the problem
    device's own sync folder, so nothing that broke that device's sync gets
    dragged back into the drive.
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking

function Get-LocalFileIndex {
    <#
    .SYNOPSIS
        Indexes a folder tree by normalised, case-insensitive relative path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [string]$Activity = 'Indexing files',
        [string[]]$ExcludeName = @('_cloud-manifest.csv')
    )

    $index = [System.Collections.Generic.Dictionary[string, psobject]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $rootFull = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd([System.IO.Path]::DirectorySeparatorChar)
    $count = 0

    Get-ChildItem -LiteralPath $rootFull -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
        if ($ExcludeName -contains $_.Name) { return }
        if ($_.Extension -eq '.partial') { return }

        $relative = ConvertTo-NormalizedRelativePath -Path $_.FullName.Substring($rootFull.Length)
        $index[$relative] = [pscustomobject]@{
            RelativePath = $relative
            Name         = $_.Name
            FullPath     = $_.FullName
            Size         = $_.Length
            ModifiedUtc  = $_.LastWriteTimeUtc
            CreatedUtc   = $_.CreationTimeUtc
            Sha256       = $null
        }

        $count++
        if ($count % 200 -eq 0) {
            Write-Progress -Activity $Activity -Status ('{0} files indexed' -f $count)
        }
    }

    Write-Progress -Activity $Activity -Completed
    Write-ToolkitLog ('{0}: {1} files under {2}' -f $Activity, $index.Count, $rootFull) -Level INFO
    return $index
}

function Get-IndexedFileHash {
    <#
    .SYNOPSIS
        Hashes an indexed file once and caches the result on the entry.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Entry)

    if ($Entry.Sha256) { return $Entry.Sha256 }
    $Entry.Sha256 = Get-FileSha256 -Path $Entry.FullPath
    return $Entry.Sha256
}

function Test-LocalDuplicatePattern {
    <#
    .SYNOPSIS
        Decides whether a local file is sync-conflict noise rather than real work.
    .DESCRIPTION
        Same two-tier rule as stage 3: high-confidence patterns stand on their own;
        the 'name-MACHINE' shape only counts when a file with the base name exists
        in the same folder of the backup, or stage 3 already flagged that path as
        a duplicate in the cloud.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)]$LocalIndex,
        [Parameter(Mandatory)][AllowNull()]$KnownDuplicatePath
    )

    $info = Get-ConflictNameInfo -Name $Entry.Name

    $result = [pscustomobject]@{
        IsDuplicate = $false
        PatternType = $info.PatternType
        Confidence  = $info.Confidence
        BasePath    = ''
        Reason      = ''
        Source      = ''
    }

    if ($KnownDuplicatePath -and $KnownDuplicatePath.Contains($Entry.RelativePath)) {
        $result.IsDuplicate = $true
        $result.Source = 'ScanReport'
        $result.Reason = 'Flagged as a duplicate by the stage 3 scan'
        return $result
    }

    if (-not $info.IsCandidate) { return $result }

    $parent = ConvertTo-NormalizedRelativePath -Path (Split-Path -Parent $Entry.RelativePath)
    $basePath = Join-RelativePath -Parent $parent -Child $info.BaseName
    $result.BasePath = $basePath

    if ($info.RequiresOriginal -and -not $LocalIndex.ContainsKey($basePath)) {
        # Just a hyphenated filename; treat it as a normal file.
        return $result
    }

    $result.IsDuplicate = $true
    $result.Source = 'Pattern'
    $result.Reason = ('Matches the {0} conflict pattern (base file: {1})' -f $info.PatternType, $basePath)
    return $result
}

function Import-KnownDuplicatePath {
    <#
    .SYNOPSIS
        Loads the duplicate paths from a stage 3 report into a case-insensitive set.
    #>
    [CmdletBinding()]
    param([string]$ReportPath)

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ([string]::IsNullOrWhiteSpace($ReportPath) -or -not (Test-Path -LiteralPath $ReportPath)) {
        return $set
    }

    try {
        foreach ($row in (Import-Csv -LiteralPath $ReportPath)) {
            $path = ConvertTo-NormalizedRelativePath -Path ([string]$row.DuplicatePath)
            if ($path) { [void]$set.Add($path) }
        }
        Write-ToolkitLog ('Loaded {0} known duplicate path(s) from {1}.' -f $set.Count, $ReportPath) -Level INFO
    }
    catch {
        Write-ToolkitLog ('Could not read the stage 3 report {0}: {1}' -f $ReportPath, $_.Exception.Message) -Level WARN
    }

    return $set
}

function Invoke-DriveComparison {
    <#
    .SYNOPSIS
        Menu option 4 - categorise every file across the backup and the cloud copy.
    #>
    [CmdletBinding()]
    param(
        [string]$BackupPath,
        [string]$CloudPath,
        [datetime]$RestorePointUtc
    )

    Write-ToolkitHeader 'Stage 4 - Compare Downloaded OneDrive Copy vs. Local PC Backup'

    $config = Get-ToolkitConfig

    if (-not $PSBoundParameters.ContainsKey('RestorePointUtc')) {
        Write-Host ''
        Write-Host 'The restore point is the date/time the drive was rolled back to in the' -ForegroundColor Gray
        Write-Host 'admin centre. Local files newer than this are the work to be recovered.' -ForegroundColor Gray
        Write-Host ''

        $default = ''
        if (-not [string]::IsNullOrWhiteSpace([string]$config['RestorePointUtc'])) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse([string]$config['RestorePointUtc'], [ref]$parsed)) {
                $default = $parsed.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
            }
        }
        $RestorePointUtc = Read-ToolkitDateTime -Prompt 'Restore point date/time (local time)' -Default $default
    }

    if (-not $CloudPath) {
        $CloudPath = Read-ToolkitDirectory -Prompt 'Downloaded cloud copy folder (stage 2 output)' -Default ([string]$config['DownloadPath']) -MustExist
    }
    if (-not $CloudPath) { Write-ToolkitLog 'No cloud copy folder given; comparison cancelled.' -Level WARN; return $null }

    if (-not $BackupPath) {
        $BackupPath = Read-ToolkitDirectory -Prompt 'Local PC backup folder' -Default ([string]$config['LocalBackupPath']) -MustExist
    }
    if (-not $BackupPath) { Write-ToolkitLog 'No backup folder given; comparison cancelled.' -Level WARN; return $null }

    $config['RestorePointUtc'] = $RestorePointUtc.ToString('o')
    $config['LocalBackupPath'] = $BackupPath
    $config['DownloadPath'] = $CloudPath
    Save-ToolkitConfig -Config $config | Out-Null

    Write-Host ''
    Write-ToolkitLog ('Restore point : {0:yyyy-MM-dd HH:mm} UTC' -f $RestorePointUtc) -Level INFO
    Write-ToolkitLog ('Cloud copy    : {0}' -f $CloudPath) -Level INFO
    Write-ToolkitLog ('Local backup  : {0}' -f $BackupPath) -Level INFO
    Write-Host ''

    $cloudIndex = Get-LocalFileIndex -Root $CloudPath -Activity 'Indexing the downloaded cloud copy'
    $localIndex = Get-LocalFileIndex -Root $BackupPath -Activity 'Indexing the local PC backup'

    $knownDuplicates = Import-KnownDuplicatePath -ReportPath ([string]$config['LastDuplicateReportPath'])

    $rows = [System.Collections.Generic.List[psobject]]::new()
    $index = 0
    $total = $localIndex.Count

    foreach ($relative in $localIndex.Keys) {
        $index++
        $local = $localIndex[$relative]

        if ($index % 25 -eq 0 -or $index -eq $total) {
            Write-Progress -Activity 'Stage 4: comparing' `
                -Status ('{0}/{1} - {2}' -f $index, $total, $local.RelativePath) `
                -PercentComplete ([int](($index / [math]::Max($total, 1)) * 100))
        }

        $duplicate = Test-LocalDuplicatePattern -Entry $local -LocalIndex $localIndex -KnownDuplicatePath $knownDuplicates

        $cloud = $null
        if ($cloudIndex.ContainsKey($relative)) { $cloud = $cloudIndex[$relative] }

        $category = ''
        $action = ''
        $reason = ''
        $localHash = $null
        $cloudHash = $null

        if ($duplicate.IsDuplicate) {
            # A conflict-named file is only safe to discard when the file it was
            # copied from actually exists somewhere. 'report (1).docx' with no
            # 'report.docx' on either side may be the only copy of real work, so
            # it goes to a human instead of being dropped on the floor.
            $baseKnown = $duplicate.Source -eq 'ScanReport' -or
                         ($duplicate.BasePath -and ($localIndex.ContainsKey($duplicate.BasePath) -or $cloudIndex.ContainsKey($duplicate.BasePath)))

            if ($baseKnown) {
                $category = 'DuplicatePatternSkip'
                $action = 'Skip'
                $reason = $duplicate.Reason
            }
            else {
                $category = 'DuplicatePatternNoOriginal'
                $action = 'Review'
                $reason = ('Matches the {0} conflict pattern, but no file called {1} exists in the backup or the cloud copy - it may be the only copy' -f $duplicate.PatternType, $duplicate.BasePath)
            }
        }
        elseif ($cloud) {
            if ($local.Size -ne $cloud.Size) {
                # Different sizes cannot be the same content; skip the hashing cost.
                $category = 'ContentMismatch'
                $action = 'Review'
                $reason = 'Present in both, sizes differ'
            }
            else {
                $localHash = Get-IndexedFileHash -Entry $local
                $cloudHash = Get-IndexedFileHash -Entry $cloud

                if ($localHash -and $cloudHash -and $localHash -eq $cloudHash) {
                    $category = 'Identical'
                    $action = 'None'
                    $reason = 'Same SHA256 on both sides'
                }
                elseif (-not $localHash -or -not $cloudHash) {
                    $category = 'ContentMismatch'
                    $action = 'Review'
                    $reason = 'One side could not be hashed (file locked or unreadable)'
                }
                else {
                    $category = 'ContentMismatch'
                    $action = 'Review'
                    $reason = if ($local.ModifiedUtc -gt $RestorePointUtc) {
                        'Present in both with different content; the local copy was edited after the restore point'
                    }
                    else {
                        'Present in both with different content'
                    }
                }
            }
        }
        elseif ($local.ModifiedUtc -gt $RestorePointUtc) {
            $category = 'UploadCandidate'
            $action = 'Upload'
            $reason = 'Newer than the restore point and missing from the cloud copy'
            $localHash = Get-IndexedFileHash -Entry $local
        }
        else {
            $category = 'LocalOnlyBeforeRestorePoint'
            $action = 'Review'
            $reason = 'Missing from the cloud copy but not modified since the restore point'
        }

        $rows.Add([pscustomobject]@{
            Category         = $category
            Action           = $action
            RelativePath     = $local.RelativePath
            Reason           = $reason
            LocalPath        = $local.FullPath
            LocalSizeBytes   = $local.Size
            LocalModifiedUtc = $local.ModifiedUtc.ToString('o')
            LocalSha256      = $localHash
            CloudPath        = if ($cloud) { $cloud.FullPath } else { '' }
            CloudSizeBytes   = if ($cloud) { $cloud.Size } else { '' }
            CloudModifiedUtc = if ($cloud) { $cloud.ModifiedUtc.ToString('o') } else { '' }
            CloudSha256      = $cloudHash
            PatternType      = $duplicate.PatternType
            PatternBasePath  = $duplicate.BasePath
        })
    }

    Write-Progress -Activity 'Stage 4: comparing' -Completed

    # Cloud-only files are informational: the rollback restored them and the
    # backup simply predates or postdates them. Nothing to do, but worth seeing.
    foreach ($relative in $cloudIndex.Keys) {
        if ($localIndex.ContainsKey($relative)) { continue }
        $cloud = $cloudIndex[$relative]
        $rows.Add([pscustomobject]@{
            Category         = 'CloudOnly'
            Action           = 'None'
            RelativePath     = $cloud.RelativePath
            Reason           = 'In the cloud copy but not in the local backup'
            LocalPath        = ''
            LocalSizeBytes   = ''
            LocalModifiedUtc = ''
            LocalSha256      = $null
            CloudPath        = $cloud.FullPath
            CloudSizeBytes   = $cloud.Size
            CloudModifiedUtc = $cloud.ModifiedUtc.ToString('o')
            CloudSha256      = $null
            PatternType      = 'None'
            PatternBasePath  = ''
        })
    }

    $reportPath = Get-ToolkitReportPath -BaseName 'comparison' -Extension 'csv'
    $rows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8
    Set-ToolkitConfigValue -Name 'LastComparisonReportPath' -Value $reportPath | Out-Null

    $uploadRows = @($rows | Where-Object { $_.Category -eq 'UploadCandidate' })
    $uploadBytes = ($uploadRows | Measure-Object -Property LocalSizeBytes -Sum).Sum
    if (-not $uploadBytes) { $uploadBytes = 0 }

    Write-Host ''
    Write-ToolkitHeader 'Comparison summary'
    foreach ($category in @('UploadCandidate', 'ContentMismatch', 'DuplicatePatternNoOriginal', 'DuplicatePatternSkip', 'Identical', 'LocalOnlyBeforeRestorePoint', 'CloudOnly')) {
        $count = @($rows | Where-Object { $_.Category -eq $category }).Count
        $colour = switch ($category) {
            'UploadCandidate' { 'Green' }
            'ContentMismatch' { 'Yellow' }
            'DuplicatePatternNoOriginal' { 'Yellow' }
            'LocalOnlyBeforeRestorePoint' { 'Yellow' }
            default { 'Gray' }
        }
        Write-Host ('  {0,-28}: {1}' -f $category, $count) -ForegroundColor $colour
    }

    Write-Host ''
    Write-Host ('  To upload in stage 5: {0} file(s), {1}' -f $uploadRows.Count, (Format-ByteSize $uploadBytes)) -ForegroundColor Cyan
    Write-Host ('  Report              : {0}' -f $reportPath)

    Write-ToolkitLog ('Stage 4 complete: {0} upload candidates, {1} need review, report at {2}.' -f
        $uploadRows.Count,
        @($rows | Where-Object { $_.Action -eq 'Review' }).Count,
        $reportPath) -Level SUCCESS -NoConsole

    return [pscustomobject]@{
        ReportPath      = $reportPath
        RestorePointUtc = $RestorePointUtc
        Rows            = $rows
        UploadCount     = $uploadRows.Count
        ReviewCount     = @($rows | Where-Object { $_.Action -eq 'Review' }).Count
    }
}

Export-ModuleMember -Function @(
    'Invoke-DriveComparison'
    'Get-LocalFileIndex'
    'Get-IndexedFileHash'
    'Test-LocalDuplicatePattern'
    'Import-KnownDuplicatePath'
)
