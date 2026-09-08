#Requires -Version 7.0
<#
.SYNOPSIS
    Stage 3 - find OneDrive sync-conflict duplicates, and optionally clean them up.

.DESCRIPTION
    Two jobs. First, produce the list of conflict copies so stage 4 does not treat
    duplicate noise as legitimate work that needs re-uploading. Second, let the
    admin reclaim quota by deleting the copies that are provably identical to the
    file they were copied from.

    Deletion is never automatic and never touches anything whose content differs
    from the canonical file - those go to a review report for a human decision.
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking
# Invoke-DriveItemDownload is shared with stage 2 rather than reimplemented here.
Import-Module (Join-Path $PSScriptRoot 'DownloadOneDrive.psm1') -DisableNameChecking

function Compare-DriveItemContent {
    <#
    .SYNOPSIS
        Compares two catalogue entries using the strongest signal available.
    .OUTPUTS
        PSCustomObject: Match ($true/$false/$null when unknown), ComparedBy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Left,
        [Parameter(Mandatory)]$Right
    )

    foreach ($field in @('Sha256', 'QuickXorHash')) {
        $leftValue = [string]$Left.$field
        $rightValue = [string]$Right.$field
        if (-not [string]::IsNullOrWhiteSpace($leftValue) -and -not [string]::IsNullOrWhiteSpace($rightValue)) {
            return [pscustomobject]@{
                Match      = ($leftValue -eq $rightValue)
                ComparedBy = $field
            }
        }
    }

    # No hash on either side. Equal sizes are suggestive but never proof, so an
    # equal-size pair is reported for review rather than as a safe deletion.
    if ($Left.Size -ne $Right.Size) {
        return [pscustomobject]@{ Match = $false; ComparedBy = 'Size' }
    }
    return [pscustomobject]@{ Match = $null; ComparedBy = 'Size' }
}

function Get-DriveFileCatalogue {
    <#
    .SYNOPSIS
        Builds the flat file list the scanner works from, from Graph or a stage 2 manifest.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Cloud', 'Manifest')][string]$Source,
        [string]$UserId,
        [string]$ManifestPath
    )

    $catalogue = [System.Collections.Generic.List[psobject]]::new()

    if ($Source -eq 'Manifest') {
        if (-not (Test-Path -LiteralPath $ManifestPath)) {
            throw ('Manifest not found: {0}' -f $ManifestPath)
        }
        foreach ($row in (Import-Csv -LiteralPath $ManifestPath)) {
            $catalogue.Add([pscustomobject]@{
                Id           = [string]$row.ItemId
                Name         = [string]$row.Name
                RelativePath = ConvertTo-NormalizedRelativePath -Path ([string]$row.RelativePath)
                ParentPath   = ConvertTo-NormalizedRelativePath -Path (Split-Path -Parent ([string]$row.RelativePath))
                Size         = [int64]$row.SizeBytes
                Sha256       = [string]$row.Sha256
                QuickXorHash = [string]$row.QuickXorHash
                CreatedUtc   = [string]$row.CreatedUtc
                ModifiedUtc  = [string]$row.ModifiedUtc
            })
        }
        Write-ToolkitLog ('Loaded {0} files from the manifest {1}.' -f $catalogue.Count, $ManifestPath) -Level INFO
        return $catalogue
    }

    Get-OneDriveItemInventory -UserId $UserId -FilesOnly -Activity 'Stage 3: scanning OneDrive' |
        ForEach-Object {
            $catalogue.Add([pscustomobject]@{
                Id           = $_.Id
                Name         = $_.Name
                RelativePath = $_.RelativePath
                ParentPath   = $_.ParentPath
                Size         = $_.Size
                Sha256       = $_.Sha256Hash
                QuickXorHash = $_.QuickXorHash
                CreatedUtc   = $_.CreatedUtc
                ModifiedUtc  = $_.ModifiedUtc
            })
        }

    Write-ToolkitLog ('Scanned {0} files from the cloud drive.' -f $catalogue.Count) -Level INFO
    return $catalogue
}

function Find-ConflictGroup {
    <#
    .SYNOPSIS
        Groups conflict-copy candidates around the file they were copied from.
    .DESCRIPTION
        Medium-confidence matches (the 'name-MACHINE' shape, which also matches
        ordinary hyphenated filenames) are only accepted when a file with the base
        name really exists in the same folder. High-confidence matches - '(1)',
        '- Copy', 'conflicted copy' - are grouped even when the original is gone,
        because a folder full of 'report (1..7).docx' is still a duplicate set.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Catalogue)

    $byPath = @{}
    foreach ($file in $Catalogue) {
        $byPath[$file.RelativePath.ToLowerInvariant()] = $file
    }

    $groups = @{}
    $ignored = 0
    $index = 0

    foreach ($file in $Catalogue) {
        $index++
        if ($index % 250 -eq 0) {
            Write-Progress -Activity 'Stage 3: matching conflict patterns' `
                -Status ('{0}/{1} files' -f $index, $Catalogue.Count) `
                -PercentComplete ([int](($index / [math]::Max($Catalogue.Count, 1)) * 100))
        }

        $info = Get-ConflictNameInfo -Name $file.Name
        if (-not $info.IsCandidate) { continue }

        $basePath = Join-RelativePath -Parent $file.ParentPath -Child $info.BaseName
        $key = $basePath.ToLowerInvariant()
        $original = if ($byPath.ContainsKey($key)) { $byPath[$key] } else { $null }

        if ($info.RequiresOriginal -and -not $original) {
            # Almost certainly just a hyphenated filename, e.g. annual-report.docx.
            $ignored++
            continue
        }

        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [pscustomobject]@{
                BasePath = $basePath
                Original = $original
                Copies   = [System.Collections.Generic.List[psobject]]::new()
            }
        }

        $groups[$key].Copies.Add([pscustomobject]@{
            File        = $file
            PatternType = $info.PatternType
            Confidence  = $info.Confidence
            Marker      = $info.Marker
            MachineLike = (Test-MachineLikeToken -Token $info.Marker)
        })
    }

    Write-Progress -Activity 'Stage 3: matching conflict patterns' -Completed
    Write-ToolkitLog ('{0} candidate group(s) found; {1} hyphenated name(s) ignored because no base file exists beside them.' -f $groups.Count, $ignored) -Level INFO -NoConsole

    return $groups
}

function ConvertTo-DuplicateReportRow {
    <#
    .SYNOPSIS
        Turns the conflict groups into flat report rows with a classification each.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Groups)

    $rows = [System.Collections.Generic.List[psobject]]::new()

    foreach ($key in $Groups.Keys) {
        $group = $Groups[$key]

        # The canonical copy is the original when it still exists, otherwise the
        # oldest member of the group - the one the others were spawned from.
        $canonical = $group.Original
        $members = @($group.Copies)

        if (-not $canonical) {
            $ordered = $members | Sort-Object { 
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParse([string]$_.File.CreatedUtc, [ref]$parsed)) { $parsed } else { [datetime]::MaxValue }
            }
            $canonicalMember = $ordered | Select-Object -First 1
            $canonical = $canonicalMember.File
            $members = @($ordered | Select-Object -Skip 1)

            if ($members.Count -eq 0) {
                # A conflict-named file whose original is gone and which has no
                # siblings. There is nothing to compare it against and deleting
                # it would remove the only copy, so it is flagged for a human.
                $rows.Add([pscustomobject]@{
                    Classification   = 'OrphanedCopy'
                    CanonicalPath    = ''
                    DuplicatePath    = $canonical.RelativePath
                    PatternType      = $canonicalMember.PatternType
                    Confidence       = $canonicalMember.Confidence
                    Marker           = $canonicalMember.Marker
                    MachineLikeTag   = $canonicalMember.MachineLike
                    OriginalPresent  = $false
                    CanonicalSize    = 0
                    DuplicateSize    = $canonical.Size
                    ComparedBy       = 'None'
                    RecoverableBytes = 0
                    DuplicateItemId  = $canonical.Id
                    CanonicalItemId  = ''
                    CreatedUtc       = $canonical.CreatedUtc
                    ModifiedUtc      = $canonical.ModifiedUtc
                })
                continue
            }
        }

        foreach ($member in $members) {
            $comparison = Compare-DriveItemContent -Left $canonical -Right $member.File

            $classification = if ($comparison.Match -eq $true) { 'ExactDuplicate' }
                elseif ($null -eq $comparison.Match) { 'ProbableDuplicate' }
                else { 'ContentConflict' }

            $rows.Add([pscustomobject]@{
                Classification   = $classification
                CanonicalPath    = $canonical.RelativePath
                DuplicatePath    = $member.File.RelativePath
                PatternType      = $member.PatternType
                Confidence       = $member.Confidence
                Marker           = $member.Marker
                MachineLikeTag   = $member.MachineLike
                OriginalPresent  = [bool]$group.Original
                CanonicalSize    = $canonical.Size
                DuplicateSize    = $member.File.Size
                ComparedBy       = $comparison.ComparedBy
                RecoverableBytes = if ($classification -eq 'ExactDuplicate') { $member.File.Size } else { 0 }
                DuplicateItemId  = $member.File.Id
                CanonicalItemId  = $canonical.Id
                CreatedUtc       = $member.File.CreatedUtc
                ModifiedUtc      = $member.File.ModifiedUtc
            })
        }
    }

    return $rows
}

function Remove-DuplicateDriveItem {
    <#
    .SYNOPSIS
        Deletes the confirmed exact duplicates from the user's OneDrive.
    .DESCRIPTION
        Only ever called on rows classified ExactDuplicate, and only after the
        admin has seen the list and typed the confirmation word.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)]$Rows
    )

    $encoded = [uri]::EscapeDataString($UserId)
    $deleted = 0
    $failed = 0
    $reclaimed = [int64]0
    $index = 0
    $results = [System.Collections.Generic.List[psobject]]::new()

    foreach ($row in $Rows) {
        $index++
        Write-Progress -Activity 'Stage 3: deleting duplicates' `
            -Status ('{0}/{1} - {2}' -f $index, $Rows.Count, $row.DuplicatePath) `
            -PercentComplete ([int](($index / [math]::Max($Rows.Count, 1)) * 100))

        $status = 'Deleted'
        $detail = ''
        try {
            Invoke-ToolkitGraphRequest -Method DELETE -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive/items/{1}' -f $encoded, $row.DuplicateItemId) | Out-Null
            $deleted++
            $reclaimed += [int64]$row.DuplicateSize
            Write-ToolkitLog ('Deleted duplicate {0}' -f $row.DuplicatePath) -Level INFO -NoConsole
        }
        catch {
            $status = 'Failed'
            $detail = $_.Exception.Message
            $failed++
            Write-ToolkitLog ('Could not delete {0}: {1}' -f $row.DuplicatePath, $detail) -Level ERROR -NoConsole
        }

        $results.Add([pscustomobject]@{
            DuplicatePath = $row.DuplicatePath
            CanonicalPath = $row.CanonicalPath
            SizeBytes     = $row.DuplicateSize
            Status        = $status
            Detail        = $detail
        })
    }

    Write-Progress -Activity 'Stage 3: deleting duplicates' -Completed

    return [pscustomobject]@{
        Deleted   = $deleted
        Failed    = $failed
        Reclaimed = $reclaimed
        Results   = $results
    }
}

function Invoke-DuplicateScan {
    <#
    .SYNOPSIS
        Menu option 3 - scan for conflict copies and offer to clean them up.
    #>
    [CmdletBinding()]
    param(
        [string]$UserId,
        [ValidateSet('Cloud', 'Manifest')][string]$Source
    )

    Write-ToolkitHeader 'Stage 3 - Duplicate / Conflict File Scanner'

    Write-Host ''
    Write-Host '  WHAT THIS DOES' -ForegroundColor Cyan
    Write-Host '  Looks for sync-conflict copies such as "report (1).docx" and lists them,' -ForegroundColor Gray
    Write-Host '  separating the ones that are provably identical to the file they came' -ForegroundColor Gray
    Write-Host '  from, from the ones that need a human decision.' -ForegroundColor Gray
    Write-Host '  The scan only reads. You see the full list before anything is deleted,' -ForegroundColor Gray
    Write-Host '  and deleting is a separate step you have to ask for.' -ForegroundColor Gray

    $config = Get-ToolkitConfig
    $manifestPath = [string]$config['LastManifestPath']

    if (-not $Source) {
        Write-Host ''
        Write-Host '  1. Scan the live cloud drive via Graph (always current)'
        Write-Host '  2. Scan the stage 2 manifest (faster, and compares by SHA256)'
        Write-Host ''
        $choice = Read-ToolkitValue -Prompt 'Source' -Default '1'
        $Source = if ($choice -eq '2') { 'Manifest' } else { 'Cloud' }
    }

    if ($Source -eq 'Manifest') {
        $manifestPath = Read-ToolkitValue -Prompt 'Manifest CSV path' -Default $manifestPath
        if (-not (Test-Path -LiteralPath $manifestPath)) {
            Write-ToolkitLog ('Manifest not found: {0}. Run stage 2 first.' -f $manifestPath) -Level ERROR
            return $null
        }
    }
    else {
        if (-not (Connect-ToolkitGraph)) { return $null }
        if (-not $UserId) {
            $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
        }
        Set-ToolkitConfigValue -Name 'TargetUserId' -Value $UserId | Out-Null
    }

    try {
        $catalogue = Get-DriveFileCatalogue -Source $Source -UserId $UserId -ManifestPath $manifestPath
    }
    catch {
        Write-ToolkitLog ('Scan failed: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    if ($catalogue.Count -eq 0) {
        Write-ToolkitLog 'No files found to scan.' -Level WARN
        return $null
    }

    $groups = Find-ConflictGroup -Catalogue $catalogue
    $rows = ConvertTo-DuplicateReportRow -Groups $groups

    $exact = @($rows | Where-Object { $_.Classification -eq 'ExactDuplicate' })
    $probable = @($rows | Where-Object { $_.Classification -eq 'ProbableDuplicate' })
    $conflicts = @($rows | Where-Object { $_.Classification -eq 'ContentConflict' })
    $orphans = @($rows | Where-Object { $_.Classification -eq 'OrphanedCopy' })
    $reclaimable = ($exact | Measure-Object -Property RecoverableBytes -Sum).Sum
    if (-not $reclaimable) { $reclaimable = 0 }

    $reportPath = Get-ToolkitReportPath -BaseName 'duplicate-scan' -Extension 'csv'
    $duplicatePath = Get-ToolkitReportPath -BaseName 'duplicates-safe-to-delete' -Extension 'csv'
    $reviewPath = Get-ToolkitReportPath -BaseName 'conflicts-needs-review' -Extension 'csv'

    $rows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8
    $exact | Export-Csv -LiteralPath $duplicatePath -NoTypeInformation -Encoding utf8
    @($probable + $conflicts + $orphans) | Export-Csv -LiteralPath $reviewPath -NoTypeInformation -Encoding utf8

    Set-ToolkitConfigValue -Name 'LastDuplicateReportPath' -Value $reportPath | Out-Null

    Write-Host ''
    Write-ToolkitHeader 'Duplicate scan summary'
    Write-Host ('  Files scanned          : {0}' -f $catalogue.Count)
    Write-Host ('  Conflict groups        : {0}' -f $groups.Count)
    Write-Host ('  Exact duplicates       : {0} ({1} reclaimable)' -f $exact.Count, (Format-ByteSize $reclaimable)) -ForegroundColor Green
    Write-Host ('  Same size, no hash     : {0} (needs review)' -f $probable.Count) -ForegroundColor Yellow
    Write-Host ('  Same name, differing   : {0} (needs review)' -f $conflicts.Count) -ForegroundColor Yellow
    Write-Host ('  Copies with no original: {0} (needs review)' -f $orphans.Count) -ForegroundColor Yellow
    Write-Host ''
    Write-Host ('  Full report   : {0}' -f $reportPath)
    Write-Host ('  Safe to delete: {0}' -f $duplicatePath)
    Write-Host ('  Needs review  : {0}' -f $reviewPath)

    if ($Source -eq 'Cloud' -and $probable.Count -gt 0) {
        Write-Host ''
        Write-ToolkitLog 'Some files could not be hash-compared from Graph alone. Run stage 2, then re-scan using the manifest for SHA256-backed comparison.' -Level WARN
    }

    Write-ToolkitLog ('Stage 3 complete: {0} exact duplicates, {1} probable, {2} conflicts.' -f $exact.Count, $probable.Count, $conflicts.Count) -Level SUCCESS -NoConsole

    if ($exact.Count -gt 0) {
        Write-Host ''
        Write-Host 'First 20 exact duplicates:' -ForegroundColor Cyan
        $exact | Select-Object -First 20 |
            Format-Table -AutoSize @{N = 'Duplicate'; E = { $_.DuplicatePath } },
                                   @{N = 'Keeping'; E = { $_.CanonicalPath } },
                                   @{N = 'Size'; E = { Format-ByteSize $_.DuplicateSize } } |
            Out-String | Write-Host

        Write-Host ''
        Write-Host '  What next?'
        Write-Host '    1. Download the copies to a folder first, then choose whether to delete (recommended)'
        Write-Host '    2. Delete the exact duplicates without keeping a local copy'
        Write-Host '    3. Nothing - the reports are on disk'
        Write-Host ''
        $nextAction = Read-ToolkitValue -Prompt 'Choice' -Default '1'

        if ($nextAction -eq '1' -or $nextAction -eq '2') {
            if (-not $UserId) {
                $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
            }
        }

        if ($nextAction -eq '1') {
            if (Connect-ToolkitGraph) {
                Invoke-DuplicateArchive -UserId $UserId -Rows $rows | Out-Null
            }
        }
        elseif ($nextAction -eq '2') {
            Write-ToolkitLog 'Deleting without a local archive - the OneDrive recycle bin will be the only way back.' -Level WARN
            if (Confirm-ToolkitAction -Prompt ('Delete these {0} exact duplicates from OneDrive to reclaim {1}?' -f $exact.Count, (Format-ByteSize $reclaimable))) {
                $typed = Read-ToolkitValue -Prompt "Type DELETE in capitals to confirm" -AllowEmpty
                if ($typed -ceq 'DELETE') {
                    if (Connect-ToolkitGraph) {
                        $result = Remove-DuplicateDriveItem -UserId $UserId -Rows $exact
                        $deletionPath = Get-ToolkitReportPath -BaseName 'duplicates-deleted' -Extension 'csv'
                        $result.Results | Export-Csv -LiteralPath $deletionPath -NoTypeInformation -Encoding utf8

                        Write-Host ''
                        Write-ToolkitLog ('Deleted {0} duplicate(s), reclaiming {1}. {2} failed.' -f $result.Deleted, (Format-ByteSize $result.Reclaimed), $result.Failed) `
                            -Level $(if ($result.Failed -gt 0) { 'WARN' } else { 'SUCCESS' })
                        Write-Host ('  Deletion log: {0}' -f $deletionPath)
                        Write-ToolkitLog 'Deleted items go to the user''s OneDrive recycle bin and can be restored from there.' -Level INFO
                    }
                }
                else {
                    Write-ToolkitLog 'Confirmation not matched; nothing was deleted.' -Level WARN
                }
            }
            else {
                Write-ToolkitLog 'Nothing deleted. The reports are on disk for review.' -Level INFO
            }
        }
        else {
            Write-ToolkitLog 'Nothing deleted. The reports are on disk for review.' -Level INFO
        }
    }

    return [pscustomobject]@{
        ReportPath      = $reportPath
        DuplicatePath   = $duplicatePath
        ReviewPath      = $reviewPath
        ExactDuplicates = $exact.Count
        NeedsReview     = ($probable.Count + $conflicts.Count + $orphans.Count)
        Rows            = $rows
    }
}

# ---------------------------------------------------------------------------
# Archive the copies before deleting them
# ---------------------------------------------------------------------------
#
# Deleting a duplicate straight from the drive leaves the OneDrive recycle bin as
# the only way back, on a retention clock, in the same tenant that just had a
# problem. Downloading each copy first gives an offline backup that can go on a
# USB stick, and makes the deletion reversible by hand.
#
# The rule that makes this safe: a copy is only ever eligible for deletion when it
# was downloaded AND verified. Anything that failed to transfer, or came down the
# wrong size, is never put forward for deletion.

$script:VerifiedStatus = @('Sha256Matched', 'SizeMatched')

function Select-DuplicateArchiveRow {
    <#
    .SYNOPSIS
        Filters scan rows down to the classifications the admin chose to archive.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$Rows,
        [string[]]$Classification = @('ExactDuplicate', 'ProbableDuplicate', 'ContentConflict', 'OrphanedCopy')
    )

    # The leading comma keeps an empty result an empty array: a bare "return @()"
    # is unrolled by the pipeline to nothing, and the caller's .Count then throws.
    return ,@($Rows | Where-Object { $Classification -contains [string]$_.Classification })
}

function Get-ArchiveDeletionCandidate {
    <#
    .SYNOPSIS
        Returns only the copies that are safe to delete: verified on disk AND eligible.

    .DESCRIPTION
        This is the gate the whole archive-then-delete flow rests on. A row reaches
        it only if the download completed and the local file matched what the drive
        said it should be, so a failed or truncated transfer can never result in the
        cloud copy being removed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$ArchiveResult,
        [string[]]$EligibleClassification = @('ExactDuplicate')
    )

    return ,@($ArchiveResult | Where-Object {
        $script:VerifiedStatus -contains [string]$_.VerifyStatus -and
        $EligibleClassification -contains [string]$_.Classification
    })
}

function Invoke-DuplicateArchive {
    <#
    .SYNOPSIS
        Downloads every flagged copy to a local folder, verifies it, then optionally
        deletes the verified ones from OneDrive.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Rows,
        [string]$Destination
    )

    if ($Rows.Count -eq 0) {
        Write-ToolkitLog 'No duplicate rows to archive.' -Level WARN
        return $null
    }

    $config = Get-ToolkitConfig

    Write-Host ''
    Write-ToolkitHeader 'Archive copies before deleting'
    foreach ($name in @('ExactDuplicate', 'ProbableDuplicate', 'ContentConflict', 'OrphanedCopy')) {
        $count = @($Rows | Where-Object { $_.Classification -eq $name }).Count
        Write-Host ('  {0,-20}: {1}' -f $name, $count)
    }

    Write-Host ''
    Write-Host '  Which copies should be downloaded?'
    Write-Host '    1. Everything flagged (recommended - the archive is the safety net)'
    Write-Host '    2. Only the hash-verified exact duplicates'
    Write-Host ''
    $scopeChoice = Read-ToolkitValue -Prompt 'Choice' -Default '1'
    $classifications = if ($scopeChoice -eq '2') {
        @('ExactDuplicate')
    }
    else {
        @('ExactDuplicate', 'ProbableDuplicate', 'ContentConflict', 'OrphanedCopy')
    }

    $targets = Select-DuplicateArchiveRow -Rows $Rows -Classification $classifications
    if ($targets.Count -eq 0) {
        Write-ToolkitLog 'Nothing matches that selection.' -Level WARN
        return $null
    }

    $totalBytes = [int64]0
    foreach ($row in $targets) {
        $size = [int64]0
        if ([int64]::TryParse([string]$row.DuplicateSize, [ref]$size)) { $totalBytes += $size }
    }

    if (-not $Destination) {
        $Destination = Read-ToolkitDirectory -Prompt 'Folder to archive the copies into (a USB drive is fine)' `
            -Default ([string]$config['DuplicateArchivePath']) -CreateIfMissing
    }
    if (-not $Destination) {
        Write-ToolkitLog 'No destination chosen; nothing was archived or deleted.' -Level WARN
        return $null
    }

    Set-ToolkitConfigValue -Name 'DuplicateArchivePath' -Value $Destination | Out-Null
    $archiveRoot = Join-Path $Destination ('duplicate-archive-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null

    $freeSpace = Get-FreeDiskSpace -Path $archiveRoot
    Write-Host ''
    Write-Host ('  Copies to archive : {0}' -f $targets.Count)
    Write-Host ('  Total size        : {0}' -f (Format-ByteSize $totalBytes))
    Write-Host ('  Archive folder    : {0}' -f $archiveRoot)
    if ($null -ne $freeSpace) {
        Write-Host ('  Free space        : {0}' -f (Format-ByteSize $freeSpace))
        if ($freeSpace -lt $totalBytes) {
            Write-ToolkitLog 'There is not enough free space at the destination for the whole archive.' -Level WARN
            if (-not (Confirm-ToolkitAction -Prompt 'Continue anyway?')) { return $null }
        }
    }

    Write-Host ''
    if (-not (Confirm-ToolkitAction -Prompt ('Download these {0} copies now?' -f $targets.Count) -DefaultYes)) {
        Write-ToolkitLog 'Archive cancelled; nothing was downloaded or deleted.' -Level WARN
        return $null
    }

    $results = [System.Collections.Generic.List[psobject]]::new()
    $rowByPath = @{}
    $downloaded = 0
    $failed = 0
    $bytesDone = [int64]0
    $index = 0

    foreach ($row in $targets) {
        $index++
        Write-Progress -Activity 'Archiving duplicate copies' -Id 1 `
            -Status ('{0}/{1} - {2}' -f $index, $targets.Count, $row.DuplicatePath) `
            -PercentComplete ([int](($index / [math]::Max($targets.Count, 1)) * 100))

        $downloadStatus = 'Failed'
        $verifyStatus = 'NotDownloaded'
        $localPath = ''
        $hash = $null
        $detail = ''

        try {
            $parent = ConvertTo-NormalizedRelativePath -Path (Split-Path -Parent ([string]$row.DuplicatePath))
            $item = Get-OneDriveItemById -UserId $UserId -ItemId ([string]$row.DuplicateItemId) -ParentPath $parent
            if (-not $item) {
                throw ('No longer in the drive (item {0})' -f $row.DuplicateItemId)
            }

            $localPath = Get-SafeLocalPath -Root $archiveRoot -RelativePath ([string]$row.DuplicatePath)
            Invoke-DriveItemDownload -Item $item -TargetPath $localPath -UserId $UserId
            $downloadStatus = 'Downloaded'
            $downloaded++

            # Verify the local copy before this row can ever be considered for
            # deletion. SHA256 when the drive gives us one, size otherwise.
            $localInfo = Get-Item -LiteralPath $localPath -ErrorAction Stop
            $hash = Get-FileSha256 -Path $localPath

            if (-not [string]::IsNullOrWhiteSpace($item.Sha256Hash) -and $hash) {
                $verifyStatus = if ($hash -eq $item.Sha256Hash) { 'Sha256Matched' } else { 'Failed' }
                if ($verifyStatus -eq 'Failed') { $detail = 'Downloaded file SHA256 does not match the drive' }
            }
            elseif ($localInfo.Length -eq $item.Size -and $localInfo.Length -ge 0) {
                $verifyStatus = 'SizeMatched'
                $detail = 'Verified by size; the drive exposed no SHA256 for this file'
            }
            else {
                $verifyStatus = 'Failed'
                $detail = 'Downloaded {0} bytes but the drive reports {1}' -f $localInfo.Length, $item.Size
            }

            $bytesDone += $localInfo.Length
        }
        catch {
            $failed++
            $detail = $_.Exception.Message
            Write-ToolkitLog ('Archive FAILED {0}: {1}' -f $row.DuplicatePath, $detail) -Level ERROR -NoConsole
        }

        $result = [pscustomobject]@{
            Classification  = [string]$row.Classification
            DuplicatePath   = [string]$row.DuplicatePath
            CanonicalPath   = [string]$row.CanonicalPath
            SizeBytes       = [string]$row.DuplicateSize
            DownloadStatus  = $downloadStatus
            VerifyStatus    = $verifyStatus
            LocalPath       = $localPath
            Sha256          = $hash
            DuplicateItemId = [string]$row.DuplicateItemId
            DeleteStatus    = 'NotRequested'
            Detail          = $detail
        }
        $results.Add($result)
        $rowByPath[[string]$row.DuplicatePath] = $row
    }

    Write-Progress -Activity 'Archiving duplicate copies' -Id 1 -Completed

    $verified = @($results | Where-Object { $script:VerifiedStatus -contains $_.VerifyStatus })

    Write-Host ''
    Write-ToolkitHeader 'Archive summary'
    Write-Host ('  Downloaded : {0} ({1})' -f $downloaded, (Format-ByteSize $bytesDone)) -ForegroundColor Green
    Write-Host ('  Verified   : {0}' -f $verified.Count) -ForegroundColor Green
    Write-Host ('  Failed     : {0}' -f $failed) -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    Write-Host ('  Folder     : {0}' -f $archiveRoot)

    # Now the deletion offer, restricted to what is provably on disk.
    $eligible = @('ExactDuplicate')
    $candidates = Get-ArchiveDeletionCandidate -ArchiveResult $results -EligibleClassification $eligible

    if ($candidates.Count -gt 0) {
        Write-Host ''
        Write-Host ('  {0} archived copy/copies are hash-verified exact duplicates and could now' -f $candidates.Count) -ForegroundColor Cyan
        Write-Host '  be deleted from OneDrive. Everything else stays where it is.' -ForegroundColor Cyan

        if ($failed -gt 0) {
            Write-ToolkitLog ('{0} copy/copies failed to archive and are excluded from deletion.' -f $failed) -Level WARN
        }

        if (Confirm-ToolkitAction -Prompt ('Delete those {0} verified copies from OneDrive?' -f $candidates.Count)) {
            $typed = Read-ToolkitValue -Prompt 'Type DELETE in capitals to confirm' -AllowEmpty
            if ($typed -ceq 'DELETE') {
                $toDelete = @()
                foreach ($candidate in $candidates) {
                    if ($rowByPath.ContainsKey($candidate.DuplicatePath)) { $toDelete += $rowByPath[$candidate.DuplicatePath] }
                }

                $deletion = Remove-DuplicateDriveItem -UserId $UserId -Rows $toDelete
                $deletedPaths = @{}
                foreach ($entry in $deletion.Results) {
                    $deletedPaths[[string]$entry.DuplicatePath] = [string]$entry.Status
                }
                foreach ($result in $results) {
                    if ($deletedPaths.ContainsKey($result.DuplicatePath)) {
                        $result.DeleteStatus = $deletedPaths[$result.DuplicatePath]
                    }
                }

                Write-Host ''
                Write-ToolkitLog ('Deleted {0} copy/copies, reclaiming {1}. {2} failed.' -f
                    $deletion.Deleted, (Format-ByteSize $deletion.Reclaimed), $deletion.Failed) `
                    -Level $(if ($deletion.Failed -gt 0) { 'WARN' } else { 'SUCCESS' })
                Write-ToolkitLog 'Deleted items also go to the OneDrive recycle bin, so there are now two ways back.' -Level INFO
            }
            else {
                Write-ToolkitLog 'Confirmation not matched; nothing was deleted.' -Level WARN
            }
        }
        else {
            Write-ToolkitLog 'Nothing deleted. The archive is on disk either way.' -Level INFO
        }
    }
    elseif ($results.Count -gt 0) {
        Write-Host ''
        Write-ToolkitLog 'No archived copy is both hash-verified and classified as an exact duplicate, so nothing is offered for deletion.' -Level INFO
    }

    # The manifest lives with the archive so the USB copy explains itself, and in
    # reports/ so the run is recorded alongside everything else.
    $manifestPath = Get-ToolkitReportPath -BaseName 'duplicate-archive' -Extension 'csv'
    $results | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding utf8
    $results | Export-Csv -LiteralPath (Join-Path $archiveRoot '_archive-manifest.csv') -NoTypeInformation -Encoding utf8
    Set-ToolkitConfigValue -Name 'LastDuplicateArchivePath' -Value $manifestPath | Out-Null

    Write-Host ''
    Write-Host ('  Manifest   : {0}' -f $manifestPath)
    Write-Host ('  Also saved : {0}' -f (Join-Path $archiveRoot '_archive-manifest.csv'))
    Write-Host ''
    Write-Host '  To restore later: the archive keeps each file at its original relative' -ForegroundColor DarkGray
    Write-Host '  path, so the folder can be pointed at as the local backup in stage 4,' -ForegroundColor DarkGray
    Write-Host '  or the files copied back by hand.' -ForegroundColor DarkGray

    Write-ToolkitLog ('Duplicate archive complete: {0} downloaded, {1} verified, {2} failed, manifest at {3}.' -f
        $downloaded, $verified.Count, $failed, $manifestPath) -Level SUCCESS -NoConsole

    return [pscustomobject]@{
        ArchiveRoot  = $archiveRoot
        ManifestPath = $manifestPath
        Downloaded   = $downloaded
        Verified     = $verified.Count
        Failed       = $failed
        Deleted      = @($results | Where-Object { $_.DeleteStatus -eq 'Deleted' }).Count
        Results      = $results
    }
}

function Import-DuplicateScanReport {
    <#
    .SYNOPSIS
        Reads a previous duplicate-scan CSV back into scan rows.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw ('Duplicate scan report not found: {0}' -f $Path)
    }

    $rows = @(Import-Csv -LiteralPath $Path)
    $required = @('Classification', 'DuplicatePath', 'DuplicateItemId')
    if ($rows.Count -gt 0) {
        foreach ($column in $required) {
            if (-not $rows[0].PSObject.Properties[$column]) {
                throw ('{0} is missing the {1} column, so it is not a duplicate scan report.' -f $Path, $column)
            }
        }
    }
    return $rows
}

function Invoke-DuplicateArchiveFromReport {
    <#
    .SYNOPSIS
        Menu option 8 - archive (and optionally delete) copies using a scan report.
    #>
    [CmdletBinding()]
    param(
        [string]$UserId,
        [string]$ReportPath
    )

    Write-ToolkitHeader 'Archive Duplicate Copies - Download, then Optionally Delete'

    Write-Host ''
    Write-Host '  WHAT THIS DOES' -ForegroundColor Cyan
    Write-Host '  Downloads every duplicate copy found by option 3 to a folder you choose' -ForegroundColor Gray
    Write-Host '  (a USB drive is the point), then checks each file arrived intact.' -ForegroundColor Gray
    Write-Host '  Only after that does it offer to delete them from OneDrive - and only' -ForegroundColor Gray
    Write-Host '  the ones it could verify. A copy that failed to download is never' -ForegroundColor Gray
    Write-Host '  deleted, so you always keep something to restore from.' -ForegroundColor Gray

    # Fail fast like the other stages, rather than asking for a report first and
    # only then discovering there is no way to reach the drive.
    if (-not (Test-ToolkitPrerequisite)) { return $null }

    $config = Get-ToolkitConfig

    if (-not $ReportPath) {
        Write-Host ''
        Write-Host '  1. Use the last duplicate scan report'
        Write-Host '  2. Use a different scan report CSV'
        Write-Host '  3. Run a fresh scan first (menu option 3)'
        Write-Host ''
        $choice = Read-ToolkitValue -Prompt 'Choice' -Default '1'

        if ($choice -eq '3') {
            Write-ToolkitLog 'Run menu option 3 to produce a scan report, then come back here.' -Level INFO
            return $null
        }

        $ReportPath = Read-ToolkitValue -Prompt 'Duplicate scan report (CSV)' -Default ([string]$config['LastDuplicateReportPath'])
    }

    try {
        $rows = Import-DuplicateScanReport -Path $ReportPath
    }
    catch {
        Write-ToolkitLog $_.Exception.Message -Level ERROR
        return $null
    }

    if ($rows.Count -eq 0) {
        Write-ToolkitLog 'That report contains no rows - the scan found no copies.' -Level SUCCESS
        return $null
    }

    Write-ToolkitLog ('Loaded {0} flagged copy/copies from {1}.' -f $rows.Count, $ReportPath) -Level INFO

    if (-not (Connect-ToolkitGraph)) { return $null }

    if (-not $UserId) {
        $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
    }
    Set-ToolkitConfigValue -Name 'TargetUserId' -Value $UserId | Out-Null

    return (Invoke-DuplicateArchive -UserId $UserId -Rows $rows)
}

Export-ModuleMember -Function @(
    'Invoke-DuplicateScan'
    'Get-DriveFileCatalogue'
    'Find-ConflictGroup'
    'ConvertTo-DuplicateReportRow'
    'Compare-DriveItemContent'
    'Remove-DuplicateDriveItem'
    'Invoke-DuplicateArchive'
    'Invoke-DuplicateArchiveFromReport'
    'Import-DuplicateScanReport'
    'Select-DuplicateArchiveRow'
    'Get-ArchiveDeletionCandidate'
)
