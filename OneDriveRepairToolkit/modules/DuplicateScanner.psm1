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

        if ($Source -eq 'Cloud' -or (Confirm-ToolkitAction -Prompt 'Delete duplicates from the live drive? (needs a Graph connection)')) {
            if (Confirm-ToolkitAction -Prompt ('Delete these {0} exact duplicates from OneDrive to reclaim {1}?' -f $exact.Count, (Format-ByteSize $reclaimable))) {
                $typed = Read-ToolkitValue -Prompt "Type DELETE in capitals to confirm" -AllowEmpty
                if ($typed -ceq 'DELETE') {
                    if (-not $UserId) {
                        $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
                    }
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

Export-ModuleMember -Function @(
    'Invoke-DuplicateScan'
    'Get-DriveFileCatalogue'
    'Find-ConflictGroup'
    'ConvertTo-DuplicateReportRow'
    'Compare-DriveItemContent'
    'Remove-DuplicateDriveItem'
)
