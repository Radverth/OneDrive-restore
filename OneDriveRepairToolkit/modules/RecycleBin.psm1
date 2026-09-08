#Requires -Version 7.0
<#
.SYNOPSIS
    Inventory the user's OneDrive recycle bin and work out what is worth recovering.

.DESCRIPTION
    After a rollback and a duplicate cleanup, the recycle bin holds a mixture of
    two very different things: the conflict copies that were deleted on purpose,
    and files that were deleted before the restore point and may still be wanted.
    This stage lists the bin and separates the two, cross-referencing the stage 2
    manifest so "is this already back in the drive?" is answered from data rather
    than guessed.

    A note on what Graph can and cannot do here, because it shapes this module:

      * Listing the bin is only available on the BETA endpoint
        GET /beta/sites/{siteId}/recycleBin/items.
      * A recycleBinItem carries id, name, size, deletedDateTime and
        deletedFromLocation - and nothing else. There is no content stream and no
        @microsoft.graph.downloadUrl, so recycle bin file content cannot be
        downloaded through Graph at any version.
      * The only way to reach the bytes is to restore the item first, and
        driveItem: restore is documented as OneDrive Personal only. For OneDrive
        for Business, restore means the SharePoint REST endpoint
        POST {siteUrl}/_api/web/recyclebin('{id}')/restore(), which needs a
        SharePoint-audience token and a SharePoint application permission this
        toolkit does not currently request.

    So this stage reports; it does not restore. The report tells the admin exactly
    which items are worth restoring by hand in the web UI.
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking

function Resolve-OneDriveSiteId {
    <#
    .SYNOPSIS
        Finds the SharePoint site ID behind a user's OneDrive.
    .DESCRIPTION
        The recycle bin is addressed per site, not per drive. The backing list's
        parentReference carries the site ID; if that is missing, the drive's
        webUrl is resolved through the sites endpoint instead.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)

    $encoded = [uri]::EscapeDataString($UserId)

    $list = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive/list?$select=id,parentReference,webUrl' -f $encoded) -AllowNotFound
    if ($list) {
        $parent = Get-GraphValue -Item $list -Key 'parentReference'
        $siteId = [string](Get-GraphValue -Item $parent -Key 'siteId')
        if (-not [string]::IsNullOrWhiteSpace($siteId)) { return $siteId }
    }

    # Fallback: turn the drive's webUrl into a sites/{hostname}:/{path} lookup.
    $drive = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive?$select=id,webUrl' -f $encoded)
    $webUrl = [string](Get-GraphValue -Item $drive -Key 'webUrl')
    if ([string]::IsNullOrWhiteSpace($webUrl)) {
        throw 'Could not determine the SharePoint site behind this OneDrive.'
    }

    $uri = [uri]$webUrl
    # webUrl looks like https://tenant-my.sharepoint.com/personal/user_domain_com/Documents
    $segments = @($uri.AbsolutePath.Trim('/') -split '/')
    if ($segments.Count -lt 2) {
        throw ('Unexpected OneDrive web URL, cannot derive the site: {0}' -f $webUrl)
    }
    $sitePath = '{0}/{1}' -f $segments[0], $segments[1]

    $site = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/sites/{0}:/{1}?$select=id' -f $uri.Host, $sitePath)
    $siteId = [string](Get-GraphValue -Item $site -Key 'id')
    if ([string]::IsNullOrWhiteSpace($siteId)) {
        throw ('Could not resolve a site ID from {0}' -f $webUrl)
    }
    return $siteId
}

function Get-RecycleBinItem {
    <#
    .SYNOPSIS
        Lists every item in a site's first-stage recycle bin, following pagination.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [string]$Activity = 'Reading the recycle bin'
    )

    $uri = 'https://graph.microsoft.com/beta/sites/{0}/recycleBin/items?$top=200' -f $SiteId
    $count = 0

    while ($uri) {
        $response = Invoke-ToolkitGraphRequest -Uri $uri
        foreach ($raw in (Get-GraphValue -Item $response -Key 'value' -Default @())) {
            $count++
            $deletedBy = Get-GraphValue -Item $raw -Key 'deletedBy'
            $deletedByUser = if ($deletedBy) { Get-GraphValue -Item $deletedBy -Key 'user' } else { $null }

            [pscustomobject]@{
                Id                  = [string](Get-GraphValue -Item $raw -Key 'id')
                Name                = [string](Get-GraphValue -Item $raw -Key 'name')
                Size                = [int64](Get-GraphValue -Item $raw -Key 'size' -Default 0)
                DeletedDateTime     = [string](Get-GraphValue -Item $raw -Key 'deletedDateTime')
                DeletedFromLocation = [string](Get-GraphValue -Item $raw -Key 'deletedFromLocation')
                DeletedBy           = [string](Get-GraphValue -Item $deletedByUser -Key 'displayName')
            }
        }

        Write-Progress -Activity $Activity -Status ('{0} items read' -f $count)
        $uri = [string](Get-GraphValue -Item $response -Key '@odata.nextLink' -Default '')
    }

    Write-Progress -Activity $Activity -Completed
    Write-ToolkitLog ('Recycle bin holds {0} item(s).' -f $count) -Level INFO -NoConsole
}

function Get-DriveManifestLookup {
    <#
    .SYNOPSIS
        Builds path and name lookups from a stage 2 manifest, for "is it back already?".
    #>
    [CmdletBinding()]
    param([string]$ManifestPath)

    $lookup = [pscustomobject]@{
        Paths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        Names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    if ([string]::IsNullOrWhiteSpace($ManifestPath) -or -not (Test-Path -LiteralPath $ManifestPath)) {
        return $lookup
    }

    try {
        foreach ($row in (Import-Csv -LiteralPath $ManifestPath)) {
            $path = ConvertTo-NormalizedRelativePath -Path ([string]$row.RelativePath)
            if ($path) { [void]$lookup.Paths.Add($path) }
            $name = [string]$row.Name
            if ($name) { [void]$lookup.Names.Add($name) }
        }
        Write-ToolkitLog ('Loaded {0} current drive path(s) from {1}.' -f $lookup.Paths.Count, $ManifestPath) -Level INFO
    }
    catch {
        Write-ToolkitLog ('Could not read the manifest {0}: {1}' -f $ManifestPath, $_.Exception.Message) -Level WARN
    }

    return $lookup
}

function ConvertTo-RecycleBinPath {
    <#
    .SYNOPSIS
        Rebuilds a drive-relative path from deletedFromLocation plus the item name.
    .DESCRIPTION
        deletedFromLocation is relative to the document library, and sometimes
        carries a leading 'Documents/' segment. Both forms are returned so a
        manifest match can be attempted either way.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$DeletedFromLocation,
        [Parameter(Mandatory)][string]$Name
    )

    $location = ConvertTo-NormalizedRelativePath -Path ([string]$DeletedFromLocation)
    $full = Join-RelativePath -Parent $location -Child $Name

    $trimmed = $full
    foreach ($prefix in @('Documents/', 'Shared Documents/')) {
        if ($full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $trimmed = $full.Substring($prefix.Length)
            break
        }
    }

    return [pscustomobject]@{
        Path        = $full
        TrimmedPath = $trimmed
    }
}

function ConvertTo-RecycleBinReportRow {
    <#
    .SYNOPSIS
        Classifies each deleted item as debris, already-back, or worth recovering.

    .DESCRIPTION
        Uses the same conflict-name patterns as stages 3 and 4, so the three
        stages cannot disagree about what counts as sync noise. A conflict-named
        item is only written off as debris when the file it was copied from is
        actually in the drive now; otherwise it is surfaced, because it may be
        the only copy left.

        When no stage 2 manifest is available there is nothing to check against,
        so classification falls back to the filename alone - recorded in
        ClassificationBasis so the report never overstates what it knows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$Items,
        [Parameter(Mandatory)]$Lookup
    )

    $hasManifest = $Lookup.Paths.Count -gt 0
    $basis = if ($hasManifest) { 'ManifestChecked' } else { 'PatternOnly' }
    $rows = [System.Collections.Generic.List[psobject]]::new()

    foreach ($item in $Items) {
        $paths = ConvertTo-RecycleBinPath -DeletedFromLocation $item.DeletedFromLocation -Name $item.Name
        $info = Get-ConflictNameInfo -Name $item.Name

        $inDriveByPath = $hasManifest -and ($Lookup.Paths.Contains($paths.Path) -or $Lookup.Paths.Contains($paths.TrimmedPath))
        $inDriveByName = $hasManifest -and $Lookup.Names.Contains($item.Name)

        $isDebris = $false
        if ($info.IsCandidate) {
            if ($hasManifest) {
                # Debris only when the file it was copied from is in the drive now.
                $isDebris = $Lookup.Names.Contains($info.BaseName)
            }
            else {
                # No drive data to check against; trust only the distinctive shapes.
                $isDebris = ($info.Confidence -eq 'High')
            }
        }

        $classification = if ($isDebris) { 'ConflictCopyDeleted' }
            elseif ($inDriveByPath) { 'AlreadyBackInDrive' }
            elseif ($inDriveByName) { 'PossiblyBackInDrive' }
            else { 'RecoverableCandidate' }

        $matchedBy = if ($inDriveByPath) { 'Path' } elseif ($inDriveByName) { 'Name' } else { 'None' }

        $note = switch ($classification) {
            'ConflictCopyDeleted'  { 'Sync-conflict copy; deleting it was the point' }
            'AlreadyBackInDrive'   { 'A file with this path is in the drive now - nothing to recover' }
            'PossiblyBackInDrive'  { 'Same filename exists elsewhere in the drive - check before restoring' }
            default                { 'Not in the current drive - restore by hand if it is still wanted' }
        }

        $rows.Add([pscustomobject]@{
            Classification      = $classification
            Name                = $item.Name
            RelativePath        = $paths.TrimmedPath
            SizeBytes           = $item.Size
            DeletedDateTime     = $item.DeletedDateTime
            DeletedFromLocation = $item.DeletedFromLocation
            DeletedBy           = $item.DeletedBy
            PatternType         = $info.PatternType
            Confidence          = $info.Confidence
            MatchedBy           = $matchedBy
            ClassificationBasis = $basis
            ItemId              = $item.Id
            Note                = $note
        })
    }

    return $rows
}

function Invoke-RecycleBinInventory {
    <#
    .SYNOPSIS
        Menu option 6 - inventory the recycle bin and report what can be recovered.
    #>
    [CmdletBinding()]
    param([string]$UserId)

    Write-ToolkitHeader 'OneDrive Recycle Bin Inventory'

    if (-not (Connect-ToolkitGraph)) { return $null }

    $config = Get-ToolkitConfig
    if (-not $UserId) {
        $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
    }

    try {
        $user = Resolve-OneDriveUser -UserId $UserId
        Set-ToolkitConfigValue -Name 'TargetUserId' -Value $UserId | Out-Null
    }
    catch {
        Write-ToolkitLog ('Could not resolve the user or their drive: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    Write-Host ''
    Write-Host ('  User       : {0} <{1}>' -f $user.DisplayName, $user.UserPrincipalName)

    try {
        $siteId = Resolve-OneDriveSiteId -UserId $UserId
        Write-Host ('  Site ID    : {0}' -f $siteId)
    }
    catch {
        Write-ToolkitLog ('Could not resolve the site behind this OneDrive: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    $manifestPath = Read-ToolkitValue -Prompt 'Stage 2 manifest to compare against (blank to skip)' `
        -Default ([string]$config['LastManifestPath']) -AllowEmpty
    $lookup = Get-DriveManifestLookup -ManifestPath $manifestPath

    if ($lookup.Paths.Count -eq 0) {
        Write-ToolkitLog 'No manifest loaded - items will be classified on their filenames alone. Run stage 2 first for a more reliable split.' -Level WARN
    }

    Write-Host ''
    Write-ToolkitLog 'Reading the recycle bin (beta endpoint)...' -Level INFO

    $items = [System.Collections.Generic.List[psobject]]::new()
    try {
        Get-RecycleBinItem -SiteId $siteId | ForEach-Object { $items.Add($_) }
    }
    catch {
        Write-ToolkitLog ('Could not read the recycle bin: {0}' -f $_.Exception.Message) -Level ERROR
        Write-ToolkitLog 'This endpoint is beta-only and is not enabled in every tenant. The bin is always readable in the OneDrive web UI.' -Level WARN
        return $null
    }

    if ($items.Count -eq 0) {
        Write-ToolkitLog 'The recycle bin is empty.' -Level SUCCESS
        return $null
    }

    $rows = ConvertTo-RecycleBinReportRow -Items $items -Lookup $lookup

    $reportPath = Get-ToolkitReportPath -BaseName 'recycle-bin-inventory' -Extension 'csv'
    $rows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8
    Set-ToolkitConfigValue -Name 'LastRecycleBinReportPath' -Value $reportPath | Out-Null

    $recoverable = @($rows | Where-Object { $_.Classification -eq 'RecoverableCandidate' })
    $totalBytes = ($rows | Measure-Object -Property SizeBytes -Sum).Sum
    if (-not $totalBytes) { $totalBytes = 0 }
    $recoverableBytes = ($recoverable | Measure-Object -Property SizeBytes -Sum).Sum
    if (-not $recoverableBytes) { $recoverableBytes = 0 }

    Write-Host ''
    Write-ToolkitHeader 'Recycle bin summary'
    Write-Host ('  Items in the bin       : {0} ({1})' -f $rows.Count, (Format-ByteSize $totalBytes))
    foreach ($category in @('RecoverableCandidate', 'PossiblyBackInDrive', 'AlreadyBackInDrive', 'ConflictCopyDeleted')) {
        $count = @($rows | Where-Object { $_.Classification -eq $category }).Count
        $colour = switch ($category) {
            'RecoverableCandidate' { 'Green' }
            'PossiblyBackInDrive'  { 'Yellow' }
            default                { 'Gray' }
        }
        Write-Host ('  {0,-23}: {1}' -f $category, $count) -ForegroundColor $colour
    }

    $dates = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DeletedDateTime) } | Sort-Object DeletedDateTime)
    if ($dates.Count -gt 0) {
        Write-Host ''
        Write-Host ('  Deleted between        : {0} and {1}' -f $dates[0].DeletedDateTime, $dates[-1].DeletedDateTime)
    }

    Write-Host ''
    Write-Host ('  Worth a look : {0} item(s), {1}' -f $recoverable.Count, (Format-ByteSize $recoverableBytes)) -ForegroundColor Cyan
    Write-Host ('  Report       : {0}' -f $reportPath)

    if ($recoverable.Count -gt 0) {
        Write-Host ''
        Write-Host 'First 20 recoverable candidates:' -ForegroundColor Cyan
        $recoverable | Sort-Object -Property @{E = { [int64]$_.SizeBytes }} -Descending | Select-Object -First 20 |
            Format-Table -AutoSize @{N = 'Deleted'; E = { $_.DeletedDateTime } },
                                   @{N = 'Path'; E = { $_.RelativePath } },
                                   @{N = 'Size'; E = { Format-ByteSize ([int64]$_.SizeBytes) } } |
            Out-String | Write-Host
    }

    Write-Host ''
    Write-Host '  Note: Graph exposes no file content for recycle bin items, so their' -ForegroundColor DarkGray
    Write-Host '  contents cannot be downloaded. Restore the items above from the OneDrive' -ForegroundColor DarkGray
    Write-Host '  web UI, then re-run stage 2 to pull them down with the rest of the drive.' -ForegroundColor DarkGray

    Write-ToolkitLog ('Recycle bin inventory complete: {0} item(s), {1} recoverable candidate(s), report at {2}.' -f
        $rows.Count, $recoverable.Count, $reportPath) -Level SUCCESS -NoConsole

    return [pscustomobject]@{
        ReportPath  = $reportPath
        ItemCount   = $rows.Count
        Recoverable = $recoverable.Count
        TotalBytes  = $totalBytes
        Rows        = $rows
    }
}

Export-ModuleMember -Function @(
    'Invoke-RecycleBinInventory'
    'Resolve-OneDriveSiteId'
    'Get-RecycleBinItem'
    'Get-DriveManifestLookup'
    'ConvertTo-RecycleBinPath'
    'ConvertTo-RecycleBinReportRow'
)
