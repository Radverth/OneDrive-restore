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
# Invoke-DriveItemDownload is shared with stage 2 rather than reimplemented here.
Import-Module (Join-Path $PSScriptRoot 'DownloadOneDrive.psm1') -DisableNameChecking

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

    Write-Host ''
    Write-Host '  WHAT THIS DOES' -ForegroundColor Cyan
    Write-Host '  Lists what is in the user''s recycle bin and works out which items look' -ForegroundColor Gray
    Write-Host '  worth recovering, as opposed to conflict copies that were meant to go.' -ForegroundColor Gray
    Write-Host '  Read-only: nothing is restored, deleted or changed. Use option 7 to' -ForegroundColor Gray
    Write-Host '  actually download the files.' -ForegroundColor Gray

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

# ---------------------------------------------------------------------------
# SharePoint REST access
# ---------------------------------------------------------------------------
#
# Downloading recycle bin content needs a different door than the rest of the
# toolkit. Graph exposes no content stream for a recycleBinItem, and its restore
# action is OneDrive Personal only, so for OneDrive for Business the bytes are
# only reachable by restoring through SharePoint REST and then downloading the
# restored file through Graph as an ordinary driveItem.
#
# SharePoint REST rejects client-secret app-only tokens, so the same certificate
# that authenticates Graph is used to sign a JWT client assertion for a
# SharePoint-audience token.

function Get-ToolkitCertificate {
    <#
    .SYNOPSIS
        Finds the toolkit's certificate (with its private key) in the Windows store.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Thumbprint)

    $clean = $Thumbprint -replace '[^0-9A-Fa-f]', ''
    foreach ($store in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
        $certificate = Get-ChildItem -Path $store -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $clean } | Select-Object -First 1
        if ($certificate) {
            if (-not $certificate.HasPrivateKey) {
                throw ('Certificate {0} was found in {1} but has no private key, so it cannot sign a token.' -f $clean, $store)
            }
            return $certificate
        }
    }

    throw ('Certificate {0} was not found in the current user or local machine store. Run stage 1 on this machine, or import the .pfx.' -f $clean)
}

function ConvertTo-Base64Url {
    [CmdletBinding()]
    param([Parameter(Mandatory)][byte[]]$Bytes)

    return ([Convert]::ToBase64String($Bytes).TrimEnd('=') -replace '\+', '-' -replace '/', '_')
}

function New-ClientAssertionJwt {
    <#
    .SYNOPSIS
        Builds an RS256 JWT client assertion signed with the toolkit certificate.
    .DESCRIPTION
        This is the certificate-credential flow Entra expects in place of a client
        secret. Kept separate from the HTTP call so it can be verified offline.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$TenantId,
        [int]$LifetimeMinutes = 10
    )

    $audience = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $TenantId
    $now = [System.DateTimeOffset]::UtcNow

    $header = [ordered]@{
        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64Url -Bytes $Certificate.GetCertHash()
    }

    $payload = [ordered]@{
        aud = $audience
        iss = $AppId
        sub = $AppId
        jti = [guid]::NewGuid().ToString()
        nbf = $now.AddMinutes(-5).ToUnixTimeSeconds()
        exp = $now.AddMinutes($LifetimeMinutes).ToUnixTimeSeconds()
    }

    $encodedHeader = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json -Compress)))
    $encodedPayload = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress)))
    $signingInput = '{0}.{1}' -f $encodedHeader, $encodedPayload

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw 'Could not obtain the certificate private key for signing.' }

    try {
        $signature = $rsa.SignData(
            [System.Text.Encoding]::UTF8.GetBytes($signingInput),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    finally {
        $rsa.Dispose()
    }

    return '{0}.{1}' -f $signingInput, (ConvertTo-Base64Url -Bytes $signature)
}

function Get-SharePointToken {
    <#
    .SYNOPSIS
        Gets an app-only access token for a SharePoint host using the certificate.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)][string]$ResourceHost
    )

    $assertion = New-ClientAssertionJwt -Certificate $Certificate -AppId $AppId -TenantId $TenantId

    $body = @{
        client_id             = $AppId
        scope                 = 'https://{0}/.default' -f $ResourceHost
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $assertion
        grant_type            = 'client_credentials'
    }

    $uri = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $TenantId

    try {
        $response = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    }
    catch {
        $detail = $_.Exception.Message
        try {
            $stream = $_.Exception.Response.Content.ReadAsStringAsync().Result
            if ($stream) { $detail = $stream }
        }
        catch { Write-Verbose 'No response body on the token error.' }
        throw ('Could not get a SharePoint token: {0}' -f $detail)
    }

    $token = [string](Get-GraphValue -Item $response -Key 'access_token')
    if ([string]::IsNullOrWhiteSpace($token)) {
        throw 'The token endpoint returned no access_token.'
    }
    return $token
}

function Invoke-SharePointRestRequest {
    <#
    .SYNOPSIS
        Calls a SharePoint REST endpoint with retry and JSON handling.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Token,
        [ValidateSet('GET', 'POST', 'DELETE')][string]$Method = 'GET',
        [int]$MaxRetry = 5
    )

    $headers = @{
        Authorization = 'Bearer {0}' -f $Token
        Accept        = 'application/json;odata=nometadata'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $previous = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try {
                $response = Invoke-WebRequest -Uri $Uri -Method $Method -Headers $headers -TimeoutSec 300 -ErrorAction Stop
            }
            finally {
                $ProgressPreference = $previous
            }

            if ($response.Content -and $response.Content.Trim()) {
                try { return ($response.Content | ConvertFrom-Json -AsHashtable) }
                catch { return $null }   # restore() answers 204 with an empty body
            }
            return $null
        }
        catch {
            $detail = Get-GraphErrorDetail -ErrorRecord $_
            $retryable = $detail.StatusCode -in @(429, 500, 502, 503, 504) -or $detail.StatusCode -eq 0
            if (-not $retryable -or $attempt -gt $MaxRetry) { throw }

            $delay = if ($detail.RetryAfter -and $detail.RetryAfter -gt 0) { [math]::Min($detail.RetryAfter, 300) }
                     else { [math]::Min([math]::Pow(2, $attempt), 60) }
            Write-ToolkitLog ('SharePoint returned HTTP {0}; retry {1}/{2} in {3}s.' -f $detail.StatusCode, $attempt, $MaxRetry, $delay) -Level WARN -NoConsole
            Start-Sleep -Seconds $delay
        }
    }
}

function Resolve-OneDriveSiteUrl {
    <#
    .SYNOPSIS
        Derives the OneDrive site collection URL and host from the drive's webUrl.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)

    $drive = Invoke-ToolkitGraphRequest -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive?$select=id,webUrl' -f [uri]::EscapeDataString($UserId))
    $webUrl = [string](Get-GraphValue -Item $drive -Key 'webUrl')
    if ([string]::IsNullOrWhiteSpace($webUrl)) {
        throw 'The drive did not return a webUrl, so the SharePoint site URL cannot be derived.'
    }

    $uri = [uri]$webUrl
    $segments = @($uri.AbsolutePath.Trim('/') -split '/')
    if ($segments.Count -lt 2) {
        throw ('Unexpected OneDrive web URL: {0}' -f $webUrl)
    }

    return [pscustomobject]@{
        Host    = $uri.Host
        SiteUrl = ('{0}://{1}/{2}/{3}' -f $uri.Scheme, $uri.Host, $segments[0], $segments[1])
        WebUrl  = $webUrl
    }
}

function ConvertTo-DriveRelativePathFromDirName {
    <#
    .SYNOPSIS
        Turns a recycle bin DirName into a path relative to the drive root.
    .DESCRIPTION
        DirName is server-relative, e.g. 'personal/user_contoso_com/Documents/Projects'.
        Everything up to and including the document library segment is stripped so
        the result lines up with Graph's drive-relative paths.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$DirName,
        [Parameter(Mandatory)][string]$LeafName
    )

    $dir = ConvertTo-NormalizedRelativePath -Path ([string]$DirName)
    $segments = @($dir -split '/' | Where-Object { $_ })

    $libraryIndex = -1
    for ($i = 0; $i -lt $segments.Count; $i++) {
        if ($segments[$i] -in @('Documents', 'Shared Documents')) { $libraryIndex = $i; break }
    }

    $relativeSegments = if ($libraryIndex -ge 0) {
        @($segments | Select-Object -Skip ($libraryIndex + 1))
    }
    elseif ($segments.Count -ge 2 -and $segments[0] -eq 'personal') {
        # No library segment present; drop the /personal/<user> prefix.
        @($segments | Select-Object -Skip 2)
    }
    else {
        $segments
    }

    return (Join-RelativePath -Parent ($relativeSegments -join '/') -Child $LeafName)
}

function Get-SharePointRecycleBinItem {
    <#
    .SYNOPSIS
        Lists recycle bin entries through SharePoint REST.
    .DESCRIPTION
        Used for the download path rather than the Graph beta endpoint, because
        restore() is addressed by the recycle bin entry GUID this call returns -
        taking both from the same API removes any question of ID mapping.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$Token,
        [ValidateSet('FirstStage', 'SecondStage', 'Both')][string]$Stage = 'FirstStage',
        [string]$Activity = 'Reading the recycle bin'
    )

    $select = 'Id,Title,LeafName,DirName,Size,DeletedDate,ItemType,ItemState,DeletedByName'
    $uri = '{0}/_api/web/RecycleBin?$select={1}&$top=1000' -f $SiteUrl.TrimEnd('/'), $select
    $count = 0

    while ($uri) {
        $response = Invoke-SharePointRestRequest -Uri $uri -Token $Token
        if (-not $response) { break }

        foreach ($raw in (Get-GraphValue -Item $response -Key 'value' -Default @())) {
            $state = [int](Get-GraphValue -Item $raw -Key 'ItemState' -Default 1)
            $keep = switch ($Stage) {
                'FirstStage'  { $state -eq 1 }
                'SecondStage' { $state -eq 2 }
                default       { $true }
            }
            if (-not $keep) { continue }

            $count++
            $leaf = [string](Get-GraphValue -Item $raw -Key 'LeafName')
            $dir = [string](Get-GraphValue -Item $raw -Key 'DirName')

            [pscustomobject]@{
                Id           = [string](Get-GraphValue -Item $raw -Key 'Id')
                Name         = $leaf
                Title        = [string](Get-GraphValue -Item $raw -Key 'Title')
                DirName      = $dir
                RelativePath = ConvertTo-DriveRelativePathFromDirName -DirName $dir -LeafName $leaf
                Size         = [int64](Get-GraphValue -Item $raw -Key 'Size' -Default 0)
                DeletedDate  = [string](Get-GraphValue -Item $raw -Key 'DeletedDate')
                DeletedBy    = [string](Get-GraphValue -Item $raw -Key 'DeletedByName')
                ItemType     = [int](Get-GraphValue -Item $raw -Key 'ItemType' -Default 0)
                ItemState    = $state
            }
        }

        Write-Progress -Activity $Activity -Status ('{0} entries read' -f $count)
        $next = Get-GraphValue -Item $response -Key 'odata.nextLink' -Default ''
        if (-not $next) { $next = Get-GraphValue -Item $response -Key '@odata.nextLink' -Default '' }
        $uri = [string]$next
    }

    Write-Progress -Activity $Activity -Completed
    Write-ToolkitLog ('SharePoint recycle bin returned {0} matching entry/entries.' -f $count) -Level INFO -NoConsole
}

function Restore-SharePointRecycleBinItem {
    <#
    .SYNOPSIS
        Restores one recycle bin entry to its original location.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$ItemId
    )

    $uri = "{0}/_api/web/RecycleBin('{1}')/restore()" -f $SiteUrl.TrimEnd('/'), $ItemId
    Invoke-SharePointRestRequest -Uri $uri -Token $Token -Method POST | Out-Null
}

# ---------------------------------------------------------------------------
# Download the contents of the recycle bin
# ---------------------------------------------------------------------------

function Select-RecycleBinDownloadTarget {
    <#
    .SYNOPSIS
        Picks which recycle bin entries to download, and puts them in restore order.
    .DESCRIPTION
        Folders go first, shallowest first, so a file whose parent folder was also
        deleted has somewhere to land. Files then follow in path order.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$Items,
        [ValidateSet('All', 'NotInDrive', 'Pattern')][string]$Scope = 'All',
        $Lookup,
        [string]$Pattern = '',
        [switch]$IncludeFolders
    )

    $selected = foreach ($item in $Items) {
        # ItemType 1 = File, 5 = Folder. Everything else (versions, list items) is
        # not a downloadable file and is skipped.
        if ($item.ItemType -eq 5) {
            if ($IncludeFolders) { $item }
            continue
        }
        if ($item.ItemType -ne 1) { continue }

        $keep = switch ($Scope) {
            'NotInDrive' {
                -not ($Lookup -and ($Lookup.Paths.Contains($item.RelativePath) -or $Lookup.Names.Contains($item.Name)))
            }
            'Pattern' {
                -not [string]::IsNullOrWhiteSpace($Pattern) -and $item.Name -like $Pattern
            }
            default { $true }
        }

        if ($keep) { $item }
    }

    # Leading comma so an empty selection stays an empty array rather than being
    # unrolled to nothing, which would make the caller's .Count throw.
    return ,@($selected | Sort-Object `
        @{ Expression = { if ($_.ItemType -eq 5) { 0 } else { 1 } } },
        @{ Expression = { @($_.RelativePath -split '/').Count } },
        @{ Expression = { $_.RelativePath } })
}

function Get-RestoredDriveItem {
    <#
    .SYNOPSIS
        Looks up a drive item by path, retrying while SharePoint catches up.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$RelativePath,
        [int]$MaxAttempt = 5,
        [int]$DelaySeconds = 2
    )

    $encodedUser = [uri]::EscapeDataString($UserId)
    $encodedPath = (($RelativePath -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    $uri = 'https://graph.microsoft.com/v1.0/users/{0}/drive/root:/{1}' -f $encodedUser, $encodedPath

    for ($attempt = 1; $attempt -le $MaxAttempt; $attempt++) {
        $raw = Invoke-ToolkitGraphRequest -Uri $uri -AllowNotFound
        if ($raw) {
            $parent = ConvertTo-NormalizedRelativePath -Path (Split-Path -Parent $RelativePath)
            return (ConvertFrom-GraphDriveItem -Item $raw -ParentPath $parent)
        }
        if ($attempt -lt $MaxAttempt) { Start-Sleep -Seconds $DelaySeconds }
    }

    return $null
}

function Invoke-RecycleBinDownload {
    <#
    .SYNOPSIS
        Menu option 7 - download the contents of the user's recycle bin.

    .DESCRIPTION
        Graph exposes no content stream for a recycle bin item, so each file is
        restored through SharePoint REST, downloaded through Graph as an ordinary
        driveItem, and then (by default) deleted again so the drive is left as it
        was found. That happens one file at a time: on an account that is already
        over quota, restoring the whole bin at once could push it further over,
        whereas this holds at most one restored file in the drive at any moment.
    #>
    [CmdletBinding()]
    param(
        [string]$UserId,
        [string]$Destination,
        [switch]$KeepRestored,
        [int]$BatchLimit = 0
    )

    Write-ToolkitHeader 'Download Recycle Bin Contents'

    Write-Host ''
    Write-Host '  WHAT THIS DOES' -ForegroundColor Cyan
    Write-Host '  Saves deleted files to a folder on this PC. OneDrive will not hand over' -ForegroundColor Gray
    Write-Host '  a deleted file directly, so each one is restored to the drive, then' -ForegroundColor Gray
    Write-Host '  downloaded, then deleted again - one file at a time, so an account that' -ForegroundColor Gray
    Write-Host '  is already over quota does not fill up. You will see the full plan, and' -ForegroundColor Gray
    Write-Host '  how much space it needs, before anything happens.' -ForegroundColor Gray

    if (-not (Connect-ToolkitGraph)) { return $null }
    $config = Get-ToolkitConfig

    if (-not $UserId) {
        $UserId = Read-ToolkitValue -Prompt 'Target user (UPN or object ID)' -Default ([string]$config['TargetUserId'])
    }

    try {
        $user = Resolve-OneDriveUser -UserId $UserId
        Set-ToolkitConfigValue -Name 'TargetUserId' -Value $UserId | Out-Null
        $site = Resolve-OneDriveSiteUrl -UserId $UserId
    }
    catch {
        Write-ToolkitLog ('Could not resolve the user or their site: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    Write-Host ''
    Write-Host ('  User       : {0} <{1}>' -f $user.DisplayName, $user.UserPrincipalName)
    Write-Host ('  Site       : {0}' -f $site.SiteUrl)
    Write-Host ('  Quota      : {0} used of {1} ({2} free)' -f
        (Format-ByteSize $user.QuotaUsed), (Format-ByteSize $user.QuotaTotal), (Format-ByteSize $user.QuotaRemaining))

    # SharePoint REST needs its own token, signed with the same certificate.
    try {
        $certificate = Get-ToolkitCertificate -Thumbprint ([string]$config['CertificateThumbprint'])
        Write-ToolkitLog ('Requesting a SharePoint token for {0}...' -f $site.Host) -Level INFO
        $token = Get-SharePointToken -TenantId ([string]$config['TenantId']) -AppId ([string]$config['AppId']) `
            -Certificate $certificate -ResourceHost $site.Host
        Write-ToolkitLog 'SharePoint token acquired.' -Level SUCCESS
    }
    catch {
        Write-ToolkitLog ('Could not authenticate to SharePoint: {0}' -f $_.Exception.Message) -Level ERROR
        Write-ToolkitLog 'Recycle bin restore needs the SharePoint application permission as well as the Graph ones. Re-run stage 1 and accept the SharePoint permission prompt, then grant admin consent.' -Level WARN
        return $null
    }

    try {
        $items = @(Get-SharePointRecycleBinItem -SiteUrl $site.SiteUrl -Token $token -Stage 'FirstStage')
    }
    catch {
        Write-ToolkitLog ('Could not read the recycle bin: {0}' -f $_.Exception.Message) -Level ERROR
        return $null
    }

    if ($items.Count -eq 0) {
        Write-ToolkitLog 'The recycle bin is empty.' -Level SUCCESS
        return $null
    }

    $manifestPath = Read-ToolkitValue -Prompt 'Stage 2 manifest to compare against (blank to skip)' `
        -Default ([string]$config['LastManifestPath']) -AllowEmpty
    $lookup = Get-DriveManifestLookup -ManifestPath $manifestPath

    Write-Host ''
    Write-Host '  What should be downloaded?'
    Write-Host '    1. Everything in the recycle bin'
    Write-Host '    2. Only items that are not in the drive now (needs a manifest)'
    Write-Host '    3. Only items matching a filename pattern'
    Write-Host ''
    $choice = Read-ToolkitValue -Prompt 'Choice' -Default '1'

    $scope = switch ($choice) { '2' { 'NotInDrive' } '3' { 'Pattern' } default { 'All' } }
    $pattern = ''
    if ($scope -eq 'Pattern') { $pattern = Read-ToolkitValue -Prompt 'Filename pattern (e.g. *.xlsx)' }
    if ($scope -eq 'NotInDrive' -and $lookup.Paths.Count -eq 0) {
        Write-ToolkitLog 'No manifest loaded, so "not in the drive now" cannot be determined. Run stage 2 first.' -Level ERROR
        return $null
    }

    $targets = Select-RecycleBinDownloadTarget -Items $items -Scope $scope -Lookup $lookup -Pattern $pattern
    if ($targets.Count -eq 0) {
        Write-ToolkitLog 'Nothing in the recycle bin matches that selection.' -Level WARN
        return $null
    }

    # Offer a trial batch: restoring is the riskiest part, so proving it works on
    # one or two files before doing hundreds is worth the extra run.
    $targets = Select-ToolkitBatch -Items $targets -Noun 'deleted file' -Limit $BatchLimit

    $totalBytes = ($targets | Measure-Object -Property Size -Sum).Sum
    if (-not $totalBytes) { $totalBytes = 0 }
    $largest = ($targets | Measure-Object -Property Size -Maximum).Maximum
    if (-not $largest) { $largest = 0 }

    if (-not $Destination) {
        $Destination = Read-ToolkitDirectory -Prompt 'Local destination folder for the recycle bin download' `
            -Default ([string]$config['RecycleBinDownloadPath']) -CreateIfMissing
    }
    if (-not $Destination) {
        Write-ToolkitLog 'No destination chosen; download cancelled.' -Level WARN
        return $null
    }

    $leaf = 'recyclebin-{0}-{1}' -f ($user.UserPrincipalName -replace '[^A-Za-z0-9._-]', '_'), (Get-Date -Format 'yyyyMMdd-HHmmss')
    $Destination = Join-Path $Destination $leaf
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    Set-ToolkitConfigValue -Name 'RecycleBinDownloadPath' -Value (Split-Path -Parent $Destination) | Out-Null

    $freeSpace = Get-FreeDiskSpace -Path $Destination
    if ($null -ne $freeSpace -and $freeSpace -lt $totalBytes) {
        Write-ToolkitLog ('Selection is {0} but only {1} is free at the destination.' -f (Format-ByteSize $totalBytes), (Format-ByteSize $freeSpace)) -Level WARN
        if (-not (Confirm-ToolkitAction -Prompt 'Continue anyway?')) { return $null }
    }

    $putBack = -not $KeepRestored
    if (-not $KeepRestored) {
        Write-Host ''
        Write-Host '  Each file is restored to the drive, downloaded, then deleted again so the' -ForegroundColor Gray
        Write-Host '  drive is left as it was. Answer no to leave the restored files in place.' -ForegroundColor Gray
        $putBack = Confirm-ToolkitAction -Prompt 'Put each file back in the recycle bin after downloading?' -DefaultYes
    }

    Write-Host ''
    Write-ToolkitHeader 'Download plan'
    Write-Host ('  Items to download : {0}' -f $targets.Count)
    Write-Host ('  Total size        : {0}' -f (Format-ByteSize $totalBytes))
    Write-Host ('  Largest single    : {0}' -f (Format-ByteSize $largest))
    Write-Host ('  Destination       : {0}' -f $Destination)
    if ($putBack) {
        Write-Host ('  Drive impact      : one file at a time, put back afterwards') -ForegroundColor Green
        Write-Host ('  Peak extra quota  : {0}' -f (Format-ByteSize $largest)) -ForegroundColor Green
    }
    else {
        Write-Host ('  Drive impact      : {0} restored and LEFT in the drive' -f (Format-ByteSize $totalBytes)) -ForegroundColor Yellow
        if ($user.QuotaRemaining -gt 0 -and $totalBytes -gt $user.QuotaRemaining) {
            Write-ToolkitLog ('That is more than the {0} of quota remaining - the drive will go over quota.' -f (Format-ByteSize $user.QuotaRemaining)) -Level WARN
        }
    }

    Write-Host ''
    Write-ToolkitLog 'This writes to the live drive: files are restored, then downloaded, then deleted again.' -Level WARN
    if (-not (Confirm-ToolkitAction -Prompt 'Proceed?')) {
        Write-ToolkitLog 'Recycle bin download cancelled.' -Level WARN
        return $null
    }
    $typed = Read-ToolkitValue -Prompt 'Type RESTORE in capitals to confirm' -AllowEmpty
    if ($typed -cne 'RESTORE') {
        Write-ToolkitLog 'Confirmation not matched; nothing was restored.' -Level WARN
        return $null
    }

    $results = [System.Collections.Generic.List[psobject]]::new()
    $downloaded = 0
    $failed = 0
    $bytesDone = [int64]0
    $index = 0

    foreach ($target in $targets) {
        $index++
        Write-Progress -Activity 'Downloading recycle bin contents' -Id 1 `
            -Status ('{0}/{1} - {2}' -f $index, $targets.Count, $target.RelativePath) `
            -PercentComplete ([int](($index / [math]::Max($targets.Count, 1)) * 100))

        $restoreStatus = 'Restored'
        $downloadStatus = 'Skipped'
        $putBackStatus = 'NotRequested'
        $detail = ''
        $localPath = ''
        $hash = $null
        $restoredHere = $false

        try {
            # A folder restored earlier in this run may already have brought this
            # file back, so check the drive before restoring anything.
            $driveItem = Get-RestoredDriveItem -UserId $UserId -RelativePath $target.RelativePath -MaxAttempt 1 -DelaySeconds 0

            if ($driveItem) {
                $restoreStatus = 'AlreadyInDrive'
            }
            else {
                Restore-SharePointRecycleBinItem -SiteUrl $site.SiteUrl -Token $token -ItemId $target.Id
                $restoredHere = $true
                $driveItem = Get-RestoredDriveItem -UserId $UserId -RelativePath $target.RelativePath
                if (-not $driveItem) {
                    throw ('Restored, but no item appeared at {0}' -f $target.RelativePath)
                }
            }

            if ($target.ItemType -eq 5) {
                $downloadStatus = 'FolderRestored'
            }
            else {
                $localPath = Get-SafeLocalPath -Root $Destination -RelativePath $target.RelativePath
                Invoke-DriveItemDownload -Item $driveItem -TargetPath $localPath -UserId $UserId
                $hash = Get-FileSha256 -Path $localPath
                $downloadStatus = 'Downloaded'
                $downloaded++
                $bytesDone += $driveItem.Size
            }

            if ($putBack -and $restoredHere -and $target.ItemType -ne 5) {
                try {
                    Invoke-ToolkitGraphRequest -Method DELETE `
                        -Uri ('https://graph.microsoft.com/v1.0/users/{0}/drive/items/{1}' -f [uri]::EscapeDataString($UserId), $driveItem.Id) | Out-Null
                    $putBackStatus = 'ReturnedToRecycleBin'
                }
                catch {
                    $putBackStatus = 'Failed'
                    $detail = 'Downloaded, but could not be deleted again: {0}' -f $_.Exception.Message
                    Write-ToolkitLog $detail -Level WARN -NoConsole
                }
            }
            elseif ($putBack -and -not $restoredHere) {
                $putBackStatus = 'LeftAlone'
                $detail = 'Was already in the drive, so it was left there'
            }
            elseif (-not $putBack) {
                $putBackStatus = 'KeptInDrive'
            }
        }
        catch {
            $failed++
            if ($restoreStatus -eq 'Restored' -and -not $restoredHere) { $restoreStatus = 'Failed' }
            $downloadStatus = 'Failed'
            $detail = $_.Exception.Message
            Write-ToolkitLog ('FAILED {0}: {1}' -f $target.RelativePath, $detail) -Level ERROR -NoConsole
        }

        $results.Add([pscustomobject]@{
            RelativePath    = $target.RelativePath
            Name            = $target.Name
            SizeBytes       = $target.Size
            DeletedDate     = $target.DeletedDate
            DeletedBy       = $target.DeletedBy
            ItemType        = if ($target.ItemType -eq 5) { 'Folder' } else { 'File' }
            RestoreStatus   = $restoreStatus
            DownloadStatus  = $downloadStatus
            PutBackStatus   = $putBackStatus
            LocalPath       = $localPath
            Sha256          = $hash
            RecycleBinId    = $target.Id
            Detail          = $detail
        })
    }

    Write-Progress -Activity 'Downloading recycle bin contents' -Id 1 -Completed

    $summaryPath = Get-ToolkitReportPath -BaseName 'recycle-bin-download' -Extension 'csv'
    $results | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
    $results | Export-Csv -LiteralPath (Join-Path $Destination '_recycle-bin-manifest.csv') -NoTypeInformation -Encoding utf8
    Set-ToolkitConfigValue -Name 'LastRecycleBinDownloadPath' -Value $summaryPath | Out-Null

    $stranded = @($results | Where-Object { $_.PutBackStatus -eq 'Failed' })

    Write-Host ''
    Write-ToolkitHeader 'Recycle bin download summary'
    Write-Host ('  Downloaded      : {0} file(s), {1}' -f $downloaded, (Format-ByteSize $bytesDone)) -ForegroundColor Green
    Write-Host ('  Failed          : {0}' -f $failed) -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    if ($putBack) {
        Write-Host ('  Put back        : {0}' -f @($results | Where-Object { $_.PutBackStatus -eq 'ReturnedToRecycleBin' }).Count)
        Write-Host ('  Left in drive   : {0}' -f $stranded.Count) -ForegroundColor $(if ($stranded.Count -gt 0) { 'Yellow' } else { 'Gray' })
    }
    Write-Host ('  Folder          : {0}' -f $Destination)
    Write-Host ('  Summary CSV     : {0}' -f $summaryPath)

    if ($stranded.Count -gt 0) {
        Write-Host ''
        Write-ToolkitLog ('{0} restored file(s) could not be deleted again and are still in the drive, using quota. They are listed in the summary CSV with PutBackStatus=Failed.' -f $stranded.Count) -Level WARN
    }

    if ($putBack -and $downloaded -gt 0) {
        Write-Host ''
        Write-Host '  Note: files put back start a fresh recycle bin retention window, and' -ForegroundColor DarkGray
        Write-Host '  appear as new entries with new IDs and a new deleted date.' -ForegroundColor DarkGray
    }

    Write-ToolkitLog ('Recycle bin download complete: {0} downloaded, {1} failed, summary at {2}.' -f $downloaded, $failed, $summaryPath) `
        -Level $(if ($failed -gt 0) { 'WARN' } else { 'SUCCESS' })

    return [pscustomobject]@{
        Destination = $Destination
        SummaryPath = $summaryPath
        Downloaded  = $downloaded
        Failed      = $failed
        Stranded    = $stranded.Count
    }
}

Export-ModuleMember -Function @(
    'Invoke-RecycleBinInventory'
    'Invoke-RecycleBinDownload'
    'Resolve-OneDriveSiteId'
    'Resolve-OneDriveSiteUrl'
    'Get-RecycleBinItem'
    'Get-DriveManifestLookup'
    'ConvertTo-RecycleBinPath'
    'ConvertTo-RecycleBinReportRow'
    'Get-ToolkitCertificate'
    'ConvertTo-Base64Url'
    'New-ClientAssertionJwt'
    'Get-SharePointToken'
    'Invoke-SharePointRestRequest'
    'ConvertTo-DriveRelativePathFromDirName'
    'Get-SharePointRecycleBinItem'
    'Restore-SharePointRecycleBinItem'
    'Select-RecycleBinDownloadTarget'
    'Get-RestoredDriveItem'
)
