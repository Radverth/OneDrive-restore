#Requires -Version 7.0
<#
.SYNOPSIS
    Offline tests for the toolkit's decision logic.

.DESCRIPTION
    Exercises everything that does not need a tenant: the conflict-pattern
    matching, duplicate grouping and classification, path safety, and a full
    stage 4 comparison over throwaway folders. No Pester dependency, so it runs
    anywhere PowerShell 7 does.

.EXAMPLE
    pwsh ./tests/Invoke-ToolkitTests.ps1
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0

function Assert-That {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowNull()]$Actual,
        [Parameter(Mandatory)][AllowNull()]$Expected
    )

    $ok = if ($null -eq $Expected) { $null -eq $Actual } else { $Actual -eq $Expected }
    if ($ok) {
        $script:Passed++
        Write-Host ('  PASS  {0}' -f $Name) -ForegroundColor Green
    }
    else {
        $script:Failed++
        Write-Host ('  FAIL  {0}' -f $Name) -ForegroundColor Red
        Write-Host ('        expected [{0}] but got [{1}]' -f $Expected, $Actual) -ForegroundColor Red
    }
}

function Assert-Throws {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Action
    )

    try {
        & $Action | Out-Null
        $script:Failed++
        Write-Host ('  FAIL  {0} (no exception thrown)' -f $Name) -ForegroundColor Red
    }
    catch {
        $script:Passed++
        Write-Host ('  PASS  {0}' -f $Name) -ForegroundColor Green
    }
}

# Redirect config/logs/reports into a sandbox so tests never touch real output.
$repoRoot = Split-Path -Parent $PSScriptRoot
$sandbox = Join-Path ([System.IO.Path]::GetTempPath()) ('odrt-tests-{0}' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $sandbox -Force | Out-Null

$global:OneDriveRepairContext = [ordered]@{
    Root           = $sandbox
    ConfigDir      = Join-Path $sandbox 'config'
    ConfigPath     = Join-Path $sandbox 'config' 'toolkit-config.json'
    LogDir         = Join-Path $sandbox 'logs'
    ReportDir      = Join-Path $sandbox 'reports'
    LogPath        = $null
    Config         = $null
    GraphConnected = $false
    GraphTenantId  = $null
    GraphAppId     = $null
}

$moduleRoot = Join-Path $repoRoot 'modules'
foreach ($module in @('Common', 'DownloadOneDrive', 'DuplicateScanner', 'CompareDrives', 'ReconcileUpload', 'RecycleBin')) {
    Import-Module (Join-Path $moduleRoot ('{0}.psm1' -f $module)) -Force -DisableNameChecking
}

Write-Host ''
Write-Host 'Conflict filename patterns' -ForegroundColor Cyan

$numbered = Get-ConflictNameInfo -Name 'Q3 budget (1).xlsx'
Assert-That -Name 'numbered copy is a candidate'        -Actual $numbered.IsCandidate -Expected $true
Assert-That -Name 'numbered copy base name'             -Actual $numbered.BaseName    -Expected 'Q3 budget.xlsx'
Assert-That -Name 'numbered copy is high confidence'    -Actual $numbered.Confidence  -Expected 'High'
Assert-That -Name 'numbered copy needs no original'     -Actual $numbered.RequiresOriginal -Expected $false

$copy = Get-ConflictNameInfo -Name 'Plan - Copy.docx'
Assert-That -Name 'copy suffix is a candidate'          -Actual $copy.IsCandidate -Expected $true
Assert-That -Name 'copy suffix base name'               -Actual $copy.BaseName    -Expected 'Plan.docx'
Assert-That -Name 'copy suffix pattern type'            -Actual $copy.PatternType -Expected 'CopySuffix'

$conflicted = Get-ConflictNameInfo -Name "Notes (tom's conflicted copy 2026-04-02).md"
Assert-That -Name 'conflicted copy is a candidate'      -Actual $conflicted.IsCandidate -Expected $true
Assert-That -Name 'conflicted copy base name'           -Actual $conflicted.BaseName    -Expected 'Notes.md'

$machine = Get-ConflictNameInfo -Name 'timesheet-DESKTOP-4F2K9L1.xlsx'
Assert-That -Name 'machine suffix is a candidate'       -Actual $machine.IsCandidate      -Expected $true
Assert-That -Name 'machine suffix base name'            -Actual $machine.BaseName         -Expected 'timesheet.xlsx'
Assert-That -Name 'machine suffix needs an original'    -Actual $machine.RequiresOriginal -Expected $true
Assert-That -Name 'machine tag looks machine-like'      -Actual (Test-MachineLikeToken -Token $machine.Marker) -Expected $true

$clean = Get-ConflictNameInfo -Name 'Statement.pdf'
Assert-That -Name 'ordinary name is not a candidate'    -Actual $clean.IsCandidate -Expected $false

$hyphenated = Get-ConflictNameInfo -Name 'annual-report.docx'
Assert-That -Name 'hyphenated name is only a candidate' -Actual $hyphenated.IsCandidate      -Expected $true
Assert-That -Name 'hyphenated name needs an original'   -Actual $hyphenated.RequiresOriginal -Expected $true
Assert-That -Name 'hyphenated tag is not machine-like'  -Actual (Test-MachineLikeToken -Token $hyphenated.Marker) -Expected $false

Write-Host ''
Write-Host 'Path helpers' -ForegroundColor Cyan

Assert-That -Name 'backslashes normalise'   -Actual (ConvertTo-NormalizedRelativePath -Path '\Docs\Sub\file.txt') -Expected 'Docs/Sub/file.txt'
Assert-That -Name 'empty path normalises'   -Actual (ConvertTo-NormalizedRelativePath -Path '') -Expected ''
Assert-That -Name 'relative paths join'     -Actual (Join-RelativePath -Parent 'Docs' -Child 'a.txt') -Expected 'Docs/a.txt'
Assert-That -Name 'root-level paths join'   -Actual (Join-RelativePath -Parent '' -Child 'a.txt') -Expected 'a.txt'
Assert-Throws -Name 'traversal is refused'  -Action { Get-SafeLocalPath -Root $sandbox -RelativePath '../../etc/passwd' }
Assert-That -Name 'url segments escape'     -Actual (ConvertTo-DrivePathUrl -RelativePath 'My Docs/a b.txt') -Expected 'My%20Docs/a%20b.txt'
Assert-That -Name 'byte formatting'         -Actual (Format-ByteSize 1536) -Expected '1.50 KB'
Assert-That -Name 'small byte formatting'   -Actual (Format-ByteSize 512)  -Expected '512 B'

$parsed = ConvertTo-UtcDateTime -Value '2026-04-02T10:30:00Z'
Assert-That -Name 'graph timestamp parses'  -Actual $parsed.ToString('yyyy-MM-dd HH:mm') -Expected '2026-04-02 10:30'
Assert-That -Name 'empty timestamp is null' -Actual (ConvertTo-UtcDateTime -Value '') -Expected $null

Write-Host ''
Write-Host 'Duplicate grouping and classification' -ForegroundColor Cyan

function New-CatalogueEntry {
    param($Path, $Size, $Hash, $Created)
    return [pscustomobject]@{
        Id           = 'id-' + ($Path -replace '[^A-Za-z0-9]', '')
        Name         = Split-Path -Leaf $Path
        RelativePath = $Path
        ParentPath   = ConvertTo-NormalizedRelativePath -Path (Split-Path -Parent $Path)
        Size         = $Size
        Sha256       = $Hash
        QuickXorHash = ''
        CreatedUtc   = $Created
        ModifiedUtc  = $Created
    }
}

$catalogue = @(
    New-CatalogueEntry -Path 'Docs/report.docx'                     -Size 1000 -Hash 'AAA' -Created '2026-01-01T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/report-DESKTOP-4F2K9L1.docx'     -Size 1000 -Hash 'AAA' -Created '2026-02-01T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/report (1).docx'                 -Size 1200 -Hash 'BBB' -Created '2026-03-01T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/annual-report.docx'              -Size 900  -Hash 'CCC' -Created '2026-01-05T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/Statement.pdf'                   -Size 500  -Hash 'DDD' -Created '2026-01-06T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/orphan (2).txt'                  -Size 100  -Hash 'EEE' -Created '2026-01-07T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/orphan (3).txt'                  -Size 100  -Hash 'EEE' -Created '2026-01-08T00:00:00Z'
    New-CatalogueEntry -Path 'Docs/lonely (1).txt'                  -Size 300  -Hash 'FFF' -Created '2026-01-09T00:00:00Z'
)

$groups = Find-ConflictGroup -Catalogue $catalogue
$rows = ConvertTo-DuplicateReportRow -Groups $groups

$exact = @($rows | Where-Object { $_.Classification -eq 'ExactDuplicate' })
$conflict = @($rows | Where-Object { $_.Classification -eq 'ContentConflict' })

Assert-That -Name 'identical machine-suffix copy is an exact duplicate' `
    -Actual (@($exact | Where-Object { $_.DuplicatePath -eq 'Docs/report-DESKTOP-4F2K9L1.docx' }).Count) -Expected 1
Assert-That -Name 'differing numbered copy is a content conflict' `
    -Actual (@($conflict | Where-Object { $_.DuplicatePath -eq 'Docs/report (1).docx' }).Count) -Expected 1
Assert-That -Name 'hyphenated file with no base file is left alone' `
    -Actual (@($rows | Where-Object { $_.DuplicatePath -eq 'Docs/annual-report.docx' }).Count) -Expected 0
Assert-That -Name 'ordinary file is not reported' `
    -Actual (@($rows | Where-Object { $_.DuplicatePath -eq 'Docs/Statement.pdf' }).Count) -Expected 0
Assert-That -Name 'orphaned copies group with the oldest kept as canonical' `
    -Actual (@($exact | Where-Object { $_.DuplicatePath -eq 'Docs/orphan (3).txt' -and $_.CanonicalPath -eq 'Docs/orphan (2).txt' }).Count) -Expected 1
Assert-That -Name 'reclaimable bytes counted only for exact duplicates' `
    -Actual (($exact | Measure-Object -Property RecoverableBytes -Sum).Sum) -Expected 1100

$orphaned = @($rows | Where-Object { $_.Classification -eq 'OrphanedCopy' })
Assert-That -Name 'a copy with no original and no siblings is flagged, not dropped' `
    -Actual (@($orphaned | Where-Object { $_.DuplicatePath -eq 'Docs/lonely (1).txt' }).Count) -Expected 1
Assert-That -Name 'an orphaned copy is never counted as reclaimable' `
    -Actual (($orphaned | Measure-Object -Property RecoverableBytes -Sum).Sum) -Expected 0

$sizeOnly = Compare-DriveItemContent `
    -Left  ([pscustomobject]@{ Sha256 = ''; QuickXorHash = ''; Size = 10 }) `
    -Right ([pscustomobject]@{ Sha256 = ''; QuickXorHash = ''; Size = 10 })
Assert-That -Name 'equal size without a hash is not proof' -Actual $sizeOnly.Match -Expected $null

$sizeDiff = Compare-DriveItemContent `
    -Left  ([pscustomobject]@{ Sha256 = ''; QuickXorHash = ''; Size = 10 }) `
    -Right ([pscustomobject]@{ Sha256 = ''; QuickXorHash = ''; Size = 20 })
Assert-That -Name 'differing size means differing content' -Actual $sizeDiff.Match -Expected $false

Write-Host ''
Write-Host 'Stage 4 comparison over real folders' -ForegroundColor Cyan

$cloudRoot = Join-Path $sandbox 'cloud'
$backupRoot = Join-Path $sandbox 'backup'
New-Item -ItemType Directory -Path (Join-Path $cloudRoot 'Docs') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $backupRoot 'Docs') -Force | Out-Null

$restorePoint = [datetime]::SpecifyKind([datetime]'2026-03-01T00:00:00', [System.DateTimeKind]::Utc)
$beforeRestore = $restorePoint.AddDays(-10)
$afterRestore = $restorePoint.AddDays(10)

function Set-TestFile {
    param($Path, $Content, [datetime]$ModifiedUtc)
    Set-Content -LiteralPath $Path -Value $Content -NoNewline -Encoding utf8
    [System.IO.File]::SetLastWriteTimeUtc($Path, $ModifiedUtc)
}

# Unchanged on both sides.
Set-TestFile -Path (Join-Path $cloudRoot 'Docs/same.txt')  -Content 'identical' -ModifiedUtc $beforeRestore
Set-TestFile -Path (Join-Path $backupRoot 'Docs/same.txt') -Content 'identical' -ModifiedUtc $beforeRestore

# Edited on both sides - a genuine conflict for a human.
Set-TestFile -Path (Join-Path $cloudRoot 'Docs/edited.txt')  -Content 'cloud version' -ModifiedUtc $beforeRestore
Set-TestFile -Path (Join-Path $backupRoot 'Docs/edited.txt') -Content 'local version' -ModifiedUtc $afterRestore

# The legitimate work done after the restore point.
Set-TestFile -Path (Join-Path $backupRoot 'Docs/new-work.txt') -Content 'recovered' -ModifiedUtc $afterRestore

# Duplicate noise that must not be re-uploaded.
Set-TestFile -Path (Join-Path $backupRoot 'Docs/new-work (1).txt') -Content 'recovered' -ModifiedUtc $afterRestore
Set-TestFile -Path (Join-Path $backupRoot 'Docs/same-DESKTOP-4F2K9L1.txt') -Content 'identical' -ModifiedUtc $afterRestore

# Missing from the cloud but untouched since the restore point.
Set-TestFile -Path (Join-Path $backupRoot 'Docs/old-local.txt') -Content 'stale' -ModifiedUtc $beforeRestore

# A conflict-named file whose original exists nowhere - it may be the only copy,
# so it must reach a human rather than being silently discarded as noise.
Set-TestFile -Path (Join-Path $backupRoot 'Docs/lonely (1).txt') -Content 'possibly the only copy' -ModifiedUtc $afterRestore

# Present only in the cloud.
Set-TestFile -Path (Join-Path $cloudRoot 'Docs/cloud-only.txt') -Content 'cloud only' -ModifiedUtc $beforeRestore

$comparison = Invoke-DriveComparison -BackupPath $backupRoot -CloudPath $cloudRoot -RestorePointUtc $restorePoint

function Get-Category {
    param($Path)
    $row = $comparison.Rows | Where-Object { $_.RelativePath -eq $Path } | Select-Object -First 1
    if ($row) { return $row.Category }
    return 'missing'
}

Assert-That -Name 'identical file needs no action'          -Actual (Get-Category 'Docs/same.txt')       -Expected 'Identical'
Assert-That -Name 'new local work is an upload candidate'   -Actual (Get-Category 'Docs/new-work.txt')   -Expected 'UploadCandidate'
Assert-That -Name 'both-sides edit goes to review'          -Actual (Get-Category 'Docs/edited.txt')     -Expected 'ContentMismatch'
Assert-That -Name 'numbered copy is skipped as noise'       -Actual (Get-Category 'Docs/new-work (1).txt') -Expected 'DuplicatePatternSkip'
Assert-That -Name 'machine-suffix copy is skipped as noise' -Actual (Get-Category 'Docs/same-DESKTOP-4F2K9L1.txt') -Expected 'DuplicatePatternSkip'
Assert-That -Name 'pre-restore-point local file is review'  -Actual (Get-Category 'Docs/old-local.txt')  -Expected 'LocalOnlyBeforeRestorePoint'
Assert-That -Name 'cloud-only file is informational'        -Actual (Get-Category 'Docs/cloud-only.txt') -Expected 'CloudOnly'
Assert-That -Name 'conflict copy with no original goes to review, not skip' `
    -Actual (Get-Category 'Docs/lonely (1).txt') -Expected 'DuplicatePatternNoOriginal'
Assert-That -Name 'exactly one file is queued for upload'   -Actual $comparison.UploadCount -Expected 1
Assert-That -Name 'comparison report was written'           -Actual (Test-Path -LiteralPath $comparison.ReportPath) -Expected $true

Write-Host ''
Write-Host 'Recycle bin inventory' -ForegroundColor Cyan

$binPath = ConvertTo-RecycleBinPath -DeletedFromLocation 'Documents/Projects' -Name 'plan.docx'
Assert-That -Name 'deleted-from location rebuilds a full path' -Actual $binPath.Path        -Expected 'Documents/Projects/plan.docx'
Assert-That -Name 'the library prefix is trimmed for matching' -Actual $binPath.TrimmedPath -Expected 'Projects/plan.docx'

$rootPath = ConvertTo-RecycleBinPath -DeletedFromLocation '' -Name 'loose.txt'
Assert-That -Name 'a root-level deletion rebuilds cleanly' -Actual $rootPath.Path -Expected 'loose.txt'

function New-BinItem {
    param($Name, $Location = 'Documents', $Size = 100)
    return [pscustomobject]@{
        Id                  = 'bin-' + ($Name -replace '[^A-Za-z0-9]', '')
        Name                = $Name
        Size                = $Size
        DeletedDateTime     = '2026-04-02T09:00:00Z'
        DeletedFromLocation = $Location
        DeletedBy           = 'Admin'
    }
}

# A manifest of what the drive holds right now.
$manifestFile = Join-Path $sandbox 'manifest.csv'
@(
    [pscustomobject]@{ RelativePath = 'report.docx';  Name = 'report.docx' }
    [pscustomobject]@{ RelativePath = 'kept.txt';     Name = 'kept.txt' }
) | Export-Csv -LiteralPath $manifestFile -NoTypeInformation -Encoding utf8

$lookup = Get-DriveManifestLookup -ManifestPath $manifestFile
Assert-That -Name 'manifest lookup loads paths' -Actual $lookup.Paths.Count -Expected 2

$binItems = @(
    New-BinItem -Name 'report-DESKTOP-4F2K9L1.docx'   # original is in the drive -> debris
    New-BinItem -Name 'kept.txt'                      # already back in the drive
    New-BinItem -Name 'gone-forever.xlsx' -Size 5000  # genuinely missing -> recoverable
    New-BinItem -Name 'orphan (1).pdf'                # copy whose original is nowhere
)

$binRows = ConvertTo-RecycleBinReportRow -Items $binItems -Lookup $lookup
function Get-BinClass { param($Name) ($binRows | Where-Object { $_.Name -eq $Name } | Select-Object -First 1).Classification }

Assert-That -Name 'deleted conflict copy is debris when the original is back' `
    -Actual (Get-BinClass 'report-DESKTOP-4F2K9L1.docx') -Expected 'ConflictCopyDeleted'
Assert-That -Name 'item already back in the drive is not flagged for recovery' `
    -Actual (Get-BinClass 'kept.txt') -Expected 'AlreadyBackInDrive'
Assert-That -Name 'item missing from the drive is a recoverable candidate' `
    -Actual (Get-BinClass 'gone-forever.xlsx') -Expected 'RecoverableCandidate'
Assert-That -Name 'deleted copy whose original is nowhere is surfaced, not written off' `
    -Actual (Get-BinClass 'orphan (1).pdf') -Expected 'RecoverableCandidate'
Assert-That -Name 'classification records that a manifest was used' `
    -Actual ($binRows[0].ClassificationBasis) -Expected 'ManifestChecked'

# With no manifest there is nothing to check against, and the report says so.
$emptyLookup = Get-DriveManifestLookup -ManifestPath ''
$patternRows = ConvertTo-RecycleBinReportRow -Items $binItems -Lookup $emptyLookup
function Get-PatternClass { param($Name) ($patternRows | Where-Object { $_.Name -eq $Name } | Select-Object -First 1).Classification }

Assert-That -Name 'without a manifest a high-confidence copy is still debris' `
    -Actual (Get-PatternClass 'orphan (1).pdf') -Expected 'ConflictCopyDeleted'
Assert-That -Name 'without a manifest an ordinary file is recoverable' `
    -Actual (Get-PatternClass 'kept.txt') -Expected 'RecoverableCandidate'
Assert-That -Name 'classification basis is downgraded without a manifest' `
    -Actual ($patternRows[0].ClassificationBasis) -Expected 'PatternOnly'

Write-Host ''
Write-Host 'Recycle bin download - path mapping and selection' -ForegroundColor Cyan

Assert-That -Name 'DirName strips the personal site and library prefix' `
    -Actual (ConvertTo-DriveRelativePathFromDirName -DirName 'personal/tom_contoso_com/Documents/Projects' -LeafName 'plan.docx') `
    -Expected 'Projects/plan.docx'
Assert-That -Name 'a root-level DirName maps to a bare filename' `
    -Actual (ConvertTo-DriveRelativePathFromDirName -DirName 'personal/tom_contoso_com/Documents' -LeafName 'plan.docx') `
    -Expected 'plan.docx'
Assert-That -Name 'DirName without a library segment drops the personal prefix' `
    -Actual (ConvertTo-DriveRelativePathFromDirName -DirName 'personal/tom_contoso_com/Sub' -LeafName 'a.txt') `
    -Expected 'Sub/a.txt'

function New-SpBinItem {
    param($Name, $Dir = 'personal/tom_contoso_com/Documents', $Type = 1, $Size = 100)
    return [pscustomobject]@{
        Id           = 'sp-' + ($Name -replace '[^A-Za-z0-9]', '')
        Name         = $Name
        Title        = $Name
        DirName      = $Dir
        RelativePath = ConvertTo-DriveRelativePathFromDirName -DirName $Dir -LeafName $Name
        Size         = $Size
        DeletedDate  = '2026-04-02T09:00:00Z'
        DeletedBy    = 'Admin'
        ItemType     = $Type
        ItemState    = 1
    }
}

$spItems = @(
    New-SpBinItem -Name 'deep.txt' -Dir 'personal/tom_contoso_com/Documents/A/B'
    New-SpBinItem -Name 'shallow.txt'
    New-SpBinItem -Name 'A' -Dir 'personal/tom_contoso_com/Documents' -Type 5
    New-SpBinItem -Name 'oldversion.txt' -Type 2   # a file version, not a file
    New-SpBinItem -Name 'kept.txt'
)

$allTargets = Select-RecycleBinDownloadTarget -Items $spItems -Scope All -IncludeFolders
Assert-That -Name 'file versions and list items are not download targets' `
    -Actual (@($allTargets | Where-Object { $_.Name -eq 'oldversion.txt' }).Count) -Expected 0
Assert-That -Name 'folders are restored before the files inside them' `
    -Actual $allTargets[0].Name -Expected 'A'
Assert-That -Name 'shallower files are restored before deeper ones' `
    -Actual (($allTargets | Where-Object { $_.ItemType -eq 1 } | Select-Object -First 1).Name) -Expected 'kept.txt'

$noFolders = Select-RecycleBinDownloadTarget -Items $spItems -Scope All
Assert-That -Name 'folders are excluded unless asked for' `
    -Actual (@($noFolders | Where-Object { $_.ItemType -eq 5 }).Count) -Expected 0

$notInDrive = Select-RecycleBinDownloadTarget -Items $spItems -Scope NotInDrive -Lookup $lookup
Assert-That -Name 'items already in the drive are excluded from NotInDrive scope' `
    -Actual (@($notInDrive | Where-Object { $_.Name -eq 'kept.txt' }).Count) -Expected 0
Assert-That -Name 'items missing from the drive are kept in NotInDrive scope' `
    -Actual (@($notInDrive | Where-Object { $_.Name -eq 'shallow.txt' }).Count) -Expected 1

$patterned = Select-RecycleBinDownloadTarget -Items $spItems -Scope Pattern -Pattern 'deep.*'
Assert-That -Name 'a filename pattern narrows the selection' -Actual $patterned.Count -Expected 1

Write-Host ''
Write-Host 'SharePoint client assertion (real signing and verification)' -ForegroundColor Cyan

# A throwaway self-signed certificate, generated in-process so the JWT signing
# path is exercised for real rather than mocked.
$rsaKey = [System.Security.Cryptography.RSA]::Create(2048)
$certRequest = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    'CN=ToolkitTest', $rsaKey,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
$testCert = $certRequest.CreateSelfSigned([System.DateTimeOffset]::UtcNow.AddDays(-1), [System.DateTimeOffset]::UtcNow.AddYears(1))

$appId = '11111111-2222-3333-4444-555555555555'
$tenantId = 'contoso.onmicrosoft.com'
$jwt = New-ClientAssertionJwt -Certificate $testCert -AppId $appId -TenantId $tenantId

$parts = $jwt -split '\.'
Assert-That -Name 'the assertion has three JWT segments' -Actual $parts.Count -Expected 3

function ConvertFrom-Base64Url {
    param([string]$Text)
    $padded = $Text.Replace('-', '+').Replace('_', '/')
    switch ($padded.Length % 4) { 2 { $padded += '==' } 3 { $padded += '=' } }
    return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($padded))
}

$header = ConvertFrom-Base64Url -Text $parts[0] | ConvertFrom-Json
$claims = ConvertFrom-Base64Url -Text $parts[1] | ConvertFrom-Json

Assert-That -Name 'the assertion is signed RS256'      -Actual $header.alg -Expected 'RS256'
Assert-That -Name 'the header carries the cert x5t'    -Actual $header.x5t -Expected (ConvertTo-Base64Url -Bytes $testCert.GetCertHash())
Assert-That -Name 'issuer is the application id'       -Actual $claims.iss -Expected $appId
Assert-That -Name 'subject is the application id'      -Actual $claims.sub -Expected $appId
Assert-That -Name 'audience is the tenant token endpoint' `
    -Actual $claims.aud -Expected ('https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $tenantId)
Assert-That -Name 'the assertion has not already expired' -Actual ($claims.exp -gt [System.DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) -Expected $true
Assert-That -Name 'nbf precedes exp' -Actual ($claims.nbf -lt $claims.exp) -Expected $true

# The real proof: Entra will verify this signature with the certificate's public key.
$padding = [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
$publicKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($testCert)
$signedBytes = [System.Text.Encoding]::UTF8.GetBytes(('{0}.{1}' -f $parts[0], $parts[1]))
$sigPadded = $parts[2].Replace('-', '+').Replace('_', '/')
switch ($sigPadded.Length % 4) { 2 { $sigPadded += '==' } 3 { $sigPadded += '=' } }
$signatureBytes = [Convert]::FromBase64String($sigPadded)

Assert-That -Name 'the signature verifies against the certificate public key' `
    -Actual ($publicKey.VerifyData($signedBytes, $signatureBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, $padding)) `
    -Expected $true

$tampered = [System.Text.Encoding]::UTF8.GetBytes(('{0}.{1}x' -f $parts[0], $parts[1]))
Assert-That -Name 'a tampered assertion fails verification' `
    -Actual ($publicKey.VerifyData($tampered, $signatureBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, $padding)) `
    -Expected $false

Assert-That -Name 'base64url output carries no padding or unsafe characters' `
    -Actual ((ConvertTo-Base64Url -Bytes ([byte[]](1, 2, 3, 4, 5))) -match '^[A-Za-z0-9_-]+$') -Expected $true

$publicKey.Dispose()
$rsaKey.Dispose()

Write-Host ''
Write-Host 'Config round-trip' -ForegroundColor Cyan

Set-ToolkitConfigValue -Name 'TargetUserId' -Value 'user@contoso.com' | Out-Null
$reloaded = Get-ToolkitConfig -Refresh
Assert-That -Name 'config value persists'    -Actual ([string]$reloaded['TargetUserId']) -Expected 'user@contoso.com'
Assert-That -Name 'config default is used'   -Actual (Get-ToolkitConfigValue -Name 'AppId' -Default 'unset') -Expected 'unset'
Assert-That -Name 'config file exists'       -Actual (Test-Path -LiteralPath $global:OneDriveRepairContext.ConfigPath) -Expected $true

Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ('{0} passed, {1} failed' -f $script:Passed, $script:Failed) -ForegroundColor $(if ($script:Failed -gt 0) { 'Red' } else { 'Green' })
Write-Host ''

if ($script:Failed -gt 0) { exit 1 }
exit 0
