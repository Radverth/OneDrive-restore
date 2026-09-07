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
foreach ($module in @('Common', 'DownloadOneDrive', 'DuplicateScanner', 'CompareDrives', 'ReconcileUpload')) {
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
