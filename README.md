# OneDrive Sync Repair & Reconciliation Toolkit

A PowerShell 7 console toolkit for repairing a OneDrive account whose sync broke
and filled the drive with conflict copies until it blew past its storage quota.

The repair has five steps. The first is manual; the rest are what this toolkit
automates:

| # | Step | Where |
|---|------|-------|
| 0 | Roll the drive back to a clean point in time | **Manual**, in the SharePoint/OneDrive admin centre |
| 1 | Register an app + certificate for Graph access | Menu option 1 (one-time) |
| 2 | Download a fresh cloud copy of the rolled-back drive | Menu option 2 |
| 3 | Scan for duplicate/conflict files, optionally clean up | Menu option 3 |
| 4 | Compare that download against the local PC backup | Menu option 4 |
| 5 | Upload only the legitimate recovered work | Menu option 5 |
| — | Inventory the recycle bin to see what was deleted | Menu option 6 (optional, any time) |

The rollback in step 0 also wipes any genuine files and edits made between the
restore point and now. Steps 2-5 exist to find that work in a local backup of
the user's PC and put it back — exactly once each, without recreating the
duplicate mess.

## Why download the drive instead of reading a sync folder

The affected user has sync problems on one device but not another, so neither
device's local OneDrive folder is a trustworthy picture of the cloud. Stage 2
pulls the drive straight down through Graph. That copy — not a sync folder — is
the cloud-side half of the stage 4 comparison, so nothing that broke the bad
device's sync gets dragged back into the drive.

## Prerequisites

- **PowerShell 7+** (`pwsh`). Stage 1 must run on Windows: it uses
  `New-SelfSignedCertificate` and the Windows certificate store.
- **Microsoft Graph PowerShell SDK.** The minimum is
  `Install-Module Microsoft.Graph.Authentication`, plus
  `Microsoft.Graph.Applications` for stage 1. `Install-Module Microsoft.Graph`
  installs everything.
- **Global Administrator** for the one-time registration and consent — or
  Application Administrator + Cloud Application Administrator + Privileged Role
  Administrator between them.
- **The admin-centre rollback already done**, with the restore point date/time
  written down. Stage 4 asks for it and uses it to decide what counts as
  recovered work.
- **A reliable local backup** of the user's PC files from after the last good
  state, with timestamps intact.
- **Enough free disk space** to hold a full copy of the user's OneDrive. Stage 2
  checks this against the drive's reported usage before it starts.

## Quick start

```powershell
cd OneDriveRepairToolkit
./Start-OneDriveRepair.ps1
```

The menu remembers what it already knows — tenant, app, certificate, target
user, restore point, and the last-used folders — and shows it above the prompt.
Any single stage can also be run directly, for repeat runs or scripting:

```powershell
./Start-OneDriveRepair.ps1 -Stage 2
```

## The stages in detail

### 1. Register Azure AD app & generate certificate (one-time)

Signs in interactively as an admin, creates a self-signed certificate (1-2 year
expiry, private key in `Cert:\CurrentUser\My`, public `.cer` exported to
`config/`), registers the application with the certificate attached, creates its
service principal, and grants admin consent for the application permissions
`Files.ReadWrite.All`, `Sites.ReadWrite.All` and `User.Read.All`.

If consent is blocked — conditional access, or a missing privileged role — the
run says so and prints the portal consent URL rather than pretending it worked.

Tenant ID, app ID and certificate thumbprint are saved to
`config/toolkit-config.json`. Every later stage authenticates with:

```powershell
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint
```

> The certificate is attached during `New-MgApplication` rather than through
> `New-MgApplicationKeyCredential`. That cmdlet calls Graph's `addKey` action,
> which requires proof-of-possession signed with a credential the application
> does not yet have. `New-MgApplicationKeyCredential` is the right tool for
> *rollover*, which is what `Add-ToolkitAppCertificate` uses it for.

### 2. Download cloud copy of the user's OneDrive

Resolves the user's drive, asks where to put the copy (never hardcoded, offers
to create the folder, checks free space against the drive's usage, and remembers
the choice for next time), then walks the whole drive breadth-first following
`@odata.nextLink` at every folder level.

The folder structure is recreated locally, including empty folders. Each file is
fetched from its `@microsoft.graph.downloadUrl` and stamped with its original
`createdDateTime` and `lastModifiedDateTime`. Throttling (`429`) is retried with
`Retry-After` honoured; other transient failures back off exponentially.

Output is a manifest — path, size, SHA256, both timestamps, item ID — written to
`reports/cloud-manifest-<timestamp>.csv` and `.json`, with a copy inside the
snapshot folder so it stays self-describing if moved.

Each run lands in its own timestamped subfolder, so re-running for a fresh
snapshot (say after cleaning up in stage 3) never clobbers an earlier one.
Re-running against the *same* folder resumes: files already present with a
matching size and timestamp are skipped.

### 3. Duplicate/conflict file scanner

Scans either the live drive or a stage 2 manifest (faster, and gives SHA256
comparisons rather than Graph's quickXorHash). It matches the usual conflict
naming shapes:

- `filename (1).ext`, `filename (2).ext` …
- `filename-Copy.ext`, `filename - Copy.ext`
- `filename (user's conflicted copy 2026-04-02).ext`
- `filename-PCNAME.ext` and `filename-user-PCNAME.ext`

Candidates are grouped around the file they were copied from and classified:

| Classification | Meaning | Action |
|---|---|---|
| `ExactDuplicate` | Hash-identical to the canonical file | Safe to delete |
| `ProbableDuplicate` | Same size, but no hash available on either side | Review |
| `ContentConflict` | Same base name, different content | Review — a human decides |
| `OrphanedCopy` | Conflict-named, but its original no longer exists anywhere | Review |

Reports go to `reports/`: the full scan, a safe-to-delete list, and a
needs-review list. Deletion is offered only for hash-verified exact duplicates,
only after the list is shown, and only after typing `DELETE` in capitals.
Deleted items land in the user's OneDrive recycle bin.

### 4. Compare downloaded copy vs. local PC backup

Asks for the restore point (saved to config), indexes both trees by normalised
relative path, and compares content with `Get-FileHash -Algorithm SHA256` — not
just size and date. Sizes are checked first, so identical-size files are the
only ones that pay the hashing cost.

| Category | Meaning | Action |
|---|---|---|
| `UploadCandidate` | Newer than the restore point, missing from the cloud copy | Upload — this is the recovered work |
| `ContentMismatch` | In both, different content | Review |
| `DuplicatePatternSkip` | Conflict-copy noise whose original exists | Skip |
| `DuplicatePatternNoOriginal` | Conflict-named, but no original anywhere | Review |
| `Identical` | Same SHA256 on both sides | None |
| `LocalOnlyBeforeRestorePoint` | Missing from the cloud, untouched since the restore point | Review |
| `CloudOnly` | In the cloud copy but not the backup | None (informational) |

Everything lands in `reports/comparison-<timestamp>.csv`, with a per-category
count on the console.

### 5. Reconcile — upload canonical files

Reads the stage 4 report and uploads **only** what was categorised
`UploadCandidate`, plus any `ContentMismatch` the admin resolves in favour of the
local copy. Files under 4 MB go up with a single content `PUT`; anything larger
uses a resumable upload session in 10 MB chunks with a progress bar. Missing
parent folders are created first.

Every upload uses `@microsoft.graph.conflictBehavior=replace`. `rename` would
create exactly the `file (1).docx` copies this whole exercise is undoing.

After upload, `fileSystemInfo.lastModifiedDateTime` is patched back to the
local file's timestamp. `createdDateTime` is sent too, but OneDrive commonly
keeps the upload time for newly created items — a platform behaviour, not
something a client can force.

There is a dry run: answer "no" to *Upload for real?* to get the full plan
written to `reports/reconcile-dryrun-<timestamp>.csv` with nothing changed.

### 6. Inventory the OneDrive recycle bin (optional)

Lists everything in the user's first-stage recycle bin and splits it into what
matters and what does not, cross-referencing the stage 2 manifest so "is this
already back in the drive?" is answered from data rather than guessed:

| Classification | Meaning |
|---|---|
| `RecoverableCandidate` | Not in the current drive — restore by hand if still wanted |
| `PossiblyBackInDrive` | Same filename exists elsewhere in the drive — check first |
| `AlreadyBackInDrive` | A file with this path is in the drive now — nothing to do |
| `ConflictCopyDeleted` | Sync-conflict copy; deleting it was the point |

A conflict-named item is only written off as debris when the file it was copied
from is actually in the drive now — the same rule stages 3 and 4 use, so the
three cannot disagree about what counts as noise. Without a stage 2 manifest
there is nothing to check against, so classification falls back to the filename
alone and every row records that in `ClassificationBasis`.

Report: `reports/recycle-bin-inventory-<timestamp>.csv`.

> **What this stage cannot do, and why.** Recycle bin *contents* cannot be
> downloaded. A Graph `recycleBinItem` carries only `id`, `name`, `size`,
> `deletedDateTime` and `deletedFromLocation` — there is no content stream and no
> `@microsoft.graph.downloadUrl` at any API version. Listing the bin is itself
> **beta-only** (`GET /beta/sites/{siteId}/recycleBin/items`), though it needs no
> permission beyond the `Sites.ReadWrite.All` stage 1 already grants.
>
> The only route to the bytes is to restore an item first, and
> [`driveItem: restore`](https://learn.microsoft.com/en-us/graph/api/driveitem-restore?view=graph-rest-1.0)
> is documented as OneDrive **Personal** only. For OneDrive for Business, restore
> means the SharePoint REST endpoint `POST {siteUrl}/_api/web/recyclebin('{id}')/restore()`,
> which needs a SharePoint-audience token and a SharePoint application permission
> this toolkit does not request. So this stage reports; restoring is done in the
> web UI, after which stage 2 will pull the restored files down with everything
> else.

## Safety rules baked into the design

- **No blind re-upload of the backup.** Only files proven newer than the restore
  point and absent from the cloud copy are uploaded.
- **Same-name-different-content always goes to a human.** Guessing loses data on
  one side or the other, silently.
- **Deletion is never automatic**, never inferred from a filename alone, and
  never applied to anything that was not hash-verified as an exact duplicate.
- **A conflict-named file whose original exists nowhere is never discarded.**
  `report (1).docx` with no `report.docx` on either side may be the only copy of
  real work, so it is surfaced for review instead of being written off as noise.
- **Hyphenated filenames are not treated as conflicts on their own.** The
  `name-MACHINE` shape also matches ordinary names like `annual-report.docx`, so
  it only counts as a conflict copy when a file with the base name actually
  exists beside it.
- **Equal file sizes are never treated as proof of identical content.** Without a
  hash on both sides, a match is reported for review, not for deletion.

## Layout

```
OneDriveRepairToolkit/
├── Start-OneDriveRepair.ps1        # menu entry point / thin dispatcher
├── modules/
│   ├── Common.psm1                 # config, logging, Graph auth + retry, shared patterns
│   ├── AppRegistration.psm1        # stage 1
│   ├── DownloadOneDrive.psm1       # stage 2
│   ├── DuplicateScanner.psm1       # stage 3
│   ├── CompareDrives.psm1          # stage 4
│   ├── ReconcileUpload.psm1        # stage 5
│   └── RecycleBin.psm1             # recycle bin inventory (menu option 6)
├── config/
│   ├── toolkit-config.json         # written at runtime (gitignored)
│   └── toolkit-config.example.json
├── logs/toolkit-YYYYMMDD.log       # rolling daily log, pruned after 30 days
├── reports/                        # every CSV/JSON the stages produce
└── tests/Invoke-ToolkitTests.ps1   # offline tests, no tenant needed
```

`Common.psm1` is not in the original design sketch but earns its place: the
config, logging, throttling-aware Graph wrapper, drive enumeration and the
conflict-name patterns are all needed by more than one stage, and the conflict
patterns in particular *must* be identical in stages 3 and 4 or the two would
disagree about what counts as duplicate noise.

## Tests

```bash
pwsh ./OneDriveRepairToolkit/tests/Invoke-ToolkitTests.ps1
```

62 assertions covering the conflict-pattern matching, duplicate grouping and
classification, recycle bin classification, path safety, config round-tripping,
and a full stage 4 comparison run against throwaway folders on disk. No tenant, no Graph modules,
no network — it runs anywhere PowerShell 7 does.

## Troubleshooting

**"Missing configuration: TenantId, AppId…"** — stage 1 has not been run, or
`config/toolkit-config.json` was moved. Run menu option 1.

**Graph connection fails right after stage 1** — consent and directory
replication can take a few minutes. Stage 1 waits 20 seconds and verifies; if it
reports a failure, wait and try stage 2.

**Consent was blocked** — grant it in the portal with the URL stage 1 prints,
then re-run stage 2.

**Downloads or uploads are slow and log `HTTP 429`** — that is Graph throttling.
The toolkit honours `Retry-After` and backs off; let it run.

**Stage 3 reports many `ProbableDuplicate` rows** — Graph did not return hashes
for those files. Run stage 2 first, then re-scan using the manifest to get
SHA256-backed comparisons.

**The recycle bin inventory fails or returns nothing** — the listing endpoint is
beta-only and is not enabled in every tenant. The bin is always readable in the
OneDrive web UI as a fallback.

**The certificate expires** — `Add-ToolkitAppCertificate` adds a new certificate
to the existing registration without disturbing the old one; update
`CertificateThumbprint` in the config afterwards.
