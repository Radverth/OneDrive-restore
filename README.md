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
| — | Download the recycle bin's contents | Menu option 7 (optional, any time) |
| — | Download the sync error files to a USB drive, then delete them | Menu option 8 (optional, any time) |

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
- **Only if you want to download recycle bin contents (option 7):** the
  `Sites.FullControl.All` application permission on Office 365 SharePoint Online,
  which stage 1 offers as an opt-in prompt, and the toolkit certificate's private
  key present on the machine running it.
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

Every menu entry says in plain English what it does and, more importantly, what
it changes. Each option carries one of these tags:

| Tag | Meaning |
|---|---|
| `[reads only]` | Nothing is changed anywhere |
| `[writes to this PC]` | Files are written locally; OneDrive is untouched |
| `[CHANGES ONEDRIVE]` | Will modify the user's OneDrive |
| `[this PC + CAN CHANGE ONEDRIVE]` | Writes locally, and can change OneDrive if you confirm |

Options whose prerequisites are not met say so instead of failing later —
"Needs option 2 first — no downloaded copy on record". Every stage then opens
with a **WHAT THIS DOES** block restating its effect before it asks for anything.

The menu also remembers what it already knows — tenant, app, certificate, target
user, restore point, and the last-used folders — and shows it below the options.
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
- `filename - Copy.ext`, `filename-Copy.ext`, `filename - Copy (2).ext`,
  `filename - Copy 2.ext`
- `filename (user's conflicted copy 2026-04-02).ext`
- `filename-PCNAME.ext` and `filename-user-PCNAME.ext`

Copy names must carry the **dash** (or an underscore). The space-only form,
`filename copy.ext`, is deliberately not matched: the dash is what makes the name
unambiguously a copy, and without it ordinary documents get caught by mistake —
`Certified copy.pdf` is a real filename, not debris.

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

This stage is read-only and needs no permission beyond the `Sites.ReadWrite.All`
stage 1 already grants. Listing the bin is a **beta** endpoint
(`GET /beta/sites/{siteId}/recycleBin/items`). To download the files themselves,
use option 7.

Report: `reports/recycle-bin-inventory-<timestamp>.csv`.

### 7. Download recycle bin contents (optional)

Downloads the actual files out of the recycle bin, to a local folder, with the
original folder structure and timestamps.

**This route is indirect, and it has to be.** A Graph `recycleBinItem` carries
only `id`, `name`, `size`, `deletedDateTime` and `deletedFromLocation` — there is
no content stream and no `@microsoft.graph.downloadUrl` at any API version, so
recycle bin content cannot be read directly. The only way to reach the bytes is
to restore the item first, and
[`driveItem: restore`](https://learn.microsoft.com/en-us/graph/api/driveitem-restore?view=graph-rest-1.0)
is documented as OneDrive **Personal** only. For OneDrive for Business, restore
means the SharePoint REST endpoint
[`POST {siteUrl}/_api/web/RecycleBin('{id}')/restore()`](https://learn.microsoft.com/en-us/answers/questions/355675/sharepoint-rest-api-for-recycle-bin).

So each file goes through three steps:

1. **Restore** it through SharePoint REST.
2. **Download** it through Graph as an ordinary driveItem.
3. **Delete it again**, so it returns to the recycle bin and the drive is left as
   it was found.

That happens **one file at a time**. On an account that is already over quota,
restoring the whole bin at once could push it further over; doing it per file
holds at most one restored file in the drive at any moment, and the plan screen
shows that peak (the largest single file) before you confirm. Answering "no" to
the put-back question leaves everything restored instead — the plan screen warns
when that exceeds the remaining quota.

You can download everything, only items that are not in the drive now (using the
stage 2 manifest), or only names matching a pattern. Folders are restored before
the files inside them, and a file already brought back by its parent folder is
detected and downloaded without a second restore.

Two extra requirements, both called out at runtime if missing:

- **A SharePoint application permission** (`Sites.FullControl.All` on Office 365
  SharePoint Online), which stage 1 offers as an opt-in prompt. Re-run stage 1 on
  an existing registration and it will add the permission without disturbing
  anything else. Everything else in the toolkit works without it.
- **The certificate's private key on this machine.** SharePoint REST rejects
  client-secret app-only tokens, so the toolkit signs a JWT client assertion with
  the same certificate stage 1 created, and exchanges it for a
  SharePoint-audience token.

Reports: `reports/recycle-bin-download-<timestamp>.csv`, plus a copy inside the
download folder. Every row records `RestoreStatus`, `DownloadStatus` and
`PutBackStatus`, so a file that was downloaded but could not be deleted again is
visible rather than silently left consuming quota.

> Two caveats worth knowing. Files put back start a **fresh retention window** and
> reappear as new bin entries with new IDs and a new deleted date. And a restore
> fails if something already occupies the original path — that is reported per
> item and the run continues.

### 8. Download the sync error files, then delete them (optional)

Downloads every sync error file the scanner flagged — the duplicate and conflict
copies the broken sync created — to a local folder (a USB drive is the point),
verifies each one landed intact, writes a CSV manifest, and only then offers to
delete the verified ones from OneDrive.

This is the answer to "I want something I can restore from if this goes wrong."
Deleting straight from the drive leaves the OneDrive recycle bin as the only way
back — on a retention clock, in the tenant that just had the problem. An offline
archive has neither of those constraints.

Reached two ways:

- **From option 3**, as the recommended next step once the scan finishes.
- **As option 8**, against the last scan report or any previous
  `duplicate-scan-<timestamp>.csv` — so you can scan now and archive later.

You choose what to download: everything flagged (the default — the archive is the
safety net, so breadth is the point), or only the hash-verified exact duplicates.

**Trial runs.** Before it starts, the stage offers to process just the first one
or two files so you can check the result and then re-run for the rest — nothing
is skipped permanently. Option 7 offers the same, since restoring is the riskiest
part of that flow. Both also take `-BatchLimit <n>` for scripted use.

**The rule that makes deletion safe:** a copy is only ever eligible for deletion
when it was **downloaded AND verified**. Verification is SHA256 against the
drive's own hash where OneDrive exposes one, and a size match otherwise; the
manifest records which was used per file. Anything that failed to transfer, came
down the wrong size, or mismatched its hash is silently excluded from the
deletion list — a failed backup can never result in the cloud copy being removed.
Deletion is further restricted to `ExactDuplicate` rows and still needs a
confirm plus typing `DELETE`, so a verified `ContentConflict` or `OrphanedCopy`
is archived but never auto-deleted.

**Before the delete prompt** the toolkit writes `pending-delete-<timestamp>.csv`
listing exactly the files it would remove — one row per file, naming the copy and
the original it is a copy of, its size, how it was verified, and where the backup
landed. The prompt points at that path and waits. Whatever is in that file is
what gets deleted; nothing else is touched. The full archive manifest is written
before the prompt too, so abandoning the run still leaves a complete record.

**How the originals are kept out.** The scanner groups each conflict copy around
the file it came from, and the canonical file — the original — is held separately
from the list of copies and is never emitted as a row. Everything downstream
filters those rows, so an original cannot reach the archive, the pending-delete
list, or the deletion. Where no original survives, the *oldest* member of the
group becomes canonical and is excluded the same way, so every group always keeps
at least one file. Deletion is then further narrowed to `ExactDuplicate` rows,
which require a hash match against that canonical.

The archive keeps each file at its **original relative path**, so restoring means
either copying files back by hand, or pointing stage 4 at the archive folder as
the local backup root and letting stage 5 upload. The manifest is written twice —
into `reports/`, and as `_archive-manifest.csv` inside the archive folder, so the
USB stick explains itself. Every row carries `DownloadStatus`, `VerifyStatus`,
`Sha256`, `DeleteStatus`, and the `CanonicalPath` each copy was a copy of.

## Safety rules baked into the design

- **No blind re-upload of the backup.** Only files proven newer than the restore
  point and absent from the cloud copy are uploaded.
- **Same-name-different-content always goes to a human.** Guessing loses data on
  one side or the other, silently.
- **Deletion is never automatic**, never inferred from a filename alone, and
  never applied to anything that was not hash-verified as an exact duplicate.
- **A copy can only be deleted once a verified local backup of it exists** (when
  deleting through option 8). A failed or truncated download excludes that file
  from deletion entirely.
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
│   ├── DuplicateScanner.psm1       # stage 3 + duplicate archive (option 8)
│   ├── CompareDrives.psm1          # stage 4
│   ├── ReconcileUpload.psm1        # stage 5
│   └── RecycleBin.psm1             # recycle bin inventory + download (options 6-7)
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

126 assertions covering the conflict-pattern matching, duplicate grouping and
classification, the archive-then-delete gate, recycle bin classification and
restore ordering, path safety, config round-tripping, and a full stage 4
comparison run against throwaway folders on disk. The SharePoint JWT client
assertion is signed with a generated certificate and its signature verified
against the public key — the same check Entra performs — including a negative
case proving a tampered assertion fails. No tenant, no Graph modules,
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

**"Could not authenticate to SharePoint" on option 7** — the app registration
does not carry the SharePoint permission yet. Re-run stage 1 and answer yes to
the SharePoint prompt; it adds the permission to the existing registration
without disturbing the certificate or the Graph permissions. If the certificate
itself cannot be found, option 7 must run on the machine that holds its private
key.

**A restore fails with a conflict** — something already occupies the file's
original path, so SharePoint will not restore over it. The item is reported with
that error and the run continues; deal with those individually.

**The certificate expires** — `Add-ToolkitAppCertificate` adds a new certificate
to the existing registration without disturbing the old one; update
`CertificateThumbprint` in the config afterwards.
