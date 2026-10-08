**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **Reporting**

# Reporting Scripts

Scripts that generate reports on Active Directory, SharePoint Online and licensing.

---

## Folders

| Folder | Description |
|--------|-------------|
| [`Licensing/`](Licensing/readme.md) | Monthly licensing and Azure cost report from Pax8 and Ingram data |

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-ComputerLastLogon.ps1`](Get-ComputerLastLogon.ps1) ([docs](#get-computerlastlogonps1)) | Last logon date of computer objects in one or more OUs, exported to CSV |
| [`Get-SharePointStorageReport.ps1`](Get-SharePointStorageReport.ps1) ([docs](#get-sharepointstoragereportps1)) | Tenant-wide storage report: sites, libraries, version history, recycle bin and long paths |
| [`Get-SharePointPermissionsReport.ps1`](Get-SharePointPermissionsReport.ps1) ([docs](#get-sharepointpermissionsreportps1)) | Who has access where, through which group and at what level — every site, list, folder and file with its own permissions. Read-only, to CSV and a single Excel workbook |
| [`Remove-SharePointFileVersionsByDate.ps1`](Remove-SharePointFileVersionsByDate.ps1) ([docs](#remove-sharepointfileversionsbydateps1)) | Deletes file versions older than a date; the current version is always kept. Reports only by default |

---

## Get-ComputerLastLogon.ps1

Reports the **last logon date** of computer objects in one or more OUs, exported to CSV.

### How it works

Two accuracy modes:

| Mode | Attribute | Lag | Speed |
|---|---|---|---|
| Default | `LastLogonTimestamp` (replicated) | up to 14 days | Fast |
| `-AllDCs` | `LastLogon` per DC, best value | None | Slower |

Use the default mode for stale-device reporting. Use `-AllDCs` when absolute accuracy is required.

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-SearchBase` | `string[]` | _(whole domain)_ | One or more OU distinguished names |
| `-AllDCs` | switch | off | Queries every DC for the most accurate `LastLogon` |
| `-InactiveDays` | int | `90` | Threshold in days after which a computer counts as _Stale_ |
| `-IncludeDisabled` | switch | off | Also includes disabled computer objects |
| `-ExportPath` | string | `C:\Temp\` | Folder for the CSV file |

### Requirements

- ActiveDirectory PowerShell module (RSAT)
- Read access to the given OUs

### Examples

```powershell
# Laptops and Computers OU
.\Get-ComputerLastLogon.ps1 `
    -SearchBase "OU=Laptops,OU=Computers,DC=bedrijf,DC=local",
               "OU=Computers,DC=bedrijf,DC=local"

# Most accurate mode — queries every DC
.\Get-ComputerLastLogon.ps1 -SearchBase "OU=Computers,DC=bedrijf,DC=local" -AllDCs

# Include disabled computers, threshold at 60 days
.\Get-ComputerLastLogon.ps1 -SearchBase "OU=Computers,DC=bedrijf,DC=local" `
    -IncludeDisabled -InactiveDays 60
```

### Status values

| Status | Meaning |
|---|---|
| `Active` | LastLogon within the `-InactiveDays` threshold |
| `Active (pwd recent)` | LastLogon looks stale because of replication lag, but the computer account password was renewed < 35 days ago — the device is online |
| `Stale` | Both LastLogon and PasswordLastSet exceed the threshold — most likely genuinely inactive |
| `Never` | Never logged on and no recent password |
| `Disabled` | Account disabled in AD |

> **Tip:** `Active (pwd recent)` are PCs that *are* active but wrongly show up as stale because of the 9-14 day replication lag of `LastLogonTimestamp`. Use `-AllDCs` for exact data when this distinction matters.

### CSV columns

| Column | Description |
|---|---|
| `Name` | Computer name |
| `Status` | See status values above |
| `Enabled` | True/False |
| `LastLogon` | Last logon date (dd/MM/yyyy HH:mm) |
| `DaysSinceLogon` | Number of days ago |
| `PasswordLastSet` | Date of the last computer account password change (dd/MM/yyyy) |
| `DaysSincePasswordSet` | Days since the password was renewed |
| `OperatingSystem` | OS name |
| `OperatingSystemVersion` | OS version |
| `IPv4Address` | IP address (if available) |
| `OU` | OU path (readable format) |
| `Created` | Creation date in AD |
| `Description` | Description from AD |
| `DistinguishedName` | Full AD path |

---

## Licensing/

See [Licensing/](Licensing/) for the monthly licensing report.

---

## Get-SharePointStorageReport.ps1

Reports storage usage across SharePoint Online with a tenant-wide scan. By default the script signs in delegated through [`Connect-M365.ps1`](../Startup/readme.md#connect-m365ps1) (device code and GDAP customer per `load.config.ps1`) and temporarily creates an App Registration (`Sites.Read.All`) for site enumeration; that app is deleted again afterwards. `-AppOnly` uses the app for the tenant in `graph.appid.json` instead. Everything is read through Microsoft Graph — sites, libraries (hidden ones such as the Preservation Hold Library included, found through Graph `/lists` with the `system`/`hidden` facets), files and versions. The file walk used to have a SharePoint REST variant for hidden libraries, fed by a client-secret token that SharePoint Online always rejects; it is gone, and a hidden library Graph will not open is now reported instead of silently skipped. Requires PowerShell 7.


### Coverage

- All SharePoint site collections (OneDrive personal sites are excluded)
- Subsites at every level
- Teams-related SharePoint locations:
    - Standard channels as libraries/folders in the parent Teams site
    - Private/shared channels as separate site collections
- Folders and files inside document libraries (only with `-Apply`)
- The detail output contains both folders and files (`ItemType`), so you see the full structure
- In `-Apply` mode everything goes into 1 ranked CSV (largest folders + files, including version history)
- The CSV also has `Level` (depth): root = `0`, top-level folder = `1`, etc.
- The CSV also has `ParentPath` for hierarchical analysis (building a tree in Excel/Power BI)

### Recycle bin

The recycle bin (stage 1 + stage 2) counts towards the tenant storage quota, so it is collected **separately** from the library scan, and only for real SharePoint site collections (not OneDrive). Graph has no recycle bin API for SharePoint sites, so this part still uses SharePoint REST with a token minted from the temporary app's client secret (or `-ClientSecret`) — and SharePoint Online rejects secret-based app-only tokens, so expect empty recycle bin figures until that moves to a certificate (not verified on a tenant):

- By default (`-Apply`, as Phase 2b) or on its own with **`-RecycleBinOnly`** (skips the library scan entirely, recycle bin only)
- Only root site collections have a recycle bin of their own (sub-webs share the root's)
- Output: an extra row per site in the summary CSV (`Library = "Recycle Bin (stage 1 + 2)"`) plus a row per deleted item in the detail CSV

```powershell
# Recycle bin only
.\Get-SharePointStorageReport.ps1 -RecycleBinOnly

# Full scan + recycle bin as an extra phase
.\Get-SharePointStorageReport.ps1 -Apply
```

### Resuming after an interruption (checkpoints) and progress

With `-Apply` (or `-RecycleBinOnly`) a checkpoint is written to the output folder after every completed library (or site recycle bin): `SharePoint_StorageReport_<hash>.state.json` + `.summary.partial.csv` + `.detail.partial.csv` + `.longpaths.partial.csv`. The `<hash>` is derived from all scan parameters (site, mode, output folder, etc.), so:

- **Starting again with the same parameters** resumes automatically from the last completed library — libraries already done are skipped (`[SKIP] Already completed in a previous run.`).
- **`-Restart`** discards an existing checkpoint and starts the scan from scratch, even when the parameters are the same.
- The checkpoint files are cleaned up automatically once the scan completes successfully — if they are still there, the previous run was interrupted.

During a long scan the script shows both scrolling log lines and (in an interactive console) nested progress bars per phase: inventorying sites/libraries, scanning folders/files per library, fetching version history, and recycle bins. When Microsoft Graph throttles (for example `activityLimitReached` during version-history lookups), a `[WAIT] throttled by Microsoft Graph — waiting ...` message appears with the wait time, instead of the script appearing to hang silently.

> **Note (delegated/SDK calls):** by default the Microsoft.Graph SDK cmdlets silently retry 429/503 *themselves*, with their own internal backoff that can honour a hefty `Retry-After` on `activityLimitReached` — that could cause minutes of silence without the script's own `[WAIT]` message ever showing up. The script therefore sets `Set-MgRequestContext -ClientTimeout <-GraphTimeoutSec> -MaxRetry 0` right after connecting, so every Graph SDK call gets a hard timeout and all retries go through the script's own, visible logic.

```powershell
# Automatically resume an interrupted tenant scan
.\Get-SharePointStorageReport.ps1 -Apply

# Ignore the checkpoint and start over from scratch
.\Get-SharePointStorageReport.ps1 -Apply -Restart
```

### Site collection totals (comparing with the admin portal)

Subsites and Teams channels share the storage quota of their root site collection, but the scan reports them as **separate site rows**; the recycle bin is kept in a **row of its own** as well. Taken separately, those figures cannot be compared one-to-one with the single "storage used" number the SharePoint admin portal shows per site collection.

With `-Apply` (Phase 2c) the script therefore adds everything back up per root site collection automatically: the library totals of all underlying subsites/channels + the recycle bin of that site collection. Output: `SharePoint_SiteCollectionTotals_<timestamp>.csv`, with per site collection `LibrariesMB`, `RecycleBinMB`, `GrandTotalMB`/`GrandTotalGB` and the number of subsites/channels that were counted. The console also shows the top 10.

If `GrandTotalGB` for a site still differs from the admin portal figure, the most likely cause is one of:
- **Timing** — the admin portal figure can lag up to 24h behind a live scan
- **Silently skipped folders** — an `[ERROR] Cannot read folder` message in the console means that subfolder tree (permission problem) was not counted
- **Failed version lookups** — fall back to "0 versions" after repeated Graph errors (rare, only after 3 failed retries)
- Compare without `-Apply` first (quick mode) — that uses the same official `quota.used` figure as the admin portal, so if that already differs, the difference is not in the `-Apply` count itself

### Long paths (Windows limits)

A library that is fine in SharePoint can still fail once it is synced with the OneDrive client or opened from Windows: the local path is longer than the SharePoint path, because the profile folder and the sync folder come in front of it. With `-Apply` (also with `-FastMode`) every file and folder is therefore measured twice:

| Path | Example | Limit |
|---|---|---|
| SharePoint (server-relative, decoded) | `/sites/Finance/Shared Documents/<folders>/<file>` | **400** characters — SharePoint refuses anything longer |
| Local OneDrive sync path | `C:\Users\<user>\<Organisation>\<Site> - <Library>\<folders>\<file>` | **260** (Windows `MAX_PATH`, 259 usable) — Explorer, many applications and older tools fail beyond it |
| Same, for Excel workbooks (`.xls*`, `.xlt*`) | | **218** — Excel will not open or save the workbook |
| Same, for PDFs (`.pdf`) | | **255** — Adobe Acrobat/Reader cannot open the file, notably from a synced or network folder |

The local path is an estimate, because its length depends on who syncs it. By default the script measures for the **enabled member account with the longest UPN** in the tenant: the profile folder is named after the UPN prefix (the part before the `@`), so a path that fits for that user fits for everyone. That needs `User.Read.All` (added to the interactive sign-in); when the users cannot be read it falls back to `C:\Users\firstname.lastname`. The organisation name is the tenant's display name from Graph, or the tenant name from the SharePoint URL when that cannot be read. `-SyncProfilePath` and `-OrganizationName` override both. A OneDrive personal site (`-IncludeOneDriveUsers`) is measured as `C:\Users\<user>\OneDrive - <Organisation>\...`.

Output: `SharePoint_LongPaths_<timestamp>.csv`, longest first, with every item whose local path is `-LongPathThreshold` (default `200`) characters or more, or that is over a limit — columns `LocalPathLength`, `SharePointPathLength`, `OverLimit` (`SharePoint (400)`, `Windows (260)`, `Excel (218)`, `Adobe (255)` or empty) and the estimated `LocalPath`. The console shows per library how many long paths it has, and at the end the count per limit and the 10 longest paths; the Markdown report gets the same top 10.

> Windows may shorten the profile folder name (for example to 20 characters for an on-premises account, or with a suffix when the name already exists). Measuring with the full UPN prefix is the cautious side: it can report a path as too long that just fits, never the other way round.

### Performance (version history lookups)

Version history is the most expensive step: by nature 1 Graph call per file. Three optimisations limit that:

- **Skipped when versioning is off** — if versioning is known for certain to be off for a library, no per-file version call is made (0 versions is the answer anyway). When the status is unknown (fallback via `Get-MgSiteDrive`) it is still fetched, to be safe.
- **Batched through Graph's `$batch` endpoint** — version lookups for the files in a library are now fetched in groups of 20 per HTTP call, instead of 1 separate call per file.
- **Shorter retry for these specific calls** — at most 3 attempts with a short backoff (instead of the standard `-MaxGraphRetry`/backoff, which for critical calls can run up to ~2 minutes per attempt). A failed version lookup falls back to "0 versions" instead of holding up the whole scan.

`-SkipVersions` remains the fastest option when version history is not needed — then no version call is made at all.

In addition, `-SiteUrl` (1 specific site) is optimised: in normal mode the script uses delegated Graph calls for that site only, which brings the start-up time in line with other commands.

For GDAP reliability the script automatically switches to an app-only bootstrap for single-site scans when `authMode=GDAP` is detected (from `load.config.ps1`/launcher context). To always force that, use `-ForceAppOnlySingleSite`.

For full-site scans under GDAP the script uses the same customer-tenant context (`$global:cid`/`-TenantId`) for both the delegated sign-in (`Connect-M365Graph`) and the temporary app bootstrap, so consent and site enumeration always happen in the right tenant. A Graph session that already holds the scopes is reused, and then left open at the end.

### Parameters

| Parameter | Description |
|---|---|
| `-SiteUrl` | Scan 1 specific site. A tenant-root URL (e.g. `https://contoso.sharepoint.com`) automatically triggers a tenant-wide scan |
| `-SkipVersions` | Leaves version history out (faster) |
| `-OutputPath` | Overrides the default output folder (`C:\Temp\` / `~/Downloads/`) |
| `-TenantId` | Entra ID tenant ID — detected automatically if not given; required together with `-ClientId` |
| `-ClientId` | Existing App Registration client ID — skips auto-create; use together with `-TenantId` and `-ClientSecret` or `-CertificateThumbprint` |
| `-ClientSecret` | Client secret for an existing app registration |
| `-CertificateThumbprint` | Certificate thumbprint for an existing app registration |
| `-AppOnly` | App-only with ClientId and CertificateThumbprint for the tenant from `graph.appid.json` — no sign-in, no temporary app. Needs `Sites.Read.All` (and `User.Read.All` for the profile path) as application permissions |
| `-Apply` | Full recursive scan of libraries, folders and files. Without this switch, quota summary only |
| `-UseHighPrivilege` | Auto mode: temporarily grants `Sites.FullControl.All` instead of `Sites.Read.All` when read-only permissions turn out to be insufficient |
| `-RecycleBinOnly` | Skips storage/library scanning — reads only recycle bin items (stage 1 + stage 2) per site collection |
| `-ForceAppOnlySingleSite` | Forces the temporary app bootstrap for `-SiteUrl` scans (useful for GDAP/delegated restrictions) |
| `-GraphTimeoutSec` | Timeout in seconds per Graph call (default: `120`) |
| `-MaxGraphRetry` | Maximum number of retries on Graph throttling/timeouts (default: `6`) |
| `-VersionBatchConcurrency` | Number of parallel `$batch` workers for fetching version history, 1-8 (default: `4`) |
| `-MaxVersionRetryPasses` | Maximum number of retry passes for version history under sustained throttling. `0` (default) scales automatically with the number of files — SharePoint enforces a hard activity ceiling of ~1500-2500 resolved version lookups per pass, so on tenants with hundreds of thousands of files a fixed low value (previously hardcoded at 8) gave up early for most of the scan. Set it higher/lower explicitly to override the auto-scaling |
| `-Restart` | Discards an existing checkpoint for this parameter combination and starts the scan from scratch |
| `-SyncProfilePath` | Profile folder used for the local path estimate (default: `C:\Users\<prefix>` of the enabled member account with the longest UPN; fallback `C:\Users\firstname.lastname`) |
| `-OrganizationName` | Organisation name in the OneDrive sync folder (default: the tenant's display name, fallback the tenant name from the URL) |
| `-LongPathThreshold` | Local path length from which an item goes into the long paths CSV, 1-1000 (default: `200`). Anything over a limit is always listed |

### Examples

```powershell
# Quick summary — site quota only, no file scan
.\Get-SharePointStorageReport.ps1

# Full tenant scan including subsites and files (auto app registration)
.\Get-SharePointStorageReport.ps1 -Apply

# Same, but with higher temporary app permissions if needed
.\Get-SharePointStorageReport.ps1 -Apply -UseHighPrivilege

# One specific Teams site only
.\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/teams/Operations" -Apply

# Full scan with an existing app registration
.\Get-SharePointStorageReport.ps1 -Apply -ClientId "..." -TenantId "..." -ClientSecret "..."

# Ignore an interrupted run and start over from scratch
.\Get-SharePointStorageReport.ps1 -Apply -Restart

# Long paths only, fast (no versions, no detail CSV), measured for a specific user
.\Get-SharePointStorageReport.ps1 -Apply -FastMode -SyncProfilePath 'C:\Users\annemarie.vandenberg' -OrganizationName 'Contoso Nederland B.V.'
```

---

## Get-SharePointPermissionsReport.ps1

Reports **who has access to what** in SharePoint Online — every site, subsite, list/library, folder and file that carries its own permissions, exported to CSV. Read-only: the script makes nothing but `GET` calls and never changes a permission.

### Coverage

- Site collection administrators
- Role assignments at web level (site and subsite), including broken inheritance
- SharePoint groups (Owners/Members/Visitors and custom groups) with their full member list
- Role assignments on lists and document libraries
- Folder and file level: every item with `HasUniqueRoleAssignments`
- Sharing links (anonymous / organisation / specific people) and who they were shared with
- External and guest users, plus `Everyone` and `Everyone except external users`
- Entra ID groups, resolved to their transitive member list

Inheritance is followed the way SharePoint models it itself: an item only shows up as a scope of its own when it has its own permissions. Everything else inherits from the nearest parent, which is reported once. That keeps the CSV a map of the permission structure rather than a row per file.

> Sites are deliberately collected twice: first tenant-wide through Graph (`getAllSites`), then queried again for subsites through both Graph and SharePoint REST (`/_api/web/webs`), and de-duplicated on URL. Classic sub-webs that Graph skips are picked up this way after all.

### Authentication

Reading role assignments is **not** possible through Microsoft Graph, and is not covered by SharePoint's Read/Write/Manage roles either: it needs the application role `Sites.FullControl.All`. The script therefore signs you in once — delegated, through [`Connect-M365Graph`](../Startup/readme.md#connect-m365ps1): a device code when `useDeviceCodeAuth` is set in `load.config.ps1`, the GDAP customer from `Connect-Tenant`, and an existing Graph session with the scopes is reused — and then creates a short-lived App Registration of its own with:

| Resource | Role | Used for |
|---|---|---|
| SharePoint | `Sites.FullControl.All` | Role assignments, site groups, item scopes |
| Graph | `Sites.Read.All` | Tenant-wide site enumeration |
| Graph | `GroupMember.Read.All` | Resolving Entra group membership |

That app is deleted again afterwards. Despite the Full Control role, the script never writes anything. If you don't want a temporary app, pass `-ClientId` + `-TenantId` + `-CertificateThumbprint` of an existing registration that already has these roles, or `-AppOnly` to take them from `graph.appid.json` (PowerShell 7). The scan itself stays app-only on purpose: SharePoint REST accepts neither a delegated Graph token (wrong audience) nor a secret-based app-only token.

> **Certificate, not secret — and that is not a preference.** SharePoint Online rejects every app-only token obtained with a client secret: you get `401` with `x-ms-diagnostics: ... Unsupported app only token`. Only certificate-based app-only authentication works against `_api`. The temporary app therefore gets a certificate that the script creates **in memory** and registers on the app; it never goes into the certificate store or onto disk, so there is nothing to clean up afterwards. If you pass `-ClientSecret` with your own app, the script warns you: the Graph half will work, the SharePoint half will not.

#### Why not just Graph?

Graph can do part of it: on a `driveItem`, `/permissions` returns the permissions, the sharing links (with type and expiry date) and, through `inheritedFrom`, whether inheritance was broken. But the rest of the picture is simply missing there — there is no Graph endpoint for:

| What | Graph | SharePoint REST |
|---|---|---|
| Role assignments at site/web level | ❌ does not exist | ✅ `/_api/web/roleassignments` |
| SharePoint groups and their members | ❌ does not exist | ✅ `/_api/web/sitegroups` |
| Site collection administrators | ❌ does not exist | ✅ `/_api/web/siteusers` |
| Name of the permission level (Full Control, Edit, custom levels) | ❌ only `read`/`write`/`owner` | ✅ `RoleDefinitionBindings` |
| Lists without a `driveItem` (ordinary lists) | ❌ | ✅ |
| Cheap filtering on unique permissions | ❌ one call per item | ✅ `HasUniqueRoleAssignments` in one sweep |

That last row is also a speed difference: through Graph you would have to make a `/permissions` call for *every* file, whereas SharePoint tells you in a single pass per list *which* items have their own permissions. A Graph-only variant *could* work with a client secret and with fewer rights (`Sites.Read.All`), but it produces a report without site owners, without groups and without permission levels — exactly where a permissions review starts.

### Output

| File | Contents |
|---|---|
| `SharePoint_Permissions_Detail_<ts>.csv` | One row per grant: scope, principal, permission levels, sharing link type, external yes/no, member count. Unreadable scopes appear as `ItemType = Error` with the reason in the `Error` column; `UnitKey` ties a row to the checkpoint unit that wrote it |
| `SharePoint_Permissions_Summary_<ts>.csv` | Per site: number of grants, unique scopes, webs, lists, folders/files with their own permissions, sharing links, anonymous links, external principals, `Everyone` grants |
| `SharePoint_Permissions_SiteAccess_<ts>.csv` | **Per site, one row per person**, with the group the access comes through and the level. See below |
| `SharePoint_Permissions_Groups_<ts>.csv` | Per group, one row per member — SharePoint groups, the Entra groups nested inside them, *and* Entra groups granted directly on a scope. All flattened to people |
| `SharePoint_Permissions_EffectiveAccess_<ts>.csv` | Only with `-IncludeEffectiveAccess`: one row per user per scope, with the group that access comes through |

> On a large tenant the effective-access file can be orders of magnitude larger than the detail CSV. That is why it is off unless you ask for it.

#### Who has access where, through which group — in one tab

The question you usually open this report with is not "which grants exist" but **"who can get into this SharePoint, and how did they get there"**. That used to be scattered: `Rechten` said *that* a group had permissions, `Groepen` said who was in it, and you had to connect the two yourself. `Site Owners has Full Control` plus `Site Owners contains five people` is not an answer yet.

That is why there is `SharePoint_Permissions_SiteAccess_<ts>.csv` (tab `Toegang`): **one row per person per site**, with the group the access comes through and the level.

| Column | Contents |
|---|---|
| `SiteTitle` / `SiteUrl` | The site, by name — 130 URLs are not "at a glance" |
| `UserDisplayName` / `UserPrincipalName` / `UserEmail` | Who |
| `IsExternal` / `AccountEnabled` | Guest or internal, account active |
| `ViaType` | `Direct`, `SharePointGroup`, `SecurityGroup`, `M365Group`, `Everyone`, … |
| `ViaName` / `ViaId` | Which group, or `(direct toegekend)`. The id is included because a title like `Site Owners` exists on every site |
| `PermissionLevels` | The level of that grant |

It is deliberately **consolidated per site collection**: someone who ends up on thirty folders in the same site through the same group is one row — not thirty. A different level or a different group *is* a separate row, because that is different access. To see it per individual folder or file, use `-IncludeEffectiveAccess`; that tab (`Effectief`) is per scope and therefore much larger.

Three things that are deliberately not dropped here:

- **Directly granted people** appear as themselves, with `ViaType = Direct` and `ViaName = (direct toegekend)`.
- **`Everyone` and `Everyone except external users`** resolve to nobody, but they are exactly what you want to see. They get one row with the claim as the name.
- Even with `-SkipGroupExpansion`, directly granted people stay visible; only the group members are missing then.

> Compared with [NovaPoint](https://github.com/Barbarur/NovaPoint/wiki/Solution-Report-PermissionsReport), which answers the same question with `AccessType` + `GroupId` and a `Users` column holding a list of users: here every user is on a row of their own. That reads less compactly, but it is the difference between being able to filter or pivot on a person and not.

#### Everything in one Excel file

With `-Excel` you get one workbook next to the CSVs, `SharePoint_Permissions_<ts>.xlsx`, with a tab per report:

| Tab | Contents |
|---|---|
| `Samenvatting` | Per site: grants, unique scopes, sharing links, external principals, `Everyone` grants, errors |
| `Rechten` | Every grant individually |
| `Toegang` | **Per site, per person: which permission and through which group.** The tab to start with |
| `Groepen` | Every group with its members — SharePoint groups, the Entra groups nested inside them, **and** Entra groups granted directly on a scope |
| `Effectief` | Only with `-IncludeEffectiveAccess`: one row per user per scope |

Every tab is a real Excel table, so it has filter buttons and a frozen header row. Numbers come in as numbers, not as text, so summing and sorting work without converting first.

#### Pivot tables

Five ready-made pivot tables are added, each on a tab of its own:

| Tab | Rows | Columns | Value | Filters |
|---|---|---|---|---|
| `Pivot rechten` | Site | Permission level | Number of grants | Principal type, scope type |
| `Pivot principals` | Principal | Scope type | Number of scopes | Site, external yes/no |
| `Pivot groepen` | Group | Member external yes/no | Number of members | Site, group type |
| `Pivot toegang` | **Site → group → person** | Permission level | Count | External yes/no, access type |
| `Pivot per persoon` | Person → site → group | Permission level | Count | External yes/no, access type |

`Pivot toegang` follows how SharePoint actually hands out permissions: a site has groups, and groups have people. Collapsed, you see which groups sit on a site; expanded, who those groups let in. `Pivot per persoon` reads the same data from the other side — what does *this* person reach and through what — which is the question during an offboarding.

> A directly granted person has no group. `ViaName` then holds `(direct toegekend)` instead of nothing: an empty level in the hierarchy reads as missing data, not as "granted without a group".

> **`PermissionLevels` cannot be pivoted, `PrimaryPermission` can.** SharePoint often gives a grant several levels at once, and those sit in one column as `Read; Limited Access`. A pivot table turns that into a separate value, so `Full Control` and `Full Control; Limited Access` end up on different rows. That is why the `Rechten` and `Effectief` tabs have an extra `PrimaryPermission` column next to the full text, holding the heaviest level of that grant. `Limited Access` always loses to a real level — SharePoint sets that itself so someone can navigate to something granted deeper down. A custom permission level counts heavier than `Lezen` (Read) but lighter than `Volledig beheer` (Full Control): it was created on purpose, so it should not drop out. Both Dutch and English level names are recognised.

Want to build a pivot table yourself: put the cursor in a tab and choose **Insert → PivotTable**; the table is already named, so the range is right straight away and grows with the data.

The CSVs are always kept; the workbook comes on top of them. That is on purpose: the CSVs are where the scan streams to and what a resumed run appends to, so they exist anyway — and if writing the workbook goes wrong (module missing, file open in Excel, not enough memory) it never costs you the report itself.

> **Row limit.** An Excel worksheet stops at 1,048,576 rows and silently drops the rest. The script therefore deliberately cuts off at 1,000,000 and tells you which tab was truncated and which CSV holds the full data. Only `Effectief` realistically gets near that on a large tenant.

`-Excel` needs the `ImportExcel` module (it is in `Install-Modules.ps1`). If it is missing, the script says so and the CSVs are kept as normal.

### Resuming after an interruption (checkpoints)

After every completed list a checkpoint is written to the output folder: `SharePoint_Permissions_<hash>.state.json`, `.keys.partial.log` and `.detail/.groups/.effective.partial.csv`. Unlike the two scripts above, the rows are streamed straight to those partials instead of being kept in memory — a tenant-wide run at item level produces millions of rows. The summary CSV therefore does not exist as a checkpoint: it is built from the detail file at the end, so that a resumed run summarises everything ever written for this `<hash>`, not just the part from the last session. The `<hash>` comes from the scan parameters, so starting again with the same parameters resumes from the last completed list. `-Restart` discards that checkpoint and starts over. The checkpoint files are only cleaned up once the final CSVs are on disk — if they are still there, the previous run was interrupted.

Which units are already done is kept in `.keys.partial.log`, one line per key, append-only. That is deliberately not a list in the JSON: rewriting it sorted after every list is quadratic work, and on a tenant with thousands of lists the checkpoint would then cost more time than the scanning itself. A half-written last line (process hard-killed while writing) is ignored — that one unit is simply scanned again.

### Robustness

A tenant-wide run takes hours and touches thousands of objects, so the failures below are not edge cases but to be expected. How the script handles them:

| Situation | Behaviour |
|---|---|
| App role not yet replicated | Before the scan, the token has to prove *that* it carries the roles (`roles` claim). Entra will happily issue a token without the role that was just granted, and Graph answers that with `401` — not `403`. Such a token is not cached; a new one is minted until the role is in it (up to ~2.5 min), followed by a clear error |
| Token rejected (`401`) | Re-authenticate once with a fresh token; if it is still rejected, **the run stops** with the reason from `x-ms-diagnostics`. A 401 never applies to just one site, so it is not reported per site |
| No access to one site (`403`) or object gone (`404`) | That one site/object is skipped, the rest carries on |
| Throttling (`429`/`503`) | Retry with `Retry-After`, otherwise exponential backoff up to 3 minutes |
| One list fails (view threshold, odd template) | Error row in the detail CSV, the other lists of that site carry on as normal. The unit is **not** marked as done, so a resumed run tries it again |
| List with unique permissions but no role assignments | Produces no rows, and that is correct. This used to crash on `Cannot bind argument to parameter 'RoleAssignments'` |
| Role assignments unreadable (`403`) | **Error row, not an empty result.** This is the only place where a 403 is not skipped: an empty list of role assignments reads as "nobody has permissions on this scope", and "I am not allowed to look" is a different fact from "there is nothing to see" |
| Gallery list rejects the field selection (`400`) | The query is narrowed step by step (four variants) until SharePoint accepts it. `Galerie van thema's` and `Galerie met basispagina's` do not have every field; better to lose the file name than the unique scopes of that list. Which variant works is a property of the list **template**, so it is learned once and reused for every later site — on a 131-site tenant that halves the round trips this costs, and the log reports each quirk once instead of once per site |
| `Lijst met gebruikersgegevens` (template 112) | The item scan is skipped. SharePoint rejects `/items` on this hidden system list at *any* field width, and the items are user records, not content — item permissions mean nothing there. The list itself is still reported. This saved 121 false error rows per tenant scan |
| Error rows from an earlier, aborted attempt | Are left out when writing as soon as the same unit later *did* succeed. Otherwise the report counts errors that have already been resolved, and the "N scopes not readable" number is wrong — exactly the number someone acts on |
| Item sweep fails | A separate error row: without it the list would look as if nothing in it had unique permissions |
| CSV is open in Excel | Five retries with increasing wait time; if it still fails, the run stops instead of silently dropping rows |
| Token expires in the middle of a large library | Every wave fetches the token again — workers get a copy and would not see a later refresh |
| SharePoint repeats a paging link | Detected and aborted instead of looping forever |
| Unexpected error anywhere | A `trap` cleans up the temporary App Registration before the script stops — a Full Control app is never left behind |

> At the end the script states explicitly whether *everything* could be read. If there is a `[WARN]` about unreadable scopes, filter the detail CSV on `ItemType = Error`: a report with gaps must not look like a report without findings.

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-TenantUrl` | string | — | Tenant root, e.g. `https://contoso.sharepoint.com`. Required for a tenant-wide run |
| `-SiteUrl` | string | — | One site collection (including subsites) instead of the whole tenant |
| `-Scope` | `Site`/`List`/`Item` | `Item` | How deep: webs only, webs + lists, or everything down to folder and file level |
| `-TenantId` | string | _(from the session)_ | Entra tenant ID. Required with `-ClientId` |
| `-ClientId` | string | — | Existing App Registration; skips the temporary app |
| `-ClientSecret` | string | — | Secret for `-ClientId`. **Does not work against SharePoint** (see Authentication); the script warns you |
| `-CertificateThumbprint` | string | — | Certificate for `-ClientId`, from `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`. This is the variant that works |
| `-AppOnly` | switch | off | ClientId and CertificateThumbprint for the tenant from `graph.appid.json` instead of a temporary app. That app needs the roles above, SharePoint `Sites.FullControl.All` included (PowerShell 7) |
| `-OutputPath` | string | `C:\Temp` | Output folder |
| `-IncludeOneDriveSites` | switch | off | Also includes personal OneDrive sites (one site per user) |
| `-IncludeHiddenLists` | switch | off | Includes hidden and system lists (Form Templates, Style Library, workflow history, …) |
| `-ListTitle` | string[] | _(all)_ | Limit to one or more list/library titles |
| `-ExcludeLimitedAccess` | switch | off | Leaves out `Limited Access` assignments. SharePoint sets those itself so someone can navigate to an item granted deeper down — noise in most reviews, but they *do* explain why someone sees a folder path |
| `-SkipGroupExpansion` | switch | off | Do not resolve group membership. Faster, but then you only know *which* group has access, not who is in it |
| `-IncludeEffectiveAccess` | switch | off | Also writes the effective-access CSV |
| `-Excel` | switch | off | Also writes a single `.xlsx` with a tab per report. Requires `ImportExcel` |
| `-GraphTimeoutSec` | int | `120` | Timeout per Graph/SharePoint call |
| `-MaxGraphRetry` | int | `6` | Number of retries on throttling or timeouts |
| `-Concurrency` | int (1-8) | `4` | Parallel workers for the per-item lookups that dominate a `-Scope Item` run. `1` turns parallelism off |
| `-Restart` | switch | off | Ignore an existing checkpoint and start over |

### Requirements

- Module `Microsoft.Graph.Authentication` (and `Microsoft.Graph.Applications` as long as you let the script create the temporary app) — `.\scripts\Startup\Install-Modules.ps1`
- An account that is allowed to create an App Registration and grant admin consent, unless you pass `-ClientId` of an existing app

### Examples

```powershell
# Everything, tenant-wide: sites, subsites, libraries, folders and files with unique permissions
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com"

# One site collection, plus a CSV with effective access per user
.\Get-SharePointPermissionsReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -IncludeEffectiveAccess

# Everything in one Excel workbook: summary, permissions, groups with members, and effective access
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -IncludeEffectiveAccess -Excel

# Faster overview: stop at list/library level and hide automatic traversal grants
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -Scope List -ExcludeLimitedAccess

# Ignore an interrupted tenant scan and start over from scratch
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -Restart
```

---

## Remove-SharePointFileVersionsByDate.ps1

Reports or deletes **old file versions** in SharePoint Online document libraries based on a cutoff date, while the **current version is kept**.

### Behaviour

- By default: preview/reporting only
- With `-Apply`: actually deletes the matching previous versions
- Works on one site or tenant-wide across all sites
- By default no OneDrive sites and no hidden libraries
- Based on Microsoft Graph (`Invoke-MgGraphRequest`) — **no** `PnP.PowerShell` and **no** Entra app registration of your own needed for the default case

### Authentication

By default the script signs in delegated with `Sites.ReadWrite.All` + `Files.ReadWrite.All` through [`Connect-M365Graph`](../Startup/readme.md#connect-m365ps1) — a device code when `useDeviceCodeAuth` is set in `load.config.ps1`, the GDAP customer from `Connect-Tenant`, Microsoft's own pre-consented app, so no App Registration or `-ClientId` of your own; a Graph session that already holds the scopes is reused and left open. Requires PowerShell 7. Only a **tenant-wide scan** (no `-SiteUrl`) additionally needs a short-lived, read-only temporary App Registration (`Sites.Read.All`) for site/library enumeration *and* fetching version history — Microsoft does not support tenant-wide site enumeration delegated. With `-VersionBatchConcurrency` above `1` (the default) a **second** temporary App Registration is also created, purely to double the throughput of version lookups: SharePoint's "activityLimitReached" throttle applies per app registration, so two apps each get their own throttle budget (same approach as `Get-SharePointStorageReport.ps1`). Both temporary apps are deleted again afterwards. Version **deletes** always go through your own delegated permissions, never through a temporary app.

Want to skip the temporary app(s) and use your own existing app registration? Then pass `-ClientId` + `-TenantId` + `-ClientSecret` (or `-CertificateThumbprint`), or `-AppOnly` to take ClientId and CertificateThumbprint from `graph.appid.json`; that app must already have the `Sites.ReadWrite.All` application permission.

> **Note:** deleting a specific version (`DELETE .../versions/{id}`) is not in Microsoft's official Graph API reference, but it is a widely used operation that is confirmed to work (for both OneDrive and SharePoint document libraries). The current/latest version cannot be deleted this way — Graph refuses that, which is exactly the keep-the-current-version guarantee.

### Resuming after an interruption (checkpoints) and progress

Like `Get-SharePointStorageReport.ps1`, this script writes a checkpoint to the output folder after every completed library: `SharePoint_VersionCleanup_<hash>.state.json` + `.summary.partial.csv` + `.detail.partial.csv`. The `<hash>` is derived from all scan parameters (cutoff date, site, mode, output folder, etc.):

- **Starting again with the same parameters** resumes automatically — libraries already done are skipped (`[SKIP] Already completed in a previous run.`).
- **`-Restart`** discards the checkpoint and starts from scratch. This only affects the progress tracking, not what (with `-Apply`) has already actually been deleted in SharePoint itself — deleted versions obviously stay deleted.
- The checkpoint files are cleaned up automatically once the scan completes successfully.

During the scan the script shows nested progress bars (sites → libraries → scanning folders/files / fetching version history) alongside the scrolling log lines, and a `[WAIT] throttled by Microsoft Graph — waiting ...` message as soon as Graph throttles, so a long pause does not feel like a hang.

> **Note (delegated/SDK calls):** as with `Get-SharePointStorageReport.ps1`, the script sets `Set-MgRequestContext -ClientTimeout <-GraphTimeoutSec> -MaxRetry 0` right after connecting — without that setting the Microsoft.Graph SDK cmdlets silently retry 429/503 themselves with their own backoff, which on `activityLimitReached` can cause minutes of silence without the script's own `[WAIT]` message showing up.

### Parameters

| Parameter | Type | Description |
|---|---|---|
| `-BeforeDate` | `datetime` | Delete versions older than this date |
| `-SiteUrl` | `string` | Optional: scan one site |
| `-TenantUrl` | `string` | Required for an all-sites scan, e.g. `https://contoso.sharepoint.com` |
| `-TenantId` | `string` | Entra ID tenant ID — detected automatically if not given; required together with `-ClientId` |
| `-ClientId` | `string` | Existing App Registration client ID — skips the temporary app; use together with `-TenantId` and `-ClientSecret` or `-CertificateThumbprint` |
| `-ClientSecret` | `string` | Client secret for an existing app registration |
| `-CertificateThumbprint` | `string` | Certificate thumbprint for an existing app registration |
| `-AppOnly` | `switch` | App-only with ClientId and CertificateThumbprint for the tenant from `graph.appid.json` (needs `Sites.ReadWrite.All` as application permission) |
| `-Apply` | `switch` | Actually performs the deletion |
| `-IncludeOneDriveSites` | `switch` | Includes OneDrive sites in the tenant scan |
| `-IncludeHiddenLibraries` | `switch` | Includes hidden document libraries |
| `-LibraryTitle` | `string[]` | Optional filter on library title |
| `-GraphTimeoutSec` | `int` | Timeout in seconds per Graph call (default: `120`) |
| `-MaxGraphRetry` | `int` | Maximum number of retries on Graph throttling/timeouts (default: `6`) |
| `-VersionBatchConcurrency` | `int` | Number of parallel `$batch` workers for fetching version history in tenant-wide scans, 1-8 (default: `4`). Above `1` the second temporary app is created as well (see Authentication) |
| `-MaxVersionRetryPasses` | `int` | Maximum number of retry passes for fetching version lists under sustained throttling. `0` (default) scales automatically with the number of files — same approach and reason as for `Get-SharePointStorageReport.ps1` above |
| `-Restart` | `switch` | Discards an existing checkpoint for this parameter combination and starts the scan from scratch |

### Examples

```powershell
# Tenant-wide preview: everything older than 1 January 2025
.\Remove-SharePointFileVersionsByDate.ps1 `
    -TenantUrl "https://contoso.sharepoint.com" `
    -BeforeDate "2025-01-01"

# Actually delete on one site
.\Remove-SharePointFileVersionsByDate.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
    -BeforeDate "2025-01-01" `
    -Apply

# Ignore an interrupted tenant scan and start over from scratch
.\Remove-SharePointFileVersionsByDate.ps1 `
    -TenantUrl "https://contoso.sharepoint.com" `
    -BeforeDate "2025-01-01" `
    -Apply -Restart
```
