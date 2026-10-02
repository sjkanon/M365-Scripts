**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **SharePoint**

# SharePoint Scripts

SharePoint Online and OneDrive content operations via PnP PowerShell, with interactive
admin sign-in on any (customer) tenant.

---

## Folders

| Folder | Description |
|--------|-------------|
| [`Provisioning/`](Provisioning/readme.md) | Provision and maintain a whole structure — metadata model, content types, libraries and group permissions — from one config file, plus a sharing audit and a drift check |

## Scripts

| Script | Description |
|--------|-------------|
| [`Find-SiteContent.ps1`](Find-SiteContent.ps1) ([docs](#find-sitecontentps1)) | Search a whole site (name, path, type, size, date or full text) and report the permissions on every hit — PnP/CSOM, signs in as you |
| [`Search-SharePointContent.ps1`](Search-SharePointContent.ps1) ([docs](#search-sharepointcontentps1)) | The same question tenant-wide through Microsoft Graph, app-only, no interactive login — files and folders |
| [`Restore-RecycleBinItems.ps1`](Restore-RecycleBinItems.ps1) ([docs](#restore-recyclebinitemsps1)) | Restore deleted files/folders from a site or OneDrive recycle bin (dry-run by default) |
| [`Trace-SharePointFile.ps1`](Trace-SharePointFile.ps1) ([docs](#trace-sharepointfileps1)) | Where did a file go? Renames, moves, copies and deletes from the audit log, folder moves included — in Brussels time, over a period you choose |
| [`Revoke-SharePointUserAccess.ps1`](Revoke-SharePointUserAccess.ps1) ([docs](#revoke-sharepointuseraccessps1)) | Take one user's access away everywhere: site collection admin, direct grants at every level, SharePoint groups and sharing links. Reports by default, removes with `-Apply` |
| [`Test-SharePointAccessScripts.ps1`](Test-SharePointAccessScripts.ps1) ([docs](#test-sharepointaccessscriptsps1)) | Verify the two access scripts without touching a tenant — shared auth block identical, and the revocation funnel behaves |

---

## Which of the two search scripts?

Both answer "where does this live and who can reach it", and both write the same kind of
CSV. They differ in what they can see and in what they need from you.

| | `Find-SiteContent.ps1` (PnP) | `Search-SharePointContent.ps1` (Graph) |
|---|---|---|
| Sign-in | Interactive, as you | App-only, unattended — fine in a scheduled task |
| Rights needed | Access to the site, or `-GrantSiteAdmin` per site | One admin consent, once, for the whole tenant |
| Scope | One site collection (+ subsites) | One site, or **every site in the tenant**, OneDrive included |
| Files and folders | Yes | Yes, and faster — `delta` reads a library in pages of a thousand and permissions come 20 per `$batch` |
| Items in ordinary lists | Yes, with their permissions | **No** — Graph exposes permissions for driveItems only |
| Site- and list-level rights | Yes: "this inherits from the library, which grants Edit to Site Members" | **No** — Graph has no API for SharePoint role assignments; it says *whether* an item inherits and from where, not what the site grants |
| Sharing links | Yes, from the `SharingLinks.*` groups | Yes, richer: link scope, edit/view, expiry date and the link URL itself |

Rule of thumb: **Graph** for "find it anywhere in the tenant and show me the links and
guests on it", **PnP** when you need the full permission story of one site, including its
lists and its groups.

---

### Find-SiteContent.ps1

Answers the two questions you normally have at the same time: **where does this live**
and **who can get at it**. Read-only — the script never changes anything.

**Two engines**

| Engine | When | What it sees |
|--------|------|--------------|
| Crawl (default) | No `-Content` given | Walks every list and library. `-Name` and `-ItemType` are pushed into a CAML query where possible, so SharePoint returns only the matches; anything CAML cannot express falls back to reading that list in full. Sees everything, also what the search index has not picked up yet |
| Search (`-Content`) | Full-text query | A KQL query against the search index, scoped to the site path — this is the one that matches text *inside* documents. Fast, but limited to what is indexed and what the signed-in account may see. Covers subsites automatically |

Both engines feed the same filters: `-Name` (wildcards), `-Path`, `-Extension`,
`-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter` / `-ModifiedBefore`,
`-MinSizeMB`. `-Name` is matched against the file name, the item title *and* the last
segment of the URL, so an item whose title differs from its file name still turns up.

Hidden and system libraries are skipped unless you pass `-IncludeHidden`, and while
crawling only the top web is searched unless you add `-IncludeSubsites` — the script
reports per web how many lists it skipped and why. **`-Everything` turns all of that
off in one go**: every subsite, every hidden and system list, no cap on the hits and
none on the permission lookups. A crawl still matches names and metadata only; use
`-Content` to search inside the documents themselves.

**Permissions per hit**

For every match the script works out where the permissions actually come from:

| Source | Meaning |
|--------|---------|
| `Item` | The item broke inheritance and carries its own role assignments |
| `List` | It inherits from a library/list that has unique permissions |
| `Site` | It inherits all the way up to the (sub)site |

Role assignments are flattened to one CSV row per principal — principal type, login,
e-mail and the role names (`Full Control`, `Edit`, …). `Limited Access` is hidden
unless you pass `-IncludeLimitedAccess`; those entries only exist so someone can reach
a deeper item and grant nothing by themselves.

Three things are called out separately because they are the ones that surprise people:

- **Sharing links** — the `SharingLinks.*` groups behind every "Copy link". Always
  expanded to the people in them and labelled *Anyone* / *Organization* / *Specific
  people*, so an anonymous link cannot hide in the noise
- **External users** — guest accounts (`#ext#`) in any assignment
- **Everyone** — "Everyone" and "Everyone except external users"

Site and list permissions are read once and cached, item permissions only for items
that actually broke inheritance, so a search with a handful of hits costs a handful of
extra calls. `-Permissions Unique` is the fast way to answer "what in this site is
shared differently from the rest"; `-Permissions None` skips permissions entirely.
`-MaxPermissionLookups` (default 1000) keeps a too-broad search from running for hours.

**Sign-in**

Same as `Restore-RecycleBinItems.ps1`: the first run against a tenant registers a
public-client Entra app (delegated `AllSites.FullControl`, admin-consented) and caches
the client ID per tenant in `pnp.appid.json` in the repo root (gitignored). A client ID
cached by the other script is reused, so this usually costs nothing. Pass `-ClientId`
to skip app registration entirely.

Reading permissions requires access to the site. `-GrantSiteAdmin` makes the signed-in
admin site collection administrator for the duration of the run and removes the rights
again afterwards (keep them with `-KeepSiteAdmin`) — that is what makes searching
someone else's OneDrive or a site you are not a member of possible.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SiteUrl` | Yes | Site collection to search (team site, communication site, or a OneDrive) |
| `-IncludeSubsites` | No | Also crawl every subsite below it |
| `-Content` | No | Full-text/KQL query — switches to the search index |
| `-Name` | No | Filter on the name, wildcards allowed (`*offerte*`) — matched against the file name, the item title *and* the last URL segment |
| `-Path` | No | Filter on the folder, substring match on the server relative URL |
| `-Extension` | No | One or more extensions, with or without the dot (`xlsx`,`pdf`) |
| `-ItemType` | No | `All` (default), `File`, `Folder` or `ListItem` |
| `-ListName` | No | Only these lists/libraries by title, wildcards allowed |
| `-ModifiedBy` | No | Who last changed it — display name or e-mail, wildcards allowed |
| `-ModifiedAfter` / `-ModifiedBefore` | No | Restrict to a change window |
| `-MinSizeMB` | No | Only files of at least this size |
| `-IncludeHidden` | No | Also search hidden lists, catalogs and system libraries |
| `-Everything` | No | Leave nothing out: `-IncludeSubsites -IncludeHidden` plus no caps at all |
| `-Permissions` | No | `Effective` (default), `Unique` (only broken inheritance) or `None` |
| `-ExpandGroups` | No | Also list the members of regular SharePoint groups (link groups are always expanded) |
| `-IncludeLimitedAccess` | No | Keep `Limited Access` assignments in the report |
| `-MaxItems` | No | Stop after this many matches (default 5000, `0` = no limit) |
| `-MaxPermissionLookups` | No | Cap on hits that get their permissions resolved (default 1000, `0` = no limit) |
| `-PageSize` | No | Items per server call while crawling (default 500) |
| `-GrantSiteAdmin` | No | Temporarily make yourself site collection admin (SharePoint Administrator required) |
| `-KeepSiteAdmin` | No | Keep those rights instead of removing them afterwards |
| `-AdminUpn` | No | UPN to grant site admin to (default: the signed-in account) |
| `-TenantId` / `-ClientId` / `-AppName` | No | Sign-in overrides, as in `Restore-RecycleBinItems.ps1` |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\SharePointFind_<timestamp>.csv`) |
| `-Disconnect` | No | Sign out of PnP when finished |

**Examples**

```powershell
# Where does anything with "offerte" in the name live, and who can see it?
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Sales -Name "*offerte*"

# Full text: which documents mention "salarisschaal", anywhere in the site tree?
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/HR `
    -Content "salarisschaal"

# Leave nothing out: all subsites, all hidden/system lists, no caps
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*veiligheid*" -Everything

# Everything in the site that is shared differently from the rest
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Permissions Unique -IncludeSubsites

# Large PDFs in one library, with the groups behind the permissions expanded
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -ListName "Gedeelde documenten" -Extension pdf -MinSizeMB 10 -ExpandGroups

# Search a OneDrive you have no rights on
.\Find-SiteContent.ps1 `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
    -Name "*.xlsx" -GrantSiteAdmin
```

**Notes**
- The CSV has one row per hit *per principal*, so it filters and pivots well: sort on
  `SharingLink`, `External` or `UniqueRights` to get straight to the interesting rows.
- `-Content` only finds what the search index knows. Freshly uploaded or recently
  changed documents can take minutes to hours to show up — crawl mode always sees them.
- A crawl reads every item in every library. On a site with hundreds of thousands of
  items that takes a while; narrow it with `-ListName` or use `-Content` instead.

---

### Search-SharePointContent.ps1

The Graph counterpart of `Find-SiteContent.ps1`: same question, app-only, and it reaches
every site and every OneDrive in the tenant without you having rights on any of them.
Read-only.

**Two engines**

| Engine | When | What it does |
|--------|------|--------------|
| Delta (default) | No `-Content` | `/drives/{id}/root/delta` walks a whole library tree in pages of a thousand items. Filtering happens client-side, so `*contains*` wildcards work |
| Search (`-Content`) | Full-text query | `/search/query` against the search index — matches text *inside* documents. KQL only does trailing wildcards (`veiligheid*`), not leading ones. An app-only search must name a geography; the script reads it from the site's data location, or finds it by trying, and `-Region` overrides |

**Permissions**

`/drives/{id}/items/{id}/permissions`, fetched 20 at a time through `/$batch`. That single
call carries the whole picture per item:

| Field | Reported as |
|-------|-------------|
| `roles` | `Read` / `Edit` / `Full Control` |
| `grantedToV2` / `grantedToIdentitiesV2` | The user, Entra group, SharePoint group or site user — one CSV row each |
| `link` | `SharingLink` (Anyone / Organization / Specific people, view or edit), `LinkUrl`, `LinkExpires` |
| `inheritedFrom` | Absent → `PermissionSource = Item` (unique rights). Present → `Inherited`, with the folder it comes from |

"Anyone with the link" permissions are counted separately in the summary and shown in red,
because those are reachable without signing in at all. External guests (`#EXT#`) and
"Everyone except external users" get their own counters too.

**What Graph cannot do** — worth knowing before you reach for this one:

- **No site- or list-level permissions.** There is no Graph API for SharePoint role
  assignments; `/sites/{id}/permissions` only returns app grants (Sites.Selected). This
  script tells you whether an item inherits and from which folder, not what the site
  itself grants to which group. Use `Find-SiteContent.ps1` for that.
- **Libraries only.** Permissions exist for driveItems, not for items in ordinary lists.
- No role definitions, and no "Limited Access" nuance.

**Sign-in — app-only, created for you**

The first run against a tenant sets the app up:

1. Sign in to Graph as a Global Administrator (once)
2. Create or reuse an app named after `-AppName`
3. Grant and admin-consent the **application** role `Sites.Read.All` (plus `Group.Read.All`
   if you use `-ExpandGroups`)
4. Create a self-signed certificate in `Cert:\CurrentUser\My` and upload its public key —
   **no secret is written to disk**
5. Cache the client ID and thumbprint per tenant in `graph.appid.json` in the repo root
   (gitignored)

Later runs connect app-only with no prompt at all, which is what makes this one usable from
a scheduled task. Bring your own app with `-ClientId` plus `-CertificateThumbprint` or
`-ClientSecret`. The certificate is non-exportable and lives in the creating user's store,
so a scheduled task has to run as that same account.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SiteUrl` | * | One site collection (or a OneDrive) to search |
| `-AllSites` | * | Search every site in the tenant instead |
| `-SiteFilter` | No | With `-AllSites`: only sites whose URL matches this wildcard |
| `-MaxSites` | No | With `-AllSites`: stop after this many sites |
| `-IncludePersonalSites` | No | With `-AllSites`: include everybody's OneDrive |
| `-IncludeSubsites` | No | With `-SiteUrl`: also search its subsites |
| `-Content` | No | Full-text/KQL query — switches to the search index |
| `-Region` | No | Geography for `-Content` (`EUR`, `NAM`, `DEU`, …). App-only search requires one; detected automatically when possible |
| `-Name` | No | Filter on the file/folder name, wildcards allowed |
| `-Path` | No | Filter on the folder path, substring match |
| `-Extension` | No | One or more extensions, with or without the dot |
| `-ItemType` | No | `All` (default), `File` or `Folder` |
| `-LibraryName` | No | Only these document libraries, wildcards allowed |
| `-ModifiedBy` | No | Who last changed it — display name or e-mail |
| `-ModifiedAfter` / `-ModifiedBefore` | No | Restrict to a change window |
| `-MinSizeMB` | No | Only files of at least this size |
| `-Permissions` | No | `Effective` (default), `Unique` or `None` |
| `-ExpandGroups` | No | Also list Entra group members (needs `Group.Read.All`) |
| `-Everything` | No | Subsites, personal sites, and no caps |
| `-MaxItems` | No | Stop after this many matches (default 5000, `0` = no limit) |
| `-MaxPermissionLookups` | No | Cap on permission lookups (default 2000, `0` = no limit) |
| `-TenantId` | * | Required with `-AllSites`; derived from `-SiteUrl` otherwise |
| `-ClientId` / `-CertificateThumbprint` / `-ClientSecret` / `-AppName` | No | Bring your own app registration |
| `-OutputPath` | No | CSV path (default: `C:\Temp\GraphSharePointFind_<timestamp>.csv`) |
| `-MaxRetries` | No | Retries on throttling (429), default 5 |

**Examples**

```powershell
# One site plus its subsites, everything with "veiligheid" in the name
.\Search-SharePointContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*veiligheid*" -IncludeSubsites

# Tenant-wide full text: which documents mention "salarisschaal"?
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -Content "salarisschaal"

# Everything in the tenant that carries permissions of its own — start with 25 sites
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -Permissions Unique -MaxSites 25

# Unattended, with an app you already have
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -ClientId 0000...-4444 -CertificateThumbprint A1B2C3 `
    -Name "*.pfx" -OutputPath C:\Reports\keys.csv
```

**Notes**
- A tenant-wide delta crawl reads every item of every library it touches. Start with
  `-MaxSites`, and add `-IncludePersonalSites` only when you mean it — that multiplies the
  work by the number of users.
- Throttling (HTTP 429) is handled: single calls and batch sub-requests are retried,
  honouring `Retry-After`.
- `-Content` searches the whole tenant index and the results are filtered back to the sites
  in scope afterwards, so a query with very many hits spends some time on results it drops.

**Required modules**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser  # to run
Install-Module Microsoft.Graph.Applications  -Scope CurrentUser   # only for the one-time app registration
```

---

### Restore-RecycleBinItems.ps1

Lists the first- and/or second-stage recycle bin, filters the items (name, path, who
deleted them, when), and restores the matches to their original location. Dry-run by
default — nothing is restored until you pass `-Apply`.

**Two modes**

| Mode | What it does |
|------|--------------|
| `-SiteUrl <url>` | One site collection — team site, communication site, or a user's OneDrive (`https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com`) |
| `-AllSites -TenantUrl <url>` | Every SharePoint site in the tenant. **OneDrive personal sites are excluded** — restore those one at a time with `-SiteUrl` |

The tenant-wide sweep also skips the My Site host, redirect sites, and sites that are
locked read-only or no-access (a restore there fails anyway). It reports how many of each
it skipped. Narrow it with `-SiteFilter "*/sites/Finance*"` and try it out with
`-MaxSites 5` before letting it loose on the whole tenant.

A site that errors out (no access, throttled, gone) is reported and the sweep continues —
one bad site does not abort the run. Per-site results end up in the summary table and in
the CSV, which carries a `Site` column in both modes.

Tenant-wide is where a dry run earns its keep: it answers "which sites hold files this
account deleted on Tuesday" without touching anything.

**Speed — batches, not threads**

Restores go through `Restore-PnPRecycleBinItem -IdList`, which hands SharePoint a whole set
of items in a single server call. 200 files then cost one round trip instead of 200, which
is where the speed comes from — a full library restore drops from tens of minutes to a
couple of minutes.

Running restores in parallel runspaces is *not* the faster route here and the script
deliberately does not do it:

- PnP PowerShell is not thread-safe; each runspace needs its own module import (seconds)
  and its own connection, and connection objects do not cross runspace boundaries cleanly
- All items land in the same site collection, so parallel calls contend on the same lists
  and trip SharePoint throttling (HTTP 429) — you get retries and partial failures, not speed
- One batched call already does server-side what the threads were trying to parallelise

Batch behaviour:

- `-BatchSize` (1-200, default 200) controls the batch size; `-BatchSize 1` restores strictly
  item by item
- Folders and files never share a batch, so the folders-first ordering survives
- A batch is all-or-nothing and its error does not say which item broke it, so a failed batch
  is automatically retried one item at a time — the good items still come back and the CSV
  names the ones that did not

**How long does it take?**

The script tells you, so you know whether to wait or grab coffee:

- Reading the recycle bin is timed and reported per site — on a site with tens of thousands
  of deleted items this alone can take minutes (use `-RowLimit` to cap it). Tenant-wide,
  this reading is usually the bulk of the runtime, not the restoring
- During `-Apply` the progress bar shows elapsed time and a live ETA, recalculated from the
  rate actually measured against that tenant instead of an up-front guess. With `-AllSites`
  there are two bars: sites, and batches within the current site
- The summary reports the real duration and, tenant-wide, a per-site table with items found,
  matched, restored, failed and the read/restore seconds for each
- The CSV has `Site`, `Batch` and `DurationSeconds` columns, which is how you spot the slow
  ones

**Sign-in — the app registration is created for you**

PnP PowerShell no longer ships a shared multi-tenant app, so an Entra app registration is
required. The first run against a tenant creates one automatically:

1. Sign in to Microsoft Graph as a Global Administrator of the target tenant
2. The script creates (or reuses) a public-client app named after `-AppName`
3. It grants and admin-consents the delegated SharePoint scope `AllSites.FullControl`
4. The client ID is cached per tenant in `pnp.appid.json` in the repo root (gitignored)

Later runs read the cached client ID and go straight to the interactive SharePoint login —
the Graph admin sign-in only happens once per tenant. Pass an existing `-ClientId` to skip
app creation entirely.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SiteUrl` | * | Site collection URL to restore in (team site, communication site, or OneDrive) |
| `-AllSites` | * | Walk every SharePoint site in the tenant instead; OneDrive is excluded |
| `-TenantUrl` | * | Tenant URL for `-AllSites`, e.g. `https://contoso.sharepoint.com` |
| `-SiteFilter` | No | With `-AllSites`: only sites whose URL matches this wildcard |
| `-MaxSites` | No | With `-AllSites`: stop after this many sites |
| `-TenantId` | No | Tenant ID/domain for the one-time Graph sign-in (default: derived from `-SiteUrl`) |
| `-ClientId` | No | Existing Entra app to sign in with — skips app registration |
| `-AppName` | No | Display name of the app to create/reuse (default: `M365-Scripts SharePoint Restore`) |
| `-Name` | No | Filter on item name, wildcards allowed (`*.xlsx`, `Budget*`) |
| `-Path` | No | Filter on original location, substring match (`Shared Documents/Finance`) |
| `-DeletedBy` | No | Filter on who deleted it — display name or e-mail, wildcards allowed |
| `-DeletedAfter` / `-DeletedBefore` | No | Restrict to a deletion time window |
| `-ItemType` | No | `All` (default), `File`, `Folder` or `ListItem` |
| `-Stage` | No | `All` (default), `FirstStage` (user bin) or `SecondStage` (site collection admin bin) |
| `-RowLimit` | No | Cap on how many recycle bin entries are fetched (use on huge bins) |
| `-BatchSize` | No | Items per server call, 1-200 (default 200). `1` = strictly one by one |
| `-GrantSiteAdmin` | No | Temporarily add yourself as site collection admin per site — needed for another user's OneDrive and effectively required for `-AllSites` |
| `-KeepSiteAdmin` | No | Keep those rights instead of removing them afterwards |
| `-AdminUpn` | No | UPN to grant site admin to (default: the signed-in account) |
| `-Apply` | No | Actually restore. Without it the script only reports |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\RecycleBinRestore_<timestamp>.csv`) |
| `-Disconnect` | No | Sign out of PnP when finished |

**Examples**

```powershell
# Dry run — show everything in both recycle bins of a site
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance

# Restore everything one user deleted last night
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-27 18:00" -Apply

# Restore Excel files from one library folder, second-stage bin only
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*.xlsx" -Path "Shared Documents/Budget" -Stage SecondStage -Apply

# Restore a user's OneDrive, temporarily granting yourself site admin
.\Restore-RecycleBinItems.ps1 `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
    -GrantSiteAdmin -Apply

# Tenant-wide dry run: which sites hold files this account deleted today?
.\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" -GrantSiteAdmin

# Tenant-wide restore after a bulk delete — try 5 sites first
.\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" `
    -GrantSiteAdmin -MaxSites 5 -Apply
```

**Notes**
- Items are restored folders-first and shallow-paths-first: a file cannot be restored while
  the folder it lived in is itself still deleted.
- Restoring fails when an item with the same name already exists in the original location —
  those items are reported in the CSV with the SharePoint error, nothing else is aborted.
- Second-stage (site collection) items are restored straight back to their original
  location, not into the first-stage bin.
- The default retention is 93 days across both stages. Items older than that are gone and
  no script can bring them back — that is a Microsoft platform limit, not a script limit.
- Requires SharePoint Administrator (or Global Administrator) for `-GrantSiteAdmin` and for
  the one-time app registration; the restore itself only needs access to the site.

**Required modules**
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser              # PowerShell 7.4+
Install-Module Microsoft.Graph.Applications -Scope CurrentUser # only for the one-time app registration
```

---

### Trace-SharePointFile.ps1

Answers "where did my file go?" for OneDrive and SharePoint: renamed, moved, copied,
deleted or restored, by whom and when. Every time is shown in **Brussels time**
(summer and winter time handled, UTC offset alongside), and you can give it a period.
Read-only: nothing in the tenant changes.

SharePoint does not remember a file's old name or location; the **Unified Audit Log**
does, so that is the source. The script reads it and rebuilds the file's trail:

| Step | What it does |
|------|--------------|
| Read | Every rename, move, copy, delete, recycle, restore and upload of files **and folders** in the window. Read per day; a slice holding more than the 50,000 records one search can return is split until it fits (down to 15 minutes), and a failing or inconsistent search is retried |
| Seed | The records that mention the file by `-Name`, `-Url` or `-ItemId` |
| Follow | Every record of the same item (`ListItemUniqueId`, which survives renames and moves) and every record starting at a path the file was renamed or moved to. A chain `A → B → C` ends at C, even though C looks nothing like the name you searched for. A move to another site is followed by its destination path |
| Folders | Renaming, moving or deleting a folder moves every file in it **without a record per file**. Folder records are replayed against the file's path at that moment, so "moved along with folder X" and "deleted along with folder X" show up too |

Per item you get the timeline, the **last known location** and a status: `Present`,
`In recycle bin` (restore it with [`Restore-RecycleBinItems.ps1`](#restore-recyclebinitemsps1)),
`In second-stage recycle bin` or `Permanently deleted`.

**Parameters**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Name` | string | — | File name as it was at some point. Without an extension it matches any extension (`Offerte` finds `Offerte.docx`); `*` and `?` are wildcards |
| `-Url` | string | — | Full URL of the file as it was. `?web=1` is ignored and a `/:w:/r/` sharing link is turned back into the path. Opaque sharing links (`/:w:/s/`, `/:w:/g/`) cannot be traced — open them and copy the address they land on |
| `-ItemId` | guid | — | The `ListItemUniqueId`, e.g. from the CSV of an earlier run |
| `-SiteUrl` | string | — | Only records of this site or OneDrive. Much faster in a large tenant |
| `-StartDate` | date or string | `-Days` before `-EndDate` | Wall-clock time in `-TimeZone`: `15-09-2026`, `15/09/2026 08:30`, `2026-09-15 08:30` (day first, Belgian style) |
| `-EndDate` | date or string | now | Same notation. A date without a time includes that whole day |
| `-Days` | int | `30` | Window length when `-StartDate` is not given |
| `-TimeZone` | string | `Europe/Brussels` | IANA or Windows ID; works in Windows PowerShell 5.1 and PowerShell 7 |
| `-FollowCopies` | switch | off | Also follow copies. By default a copy is reported but not followed — the original stays where it was |
| `-IncludeActivity` | switch | off | Also opens, edits, downloads, sync and check-in/out: who last worked in it. Many more records, so slower |
| `-OutputPath` | string | `C:\Temp\FileTrail_<name>_<ts>.csv` | CSV path; the raw audit records of the trail go to the same name with `.json` |
| `-TenantId` | string | — | Tenant domain for `Connect-ExchangeOnline`; not needed when already connected |
| `-PassThru` | switch | off | Also return the timeline rows as objects |

**Examples**

```powershell
# Where did "Offerte Janssens.docx" go in the last 30 days?
.\Trace-SharePointFile.ps1 -Name "Offerte Janssens.docx" -TenantId contoso.onmicrosoft.com

# A period in Brussels time, one OneDrive only
.\Trace-SharePointFile.ps1 -Name "Budget*" -StartDate '01-09-2026' -EndDate '15-09-2026' `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jan_contoso_com

# From the link someone once sent, one afternoon
.\Trace-SharePointFile.ps1 -Url "https://contoso.sharepoint.com/sites/Sales/Shared Documents/2026/Prijslijst.xlsx" `
    -StartDate '2026-09-12 13:00' -EndDate '2026-09-12 18:00'
```

**Notes**

- Needs the **View-Only Audit Logs** or **Audit Logs** role in Exchange Online, and the module `ExchangeOnlineManagement`. Runs in Windows PowerShell 5.1 and PowerShell 7
- The audit log runs 30–90 minutes (occasionally 24 hours) behind. Audit Standard keeps **180 days**; the script warns when the window starts earlier
- Only what happened **inside** the window can be followed. A folder renamed before `-StartDate` is invisible; if the trail seems to start halfway, widen the window
- Renames by the OneDrive sync client (in Explorer) are audited like those in the browser; the `UserAgent` column tells them apart
- The last known location is what the audit log says — the script does not check that the file is still there
- CSV columns: `Item, Time, TimeUtc, Action, Operation, User, From, To, ViaFolder, ItemId, ClientIP, UserAgent, RecordId`. `ViaFolder` is filled when the step came from a folder action

---

### Revoke-SharePointUserAccess.ps1

The counterpart of [`Get-SharePointPermissionsReport.ps1`](../Reporting/readme.md#get-sharepointpermissionsreportps1): that script tells you who can get at what, this script takes it away. It finds every place where one named user has access and removes it:

- **Site collection administrator** — first, because that role overrides every role assignment below it; leaving it in place would make the rest cosmetic
- **Direct role assignments** on a site, subsite, list/library, folder or individual file
- **SharePoint groups** (Owners, Members, Visitors and custom groups)
- **Sharing links** — the `SharingLinks.*` groups a shared link puts its recipients in. That is how "anyone with the link" and "specific people" actually give a person access

Reporting is the default. Nothing changes without `-Apply`, and every run writes a CSV with exactly what was found and what was done with it.

#### What it deliberately does *not* do

| | Why |
|---|---|
| Change Entra ID group membership | **Unless `-RemoveFromEntraGroups` is given.** By default whoever gets in through a security or M365 group keeps that access — the group *is* the grant — and those routes are reported explicitly, with the group name, so you do not think it is closed while it is still open |
| Remove `Everyone` / `Everyone except external users` | That takes access away from the whole tenant, not from this person. Reported, not touched |
| Clean up ownership and metadata | A revoked user remains the author of what they created |

> **So offboarding is two steps.** Run this script, then deal with the Entra groups listed in the CSV under `Action = CannotRevoke`. Without that second step the access is not gone.

#### Robustness

This script removes permissions, so its failure modes differ from those of a report: silently hitting the wrong person, or not being able to tell afterwards what you removed.

| Situation | Behaviour |
|---|---|
| Identifying the user | **Exact comparisons only.** On UPN, e-mail, the claim suffix, and the decoded guest name (`jan_partner.com#ext#@tenant` becomes `jan@partner.com`). Never on a substring: `an@contoso.com` is contained in `jan@contoso.com`, and that is exactly how you revoke the wrong person |
| Two accounts with the same address | The site is **not** touched; the script stops with both login names in the error. Choosing is up to you, not the script |
| Audit CSV | Written row by row during the run, not at the end. A run that revokes two hundred things and then crashes must still be able to tell you *what* is gone |
| CSV is open in Excel | Five attempts with increasing wait time, then the run stops — better an aborted run than removing permissions without a trace |
| Removal returns `404` | `AlreadyGone`, not an error. On a second run that is the normal outcome; counted as an error, a clean run would look broken |
| `-WhatIf` | Same branch as a dry run, so `WouldRevoke` in the CSV — not `Skipped`, which would suggest someone declined a prompt |
| Site collection administrator cannot be removed | **Reported loudly, and the run counts it as failed.** That role reaches every scope in the site, so all other removals there are then cosmetic |
| Throttling (`429`/`503`) | Retry with `Retry-After`; a rejected token is fetched fresh once before the run stops |
| Unexpected error | A `trap` cleans up the temporary Full Control app before the script stops |

#### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-UserPrincipalName` | string | — | **Required.** The user, e.g. `jan@contoso.com`. For a guest the real address (`jan@partner.com`) works too — the script finds the `#ext#` variant itself |
| `-TenantUrl` | string | — | Tenant root. Required for a tenant-wide run |
| `-SiteUrl` | string | — | One site collection instead of the whole tenant |
| `-Apply` | switch | off | Actually revoke. Without it, report only |
| `-Scope` | `Site`/`List`/`Item` | `Item` | How deep to search for direct grants |
| `-IncludeGroupAccess` | switch | off | Also reports the sites the user reaches through Entra groups, *including* where they have nothing else. Report only |
| `-KeepSharingLinks` | switch | off | Leave sharing links alone; all other routes are still revoked |
| `-RemoveFromEntraGroups` | switch | off | **Also remove the user from the Entra ID groups that were seen granting access** — only those, never every group they belong to. Needs Graph `GroupMember.ReadWrite.All`, which the temporary app asks for only with this switch. See below |
| `-RemoveFromSite` | switch | off | Afterwards also removes the user from the user list of every site collection. Catches what the scope-by-scope pass missed, but the name then renders as a deleted account in older metadata |
| `-IncludeOneDriveSites` | switch | off | Also search personal OneDrive sites |
| `-IncludeHiddenLists` | switch | off | Also hidden and system lists |
| `-TenantId` / `-ClientId` / `-CertificateThumbprint` | string | — | Your own app registration instead of the temporary one. It needs SharePoint `Sites.FullControl.All` plus Graph `Sites.Read.All`, `User.Read.All` and `GroupMember.Read.All`. Without `User.Read.All` the user lookup comes back `403` and the run stops rather than mistaking that for a missing account |
| `-ClientSecret` | string | — | Works for Graph but **not** for SharePoint (see authentication under the report) |
| `-OutputPath` | string | `C:\Temp` | Output folder |
| `-GraphTimeoutSec` / `-MaxGraphRetry` | int | `120` / `6` | Timeout and retries |

Authentication is identical to the report: a short-lived, certificate-based app registration with SharePoint `Sites.FullControl.All`, which is deleted again afterwards.


#### Removing the Entra ID groups too

`-RemoveFromEntraGroups` completes the second half of an offboarding instead of only reporting it. **Only the groups this run actually caught holding a role assignment on a scope in range are touched** — never every group the user belongs to. A leaver can be in fifty groups; the ones that grant SharePoint access are the ones in scope.

> **This reaches past SharePoint.** An Entra group is not a SharePoint object. The same membership commonly carries a Teams team, a mailbox, licences and app assignments, none of which this report can see. Read the report from a run without `-Apply`, then re-run with the switch.

Four cases are reported rather than forced, because forcing them would either fail or do the wrong thing:

| Case | Why it is left alone |
|---|---|
| Dynamic group | Membership follows a rule and is not stored, so there is nothing to remove. Change the rule, or the user attributes it matches |
| Synced from on-premises AD | Read-only in the cloud. The membership has to go in Active Directory |
| Member through a nested group | The user is not a direct member, so removing them here would fail. The access has to be cut at the group that actually holds them, which the report names |
| User not resolved in Entra | Nothing to remove them from; the SharePoint side still runs |

The removal goes through the same funnel as every other change, so `-Apply`, `-WhatIf`, the confirmation prompt and the audit CSV all behave identically. Needs Graph `GroupMember.ReadWrite.All`, which the temporary app asks for **only** when the switch is given — a report-only run holds no permission that can change group membership.

#### Output

`SharePoint_Revoke_<user>_<ts>.csv`, one row per grant found, with an `Action` column:

| Action | Meaning |
|---|---|
| `WouldRevoke` | Found, and would be removed — this is what you get without `-Apply` |
| `Revoked` | Removed |
| `Failed` | Attempt failed; the reason is in `Detail` |
| `AlreadyGone` | Nothing left to remove — the normal outcome on a second run |
| `CannotRevoke` | Through an Entra group or `Everyone` — has to be resolved elsewhere |
| `Kept` | Deliberately left in place by `-KeepSharingLinks` |
| `Skipped` | Declined at the confirmation prompt |

#### Examples

```powershell
# What can Jan reach? Changes nothing
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com"

# The same, and now actually revoke
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -Apply

# Remove a guest from one site collection, sharing links included
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName gast@partner.com -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

# Offboarding checklist: also the Entra groups that give access
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -IncludeGroupAccess
```

> Running unattended? Pass `-Confirm:$false`, otherwise the script asks for confirmation for every removal (`ConfirmImpact = 'High'`).

---

### Test-SharePointAccessScripts.ps1

Checks `Revoke-SharePointUserAccess.ps1` and `Get-SharePointPermissionsReport.ps1` without touching a tenant. Run it after every change to either of them; exit code 0 means both are in order.

Two things are checked, both of them errors that go wrong silently in production:

1. **The shared authentication block is byte-identical.** Both scripts contain the same app-only auth and SharePoint REST layer, delimited by `SHARED BLOCK START/END`. That layer took four live runs against a tenant to get right — certificate instead of secret, tokens that have to prove their app roles before they are cached, 401 as fatal instead of per site, paging that cannot get stuck. A second copy that silently drifts is a correctness risk in precisely the script that removes permissions. On a difference, the first line that differs is shown.
2. **The revocation funnel behaves.** A dry run must record its intent and execute nothing, `-Apply` must execute *and* record, a failure must end up in the audit trail instead of disappearing, and the grants the script must refuse to remove (through an Entra group, or to everyone) must stay refused.

```powershell
.\Test-SharePointAccessScripts.ps1
```
