# SharePoint Scripts

SharePoint Online and OneDrive content operations via PnP PowerShell, with interactive
admin sign-in on any (customer) tenant.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Find-SiteContent.ps1`](#find-sitecontentps1) | Search a whole site (name, path, type, size, date or full text) and report the permissions on every hit |
| [`Restore-RecycleBinItems.ps1`](#restore-recyclebinitemsps1) | Restore deleted files/folders from a site or OneDrive recycle bin (dry-run by default) |

---

### Find-SiteContent.ps1

Answers the two questions you normally have at the same time: **where does this live**
and **who can get at it**. Read-only — the script never changes anything.

**Two engines**

| Engine | When | What it sees |
|--------|------|--------------|
| Crawl (default) | No `-Content` given | Walks every list and library item by item. Sees everything, also what the search index has not picked up yet. Slower on big sites |
| Search (`-Content`) | Full-text query | A KQL query against the search index, scoped to the site path — this is the one that matches text *inside* documents. Fast, but limited to what is indexed and what the signed-in account may see. Covers subsites automatically |

Both engines feed the same filters: `-Name` (wildcards), `-Path`, `-Extension`,
`-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter` / `-ModifiedBefore`,
`-MinSizeMB`. Hidden and system libraries are skipped unless you pass `-IncludeHidden`.
While crawling, only the top web is searched unless you add `-IncludeSubsites` — the
script reports how many subsites it skipped.

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
| `-Name` | No | Filter on item/file name, wildcards allowed (`*offerte*`) |
| `-Path` | No | Filter on the folder, substring match on the server relative URL |
| `-Extension` | No | One or more extensions, with or without the dot (`xlsx`,`pdf`) |
| `-ItemType` | No | `All` (default), `File`, `Folder` or `ListItem` |
| `-ListName` | No | Only these lists/libraries by title, wildcards allowed |
| `-ModifiedBy` | No | Who last changed it — display name or e-mail, wildcards allowed |
| `-ModifiedAfter` / `-ModifiedBefore` | No | Restrict to a change window |
| `-MinSizeMB` | No | Only files of at least this size |
| `-IncludeHidden` | No | Also search hidden lists, catalogs and system libraries |
| `-Permissions` | No | `Effective` (default), `Unique` (only broken inheritance) or `None` |
| `-ExpandGroups` | No | Also list the members of regular SharePoint groups (link groups are always expanded) |
| `-IncludeLimitedAccess` | No | Keep `Limited Access` assignments in the report |
| `-MaxItems` | No | Stop after this many matches (default 5000) |
| `-MaxPermissionLookups` | No | Cap on hits that get their permissions resolved (default 1000) |
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
