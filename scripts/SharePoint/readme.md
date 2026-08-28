# SharePoint Scripts

SharePoint Online and OneDrive content operations via PnP PowerShell, with interactive
admin sign-in on any (customer) tenant.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Restore-RecycleBinItems.ps1`](#restore-recyclebinitemsps1) | Restore deleted files/folders from a site or OneDrive recycle bin (dry-run by default) |

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
