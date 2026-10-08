**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **Teams**

# Teams

Microsoft Teams / SharePoint export and archiving tooling.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Invoke-TeamsArchive.ps1`](Invoke-TeamsArchive.ps1) ([docs](#invoke-teamsarchiveps1)) | Runs a Teams/SharePoint export and archiving flow for a list of teams and channels read from Excel |

---

### Invoke-TeamsArchive.ps1

Teams archiver: exports members, files and chat of the listed channels, then (optionally) archives teams or channels. Everything runs on Microsoft Graph. PowerShell 7+ required, run as Global Admin of the customer (or with GDAP rights on it).

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Step10Action` | `interactive` (default), `archive`, `undo`, or `skip` |
| `-Step10Only` | Run only Step 10 (archive/unarchive), non-interactively |
| `-ChannelAction` | `none` (default), `archive`, or `undo` — per-channel quick mode |
| `-ChannelArchiveTag` | Marker text used for the rename fallback (default: `[ARCHIEF]`) |
| `-ChannelFallbackToRename` | Fall back to a rename marker if the Graph archive/unarchive API call fails |
| `-DryRun` | Simulate — keeps the full sign-in and validates Steps 6-9 by probing counts, without writing exports or mutating archive state |
| `-WorksheetName` | Worksheet holding the team list. Default: the first worksheet of the Excel file |
| `-TenantId` | Customer tenant ID. Default offered in the wizard: the GDAP customer from `Connect-Tenant` |
| `-ClientId` / `-CertificateThumbprint` | App-only instead of delegated, with this app registration and certificate |
| `-AppOnly` | App-only with the app registration for the tenant in `graph.appid.json` |
| `-PnPClientId` | PnP app for the site-admin fallback. Default: the tenant's entry in `pnp.appid.json` |

**Examples**

```powershell
# Delegated (default): device code per load.config.ps1, dry run first
.\Invoke-TeamsArchive.ps1 -DryRun

# Archive the listed channels only, app-only
.\Invoke-TeamsArchive.ps1 -Step10Only -ChannelAction archive -AppOnly -TenantId <tenant-guid>
```

**Notes**

The Excel file needs the columns `TeamName`, `ChannelName` and `Archive` (rows with `Archive` = `Archive` are processed). Nothing in the script is tied to one customer: the tenant ID, SharePoint URL, Excel file and archive folder are asked by the setup wizard (defaults `C:\Temp\Teams_Channels.xlsx` and `C:\Temp\Teams_Archive`).

Sign-in (v9.0):
- Goes through [`Connect-M365.ps1`](../Startup/readme.md#connect-m365ps1). **Delegated by default**: you sign in as the admin, with a device code when `useDeviceCodeAuth` is set in `load.config.ps1` (and, without `load.config.ps1`, a device code as before). The clean-session restart carries those settings and the GDAP customer over to the new session.
- **No temporary app registration any more.** Earlier versions created one every run, consented it and removed it again; the delegated scopes are now requested on the Microsoft Graph Command Line Tools app at sign-in.
- **App-only** with `-ClientId` + `-CertificateThumbprint`, or `-AppOnly`. The app then needs the application permissions `Group.Read.All`, `Sites.Read.All`, `TeamMember.Read.All`, `ChannelMessage.Read.All` (a protected API Microsoft must approve) and `TeamSettings.ReadWrite.All` / `ChannelSettings.ReadWrite.All` for archiving.
- The Graph session is closed at the end; it is the restarted session's own.

What runs where:
- **Graph** for everything that can: finding the teams (`/groups`, filtered on `resourceProvisioningOptions` = `Team`), members (`/teams/{id}/members`), channels (`/teams/{id}/channels`), the channel's files location (`filesFolder`), the file listing and download (`/drives/{id}/items/{id}/children` and `/content`), chat (`/messages`, `/replies`) and archiving (`/teams/{id}/archive`, `/channels/{id}/archive`). The MicrosoftTeams module is no longer used or installed.
- **PnP** only for the site-admin fallback: when a channel's files answer "access denied" to a delegated sign-in, the signed-in admin is made site collection admin once per site (`Set-PnPSite -Owners`) — Graph has no API for that. It uses the PnP app from `pnp.appid.json` (or `-PnPClientId`); without one it says how to do it by hand. App-only never needs it.

Behaviour:
- Resolves channel file locations via Graph `filesFolder` for all channel types (standard/private/shared), with the channel list cached per team.
- Normalizes TeamName/ChannelName values from Excel (trim) and matches channel names normalized (trim + whitespace collapse + lowercase) to avoid false "Kanaal niet gevonden" cases.
- Handles SharePoint NotFound during export as a controlled skip.
- Downloads files with per-file retries and a post-download count check. The folder structure inside the channel is now kept under `Files` (it used to be flattened, so two files with the same name in different subfolders overwrote each other and failed the count check).
- Stores output per channel: `Teams > Team > Channel > Files, Chat, Members`. `members.csv` keeps the columns `Name`, `User`, `Role`.
- Does not archive Teams by default; archiving requires explicit confirmation during Step 10, which also supports `unarchive`, a non-interactive quick mode (`-Step10Only -Step10Action undo|archive|skip`) and real per-channel archive/unarchive (`-Step10Only -ChannelAction archive|undo`), with an optional rename-marker fallback (`-ChannelFallbackToRename`, `-ChannelArchiveTag`).
- Archive/unarchive of a team is a team-level action in Microsoft Teams.
- Dry-run keeps the full sign-in, validates Steps 6-9 by probing counts without writing to disk, uses those counts in the Step 11 report, and only simulates Step 10 (`[DRYRUN]`).
- The clean-session restart after the module cleanup passes every parameter on (including `-DryRun`), returns the restarted run's exit code, and no longer leaves its marker variable behind in the calling session (a second run in the same session used to skip the cleanup).
