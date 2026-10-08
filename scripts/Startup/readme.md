**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **Startup**

# Startup

Entry-point scripts and the core M365 function library.

---

## Files

| File | Description |
|------|-------------|
| [`functies.ps1`](functies.ps1) ([docs](#functiesps1)) | M365 function library — dot-sourced by `menu.ps1` on first use |
| [`RequiredModules.psd1`](RequiredModules.psd1) ([docs](#requiredmodulespsd1)) | The one list of modules this repo needs — read by `load.ps1`, `Install-Modules.ps1` and `Update-Modules.ps1` |
| [`Test-RequiredModules.ps1`](Test-RequiredModules.ps1) ([docs](#test-requiredmodulesps1)) | Reports modules that scripts load but `RequiredModules.psd1` does not list — run by the docs hook after every edit |
| [`Connect-M365.ps1`](Connect-M365.ps1) ([docs](#connect-m365ps1)) | The one way scripts sign in: Graph first, delegated by default (device code and GDAP from `load.config.ps1`), app-only on request — dot-sourced by the scripts |
| [`Install-Modules.ps1`](Install-Modules.ps1) ([docs](#install-modulesps1)) | Bootstrap script — installs and imports all required PowerShell modules |
| [`Update-Modules.ps1`](Update-Modules.ps1) ([docs](#update-modulesps1)) | Checks the required modules (missing, too old, update available) and installs/updates them; optionally updates every other installed module too |
| [`Test-PowerShellSyntax.ps1`](Test-PowerShellSyntax.ps1) ([docs](#test-powershellsyntaxps1)) | Parse-checks `.ps1` files in the repo for syntax errors, no execution |
| [`Update-ScriptIndex.ps1`](Update-ScriptIndex.ps1) ([docs](#update-scriptindexps1)) | Regenerates [`scripts/INDEX.md`](../INDEX.md) — the searchable A–Z list of every script |
| [`Test-MarkdownLinks.ps1`](Test-MarkdownLinks.ps1) ([docs](#test-markdownlinksps1)) | Checks every link in every readme — files that must exist, anchors that must match a heading |
| [`Convert-MarkdownToHtml.ps1`](Convert-MarkdownToHtml.ps1) ([docs](#convert-markdowntohtmlps1)) | Builds a self-contained, styled HTML page from a markdown document — for pasting into IT Glue or printing |
| [`Update-ReadmeHeader.ps1`](Update-ReadmeHeader.ps1) ([docs](#update-readmeheaderps1)) | Writes the language switcher and breadcrumb at the top of every readme, in English, Dutch and French |

---

## Delegated GDAP Startup

`load.ps1` now supports storing delegated defaults in `load.config.ps1`:

- `authMode` (`GDAP` or `Direct`)
- `defaultCustomerDomain` (optional)
- `useDeviceCodeAuth` (`$true` / `$false`)

When `authMode` is `GDAP` and a `defaultCustomerDomain` is configured, the menu auto-runs `Connect-Tenant` after `functies.ps1` is loaded.

To register the launcher at Windows sign-in:

```powershell
.\load.ps1 -SetupStartup
```

To remove it:

```powershell
.\load.ps1 -RemoveStartup
```

At every start, before the menu opens, `load.ps1` checks the modules in [`RequiredModules.psd1`](#requiredmodulespsd1) with [`Update-Modules.ps1`](#update-modulesps1): it lists what is missing, older than its minimum, or behind the PowerShell Gallery, and installs or updates it right away, without asking. The gallery is asked at most once every 24 hours, so a normal start costs under a second. Skip the check once with:

```powershell
.\load.ps1 -SkipModuleCheck
```

You can also toggle startup from the launcher menu:

- `F` = Enable-LauncherStartup
- `G` = Disable-LauncherStartup

---

## functies.ps1

Central function library for multi-tenant M365 management via Microsoft Graph and Exchange Online. Loaded automatically by the menu on first use of a B–E option.

### Configuration

Edit the `#region Configuration` block at the top to match your organisation:

```powershell
$script:MspAdminAlias       = 'msp-admin'
$script:MspAdminDisplayName = 'MSP - Admin Account'
```

### Select a customer tenant

```powershell
Connect-Tenant -Domain "customer.com"
# Sets $global:cid and $global:connectmsoldomain
# Under GDAP, all subsequent functions target the selected tenant
```

### Sign-in

Every function signs in through [`Connect-M365.ps1`](#connect-m365ps1), which `functies.ps1` dot-sources: Microsoft Graph, delegated, with a device code when `useDeviceCodeAuth` is set in `load.config.ps1`, otherwise in the browser. A session that already holds the scopes a function needs is reused.

- **GDAP** (`authMode = 'GDAP'`): after `Connect-Tenant`, every Graph and Exchange function connects to that customer (`$cid`, Exchange through `-DelegatedOrganization`).
- **Direct**: the functions stay in your own tenant; `$cid` is not used.
- The partner session that `Connect-Tenant` and `Test-GdapConnection` need for `Get-MgContract` is opened by `Connect-PartnerGraph`, which remembers your own tenant (`$global:partnerTenantId`) at the first start, so selecting another customer still works after a function has switched to the current one.
- Teams (`Invoke-Menu` option 3) goes through `Connect-M365Teams`, Exchange through `Connect-M365Exchange`.

### Functions

**Connection**

| Function | Description |
|----------|-------------|
| `Connect-Tenant` | Select a CSP customer by domain, populates `$cid` and `$connectmsoldomain` |
| `Test-GdapConnection` | Validates the delegated GDAP/CSP contract, then tries a delegated Graph connection to the customer (`Get-MgOrganization`) and a delegated Exchange connection |
| `Test-ExoConnection` | Connects to Exchange Online, or reuses a session to the right organisation |
| `Connect-PartnerGraph` | Graph in your own (partner) tenant, for `Get-MgContract` |

**Exchange Online**

| Function | Description |
|----------|-------------|
| `Enable-CopyOfSentItems` | Enables copy of sent items for all mailboxes |
| `Add-SharedMailboxAccess` | Grants FullAccess + SendAs on a shared mailbox |
| `Set-MailboxLocale` | Sets language and timezone on all mailboxes (default: NL / W. Europe) |
| `Add-MailboxAlias` | Adds an alias to a mailbox |
| `Get-MailboxAliases` | Lists all SMTP aliases per mailbox |
| `Export-DistributionGroups` | Exports all distribution groups to CSV (`C:\Temp\`) |
| `Set-AutoReply` | Configures an out-of-office reply |

**Entra ID / Graph**

| Function | Description |
|----------|-------------|
| `Get-TenantAdmins` | Lists all Global Administrators |
| `Add-TenantDomain` | Adds a domain and walks through verification |
| `Get-TenantLicenses` | Shows licence overview with usage and availability |
| `Get-TenantUsers` | Lists all users with UPN, display name, and licences |
| `Add-TenantAdmin` | Grants Global Administrator rights to a user |
| `Get-EntraApplication` | Finds an Enterprise App by name |
| `Reset-UserPassword` | Resets a user's password |
| `Export-SignInLogs` | Exports sign-in logs to CSV in `C:\Temp\` (default: last 30 days) |

**MSP Admin Account**

| Function | Description |
|----------|-------------|
| `New-MspAdmin` | Creates the MSP admin account as Global Admin in the customer tenant |
| `Set-MspAdminAsGroupOwner` | Sets the MSP admin account as group owner |
| `Reset-MspAdminPassword` | Resets the MSP admin account password |

---

## RequiredModules.psd1

The modules this repository depends on, in one PowerShell data file. `load.ps1`,
`Install-Modules.ps1` and `Update-Modules.ps1` all read it, so they can no longer disagree
— before, each had its own list, and `load.ps1` only checked seven modules.

**To add a module:** add a line here. The next start of `load.ps1` on every machine sees it
as missing and installs it. Raising a `MinimumVersion` works the same way. Forgetting is
hard: [`Test-RequiredModules.ps1`](#test-requiredmodulesps1) runs after every edit and reports
a module a script loads that is not in this file.

| Key | Meaning |
|-----|---------|
| `Name` | Module name on the PowerShell Gallery |
| `MinimumVersion` | Older than this counts as *too old* (not just *update available*) |
| `WindowsOnly` | Skipped on macOS and Linux |
| `MinimumPSVersion` | Skipped on an older PowerShell — `PnP.PowerShell` 3 needs 7.4 |
| `ImportAtStartup` | Imported by `load.ps1` before the menu opens |

Next to `Modules` there is `NotManaged`: modules scripts load that are deliberately not installed from the gallery, each with its reason — `ActiveDirectory` and `WebAdministration` (Windows features), `AzureAD` (retired), `Microsoft.Graph` (the whole SDK, only named in install hints).

Current list: `ExchangeOnlineManagement`, the Graph submodules `Authentication`, `Sites`,
`Identity.DirectoryManagement`, `Identity.SignIns`, `Identity.Governance`, `Applications`,
`Calendar`, `Groups`, `Users`, `Reports`, plus `PnP.PowerShell`, `MicrosoftTeams`, `ImportExcel`,
`Az.Accounts`, `Az.OperationalInsights`, `DCToolbox`, `IntuneBackupAndRestore`, and on Windows
`WindowsAutopilotIntune` and `IntuneWin32App`.

---

## Test-RequiredModules.ps1

A script that starts using a new module works on the machine it was written on and fails
everywhere else until the module is in [`RequiredModules.psd1`](#requiredmodulespsd1). This
script searches every `.ps1`/`.psm1` for `#Requires -Modules`, `Import-Module` and
`Install-Module` with a literal name, and reports each name that is in neither `Modules` nor
`NotManaged`, with the files that use it. A name in a variable (`$mod`) cannot be checked.

The docs hook (`.claude/hooks/sync-docs.ps1`) runs it after every edit of a `.ps1`, `.psd1` or
`.md`, and the git pre-commit hook warns with it, so a new module is caught when the script is
saved, not when someone else's run fails.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Root` | Repository root (default: two levels above this script) |

**Examples**

```powershell
pwsh -File scripts/Startup/Test-RequiredModules.ps1
```

Exit codes: `0` = every module a script loads is listed, `1` = something is missing.

---

## Connect-M365.ps1

The sign-in every script uses. **Microsoft Graph is the standard**; Exchange Online, Teams
and PnP are only connected for work Graph has no API for (mailbox and SendAs permissions,
message trace, DKIM, EOP policies, Teams `Cs*` policies, SharePoint role assignments, ...).

```powershell
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')   # depth depends on the script's folder
$graph = Connect-M365Graph -Scopes 'User.Read.All' -TenantId $TenantId
# ... work ...
Disconnect-M365Graph $graph    # disconnects only what this call connected
```

| Function | What it does |
|----------|--------------|
| `Connect-M365Graph` | Microsoft Graph. `-Scopes`, `-TenantId`, `-ClientId` + `-CertificateThumbprint`/`-ClientSecret`, `-AppOnly`, `-DeviceCode`, `-Interactive`, `-Force`; `-Force` signs in again even when the session would fit. Safe under `Set-StrictMode`, also without `load.ps1` |
| `Disconnect-M365Graph` | Disconnects only when `Connect-M365Graph` opened the session |
| `Connect-M365Exchange` | Exchange Online, `-IncludeCompliance` adds Security & Compliance (`Connect-IPPSSession`), `-EnableSearchOnlySession` for Content Search |
| `Disconnect-M365Exchange` | Closes only the sessions `Connect-M365Exchange` opened (by connection id), never the caller's |
| `Connect-M365Teams` / `Disconnect-M365Teams` | Microsoft Teams PowerShell, for the `Cs*` policies |
| `Connect-M365PnP` | PnP.PowerShell to a site; returns the connection. ClientId from `-ClientId` or `pnp.appid.json`; `-AppOnly` uses the certificate app from `graph.appid.json` |
| `Invoke-M365GraphPaged` | GET a Graph collection and follow `@odata.nextLink` to the end |
| `Resolve-M365TenantId` | The tenant to use: `-TenantId`, else the GDAP customer, else your own tenant |

**How it signs in**

- **Delegated, the default.** You sign in as yourself, with a device code when
  `useDeviceCodeAuth` is set in `load.config.ps1` (or `-DeviceCode` is passed), otherwise in
  the browser with your `upn` pre-filled. Under GDAP (`authMode = 'GDAP'`) the customer
  tenant is `$global:cid` / `$global:connectmsoldomain` from `Connect-Tenant`, or
  `$env:M365_CUSTOMER_TENANTID`. Exchange reaches the customer with `-DelegatedOrganization`;
  `-Organization` only works for app-only sign-in. Outside GDAP a delegated Exchange sign-in
  lands in the tenant of the account you sign in with.
- **App-only, on request.** `-ClientId` with `-CertificateThumbprint` (or `-ClientSecret`,
  Graph only), or `-AppOnly` to read ClientId and thumbprint for the tenant from
  `graph.appid.json` in the repo root (gitignored). The app must be consented in that
  tenant: GDAP gives delegated rights, not app-only access.
- **Existing sessions are reused** when they are the right kind, for the right tenant and
  (delegated) already hold every requested scope. A delegated reconnect keeps the scopes
  the earlier session had, so a second script in the same window does not take them away.

**Notes**

- Requires PowerShell 7. Each `Connect-*` throws with an install hint when its module is missing.
- Run from the repo: scripts dot-source this file by relative path, so a script copied on
  its own needs this file next to it.

---

## Install-Modules.ps1

Installs and imports every module in [`RequiredModules.psd1`](#requiredmodulespsd1). Run once on a new machine or after a clean PowerShell install. Modules that are already installed are left alone — updating is [`Update-Modules.ps1`](#update-modulesps1)'s job.

```powershell
.\scripts\Startup\Install-Modules.ps1
```

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Force` | Reinstall modules even if already present |
| `-Scope` | `CurrentUser` (default) or `AllUsers` (elevated) |

---

## Update-Modules.ps1

Checks every module in [`RequiredModules.psd1`](#requiredmodulespsd1) and gives each a status:

| Status | Meaning | Without `-CheckOnly` |
|--------|---------|----------------------|
| `Missing` | Not installed | Installed |
| `BelowMinimum` | Older than its `MinimumVersion` | Updated |
| `UpdateAvailable` | A newer version is on the PowerShell Gallery | Updated |
| `OK` | Current | — |
| `Unknown` | Gallery unreachable, or the module is no longer on it | — |
| `Skipped` | Not for this platform or PowerShell version | — |

Then, unless `-RequiredOnly`, it updates every other module installed through PowerShellGet, as it always did. `load.ps1` runs it at startup as `-RequiredOnly -Auto -MaxAgeHours 24`.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-CheckOnly` | Report only; install or update nothing |
| `-RequiredOnly` | Only the modules in `RequiredModules.psd1`, not everything else installed |
| `-MaxAgeHours` | Reuse gallery versions cached within this many hours (default `0` = always ask the gallery) |
| `-Scope` | Scope for newly installed modules: `CurrentUser` (default) or `AllUsers` |
| `-Quiet` | No line per module, only errors |
| `-PassThru` | Return a status object per required module (`Name`, `Installed`, `Minimum`, `Latest`, `Status`, `Reason`) |
| `-Auto` | For startup: install what is missing and update what is outdated without asking, showing only what it does. All in order gives one line, `Modules OK` |
| `-Prompt` | As `-Auto`, but lists what it would do and asks first |

**Examples**

```powershell
# What is missing or outdated? Changes nothing
.\scripts\Startup\Update-Modules.ps1 -RequiredOnly -CheckOnly

# Install what is missing and update what is behind, only the repo's modules
.\scripts\Startup\Update-Modules.ps1 -RequiredOnly

# The above, then update every other installed module as well
.\scripts\Startup\Update-Modules.ps1
```

**In your PowerShell profile.** If you start PowerShell with your own profile (`$PROFILE`) instead of `load.ps1`, add the line `load.ps1` uses, so the check runs at every PowerShell start:

```powershell
& "C:\path\to\M365-Scripts\scripts\Startup\Update-Modules.ps1" -RequiredOnly -Auto -MaxAgeHours 24
```

With everything current it prints one line and costs well under a second; the gallery is asked at most once a day.

**Notes**

- The installed version is read with `Get-Module -ListAvailable`, so a module that was not installed through PowerShellGet (a manual copy, an MSI) still counts as installed. Such a module gets the new version side by side through `Install-Module`, because `Update-Module` refuses modules it did not install.
- Gallery versions come from `Find-PSResource` when PSResourceGet is present (about 3 s for the whole list), otherwise `Find-Module` (about 9 s). They are cached in `%LOCALAPPDATA%\M365-Scripts\module-gallery-cache.json`; `-MaxAgeHours` decides how old that cache may be.
- Offline, the gallery lookup fails quietly: missing and too-old modules are still reported, *update available* is not.
- PowerShell 7 and Windows PowerShell 5.1 have separate module folders, so the two can report different results on the same machine. That is correct, not a bug.
- Old versions are not removed. Run as administrator to update modules installed for all users.

---

## Test-PowerShellSyntax.ps1

Parse-checks `.ps1` (and optionally `.psm1`) files for syntax errors without executing them — uses `[System.Management.Automation.Language.Parser]::ParseFile()`.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Path` | No | File or folder to check (default: repo root) |
| `-Recurse` | No | Recurse into subfolders when `-Path` is a folder |
| `-IncludePsm1` | No | Also check `.psm1` module files |

**Examples**

```powershell
# Check the whole repo
.\Test-PowerShellSyntax.ps1 -Recurse

# Check a single file
.\Test-PowerShellSyntax.ps1 -Path .\scripts\Entra\New-M365User.ps1
```

Exit codes: `0` = no errors, `1` = syntax errors found, `2` = path/argument error.

---

## Update-ScriptIndex.ps1

Builds [`scripts/INDEX.md`](../INDEX.md): one page listing every script in the repository
A–Z, with a link to the file, a link to its folder readme, and a one-line description.

It exists because finding a script on GitHub otherwise means guessing which workload
folder it is under and opening readmes until it turns up. One generated page is
Ctrl-F-able and clickable, and — being generated — cannot drift from the files the way a
hand-kept table does.

**Where the description comes from**

| Order | Source |
|-------|--------|
| 1 | The script's `.SYNOPSIS` block, joined across the lines it wraps over |
| 2 | Failing that, the first real line of a leading `#` comment block |
| 3 | Failing that, nothing — and the script is listed under *Scripts without a description*, so the gap is visible instead of silently blank |

A single `#` comment sitting directly on top of code is deliberately **not** used: a line
like `# URL van de theme` above a `$ThemeUrl` assignment describes that variable, not the
script, and reading it as a description puts something worse than nothing in the table.
A header block runs to several lines, or is set off from the code by a blank line.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Root` | Repository root (default: two levels above this script) |
| `-Check` | Write nothing; exit `1` when the committed page no longer matches the scripts on disk |
| `-WhatIf` | Report what would change without writing |

**Examples**

```powershell
# Rebuild the index after adding, renaming or removing a script
pwsh -File scripts/Startup/Update-ScriptIndex.ps1

# Fail when the index is stale — for a hook or a pipeline
pwsh -File scripts/Startup/Update-ScriptIndex.ps1 -Check
```

> Rerun it whenever a script is added, renamed, moved or removed — the same moment the
> [working rules](../../.claude/CLAUDE.md) already ask you to update the readmes and
> `menu.ps1`. It rewrites nothing when the page is already current, so it is safe to run
> on every commit.

Exit codes: `0` = written or already current, `1` = `-Check` found the page stale.

---

## Convert-MarkdownToHtml.ps1

The service desk documents in this repo are markdown, but IT Glue and most ticket systems want rich text. Converting by hand means the HTML is stale the first time the markdown changes — and `Update-TeamsClient-ITGlue.md` changed six times in two days — so the page is generated instead.

Supported, because it is what these documents use: headings, tables with a header row, fenced code blocks (including the indented ones inside numbered steps), blockquotes, ordered and unordered lists, horizontal rules, and inline code, bold, italic and links. Anything else passes through as text rather than being guessed at.

The CSS is embedded, so the page is standalone — nothing to host, and nothing to break when the file is copied elsewhere. Void elements are written self-closing (`<hr/>`, `<br/>`), so the output parses as XML as well as HTML and can be checked structurally instead of by eye.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Path` | The markdown file to convert (required) |
| `-Destination` | Where to write the HTML (default: same folder and name, `.html`) |
| `-Title` | Page and browser title (default: the document's first `#` heading) |
| `-Check` | Write nothing; exit `1` when the HTML on disk no longer matches the markdown |
| `-WhatIf` | Show what would be written and change nothing |

**Examples**

```powershell
# Build the IT Glue page next to its markdown source
pwsh -File scripts/Startup/Convert-MarkdownToHtml.ps1 -Path scripts/Device/Update-TeamsClient-ITGlue.md

# Has the committed page fallen behind? Exit code 1 if it has
pwsh -File scripts/Startup/Convert-MarkdownToHtml.ps1 -Path scripts/Device/Update-TeamsClient-ITGlue.md -Check
```

From the menu: `menu.ps1`, key **M**. It offers the Teams IT Glue document as the default path and asks whether to only check.

**Notes**

- **Getting it into IT Glue:** open the `.html` in a browser, select all, copy, and paste into the IT Glue document editor. The editor keeps the headings, tables and code blocks and drops the CSS — which is what you want there, since IT Glue applies its own.
- The generated-on line is excluded from the `-Check` comparison, so rerunning it on an unchanged document does not report a difference.
- Rerun it after editing the markdown. `-Check` is what a pre-commit hook or a pipeline would call.

Exit codes: `0` = written or already current, `1` = `-Check` found the page stale, or the page does not exist yet.

---

## Test-MarkdownLinks.ps1

Walks every `.md` file in the repository and reports links that go nowhere. A dead
link in a readme is invisible until someone clicks it, which is usually the moment
they needed it.

It checks two kinds:

| Kind | What can go wrong |
|------|-------------------|
| A link to a file or folder | The file was renamed, moved or removed and the readme still points at the old path. Percent-encoded spaces (`Time%20sync/readme.md`) are decoded before the path is tested, because that is what GitHub serves |
| An in-page anchor (`#set-usermanagerps1`) | The heading it points at was renamed, or the anchor was typed by hand and never matched. These rot in silence — nothing warns you |

Anchors are resolved the way GitHub builds them: the heading is lower-cased, markdown
formatting is stripped, everything that is not a letter, digit, space, `_` or `-` is
dropped, and spaces become hyphens — so `### Watch-RDSLive.ps1` is `#watch-rdsliveps1`,
not `#watch-rdslivesps1`. Repeated headings get GitHub's `-1`, `-2` suffix. Invisible
characters (variation selectors, zero-width joiners — the bytes that make an emoji an
emoji) are stripped from both the heading and the link before they are compared, so an
emoji heading in the table of contents does not read as broken.

Fenced code blocks and inline code are skipped, so a readme that *documents* link
syntax does not report itself as broken — the `([docs](#…))` example two paragraphs up
is text about links, not a link.

External links (`http`, `https`, `mailto`) are counted but not fetched: this is a
structural check, and it has to work offline.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Root` | Repository root (default: two levels above this script) |
| `-Path` | Check one file or folder instead of the whole repository |

**Examples**

```powershell
# Check every readme in the repository
pwsh -File scripts/Startup/Test-MarkdownLinks.ps1

# Only one workload folder
pwsh -File scripts/Startup/Test-MarkdownLinks.ps1 -Path scripts/Exchange
```

Exit codes: `0` = every internal link resolves, `1` = something is broken (each one
listed with the file it is in and why it failed).

---

## Update-ReadmeHeader.ps1

Every folder has its readme three times: `readme.md` (English), `readme.nl.md` (Dutch) and
`readme.fr.md` (French). Each opens with the same two lines, different only in their paths:
a language switcher to the same page in the other languages, and a breadcrumb back up the
tree in which every level links to its own readme **in the current language**.

Kept by hand, those paths are exactly what goes wrong — one `../` too few after a folder
moves, or a Dutch page that links to the English parent. So they are generated from the
folder the readme sits in, and nothing else. Everything above the first heading that is a
switcher or breadcrumb line is replaced; the rest of the file is not touched.

A folder with a `readme.md` but no Dutch or French version is reported: its switcher would
link to nothing.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Root` | Repository root (default: two levels above this script) |
| `-Check` | Write nothing; exit `1` when a header is out of date or a language version is missing |
| `-WhatIf` | Report which headers would be rewritten without writing |

**Examples**

```powershell
# After adding or moving a folder readme (write the .nl.md and .fr.md versions first)
pwsh -File scripts/Startup/Update-ReadmeHeader.ps1

# Fail when a header is stale or a translation is missing — for a hook or a pipeline
pwsh -File scripts/Startup/Update-ReadmeHeader.ps1 -Check
```

> Run [`Test-MarkdownLinks.ps1`](#test-markdownlinksps1) afterwards: this script writes the
> links, that one proves they resolve.
