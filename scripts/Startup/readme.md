# Startup

Entry-point scripts and the core M365 function library.

---

## Files

| File | Description |
|------|-------------|
| [`functies.ps1`](functies.ps1) | M365 function library — dot-sourced by `menu.ps1` on first use |
| [`Install-Modules.ps1`](Install-Modules.ps1) | Bootstrap script — installs and imports all required PowerShell modules |
| [`Update-Modules.ps1`](Update-Modules.ps1) | Updates every installed PowerShell module to its latest version |
| [`Test-PowerShellSyntax.ps1`](Test-PowerShellSyntax.ps1) | Parse-checks `.ps1` files in the repo for syntax errors, no execution |
| [`Update-ScriptIndex.ps1`](Update-ScriptIndex.ps1) | Regenerates [`scripts/INDEX.md`](../INDEX.md) — the searchable A–Z list of every script |
| [`Test-MarkdownLinks.ps1`](Test-MarkdownLinks.ps1) | Checks every link in every readme — files that must exist, anchors that must match a heading |
| [`Convert-MarkdownToHtml.ps1`](Convert-MarkdownToHtml.ps1) | Builds a self-contained, styled HTML page from a markdown document — for pasting into IT Glue or printing |

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
# All subsequent functions automatically target the selected tenant
```

### Functions

**Connection**

| Function | Description |
|----------|-------------|
| `Connect-Tenant` | Select a CSP customer by domain, populates `$cid` and `$connectmsoldomain` |
| `Test-GdapConnection` | Validates delegated GDAP/CSP contract + tries delegated Exchange connection |
| `Test-ExoConnection` | Checks / restores the Exchange Online connection |

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

## Install-Modules.ps1

Installs and imports all PowerShell modules required by this repository. Run once on a new machine or after a clean PowerShell install.

```powershell
.\scripts\Startup\Install-Modules.ps1
```

Core modules installed include `ExchangeOnlineManagement` and required Microsoft Graph submodules (`Microsoft.Graph.Authentication`, `Microsoft.Graph.Sites`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Identity.SignIns`, `Microsoft.Graph.Identity.Governance`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Users`, `Microsoft.Graph.Calendar`).
Windows-only compatibility modules are also included when applicable (`WindowsAutopilotIntune`, `AzureAD`).

---

## Update-Modules.ps1

Updates every installed PowerShell module to its latest version. Run as administrator for system-wide modules.

Also ensures a minimum version for the specific Graph submodules this repo depends on (`Microsoft.Graph.Authentication`, `Microsoft.Graph.Sites`, `Identity.SignIns`, `Identity.Governance`, `Applications`, `Groups`) before updating everything else installed on the machine.

```powershell
.\scripts\Startup\Update-Modules.ps1
```

> No parameters. Iterates every module returned by `Get-InstalledModule`, so it can take a while on a machine with many modules installed.

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
