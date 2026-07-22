# Startup

Entry-point scripts and the core M365 function library.

---

## Files

| File | Description |
|------|-------------|
| `functies.ps1` | M365 function library — dot-sourced by `menu.ps1` on first use |
| `Install-Modules.ps1` | Bootstrap script — installs and imports all required PowerShell modules |
| `Update-Modules.ps1` | Updates every installed PowerShell module to its latest version |
| `Test-PowerShellSyntax.ps1` | Parse-checks `.ps1` files in the repo for syntax errors, no execution |

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

Modules installed: `ExchangeOnlineManagement`, `Microsoft.Graph`, `ImportExcel`, `PSWindowsUpdate`
Key Graph submodules include `Microsoft.Graph.Authentication` and `Microsoft.Graph.Identity.DirectoryManagement` for delegated/GDAP tenant contract lookups.

---

## Update-Modules.ps1

Updates every installed PowerShell module to its latest version. Run as administrator for system-wide modules.

Also ensures a minimum version for the specific Graph submodules this repo depends on (`Microsoft.Graph.Authentication`, `Identity.SignIns`, `Identity.Governance`, `Applications`, `Groups`) before updating everything else installed on the machine.

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
