# Startup Scripts

Generic MSP M365 management function library for multi-tenant Microsoft 365 administration.
Part of the [M365-Scripts](../../readme.md) repository.

---

## Files

| File | Description |
|---|---|
| `functies.ps1` | Full function library — dot-source this file at startup |

---

## functies.ps1

Central function library for multi-tenant M365 management via Microsoft Graph and Exchange Online.
Fully rewritten from MSOnline + AzureAD to Microsoft Graph.

### Requirements

| Requirement | Value |
|---|---|
| PowerShell | 7.0 or later |
| Modules | See `#Requires` at the top of the file |
| Connection | Started automatically via `Connect-MgGraph` on dot-source |

### Configuration

Edit the `#region Configuration` block at the top of the file to match your organisation:

```powershell
$script:MspAdminAlias       = 'msp-admin'        # Mailbox alias for the MSP admin account
$script:MspAdminDisplayName = 'MSP - Admin Account'  # Display name for the MSP admin account
```

### Getting started

Dot-source the file from your PowerShell profile or startup script:

```powershell
$upn      = "admin@yourdomain.com"   # UPN of the administrator
$realname = "Sjoerd"                 # Display name for greeting (optional)

. "$PSScriptRoot\scripts\Startup\functies.ps1"
```

### Select a customer tenant (CSP)

```powershell
Connect-Tenant -Domain "customer.com"
# Populates $global:cid and $global:connectmsoldomain
# All subsequent functions automatically target the selected tenant
```

---

## Functions

### Connection

| Function | Description |
|---|---|
| `Connect-Tenant` | Select a CSP customer by domain name, populates `$cid` and `$connectmsoldomain` |
| `Test-ExoConnection` | Checks / restores the Exchange Online connection for the selected customer |
| `Invoke-Menu` | Interactive menu to load service modules (EXO, Entra, Teams, Intune) |

### Exchange Online

| Function | Description |
|---|---|
| `Enable-CopyOfSentItems` | Enables copy of sent items for all mailboxes |
| `Add-SharedMailboxAccess` | Grants a user FullAccess + SendAs on a shared mailbox |
| `Set-MailboxLocale` | Sets language and timezone on all mailboxes (default: NL / W. Europe) |
| `Add-MailboxAlias` | Adds an alias to a mailbox |
| `Get-MailboxAliases` | Lists all SMTP aliases per mailbox |
| `Export-DistributionGroups` | Exports all distribution groups to CSV (`%TEMP%\ExportDGs.csv`) |
| `Set-AutoReply` | Configures an out-of-office reply (enabled / disabled / scheduled) |

### Microsoft Entra ID / Graph

| Function | Description |
|---|---|
| `Get-TenantAdmins` | Lists all Global Administrators in the customer tenant |
| `Add-TenantDomain` | Adds a domain and walks through the verification process |
| `Get-TenantLicenses` | Shows licence overview with usage and availability |
| `Get-TenantUsers` | Lists all users with UPN, display name, and licences |
| `Add-TenantAdmin` | Grants a user Global Administrator rights |
| `Get-EntraApplication` | Finds an Enterprise App by name |
| `Reset-UserPassword` | Resets a user's password in a CSP customer tenant |
| `Export-SignInLogs` | Exports sign-in logs to CSV (default 30 days, max 30) |

### MSP Admin Account

| Function | Description |
|---|---|
| `New-MspAdmin` | Creates the MSP admin account as Global Admin in the customer tenant |
| `Set-MspAdminAsGroupOwner` | Sets the MSP admin account as owner of a group |
| `Reset-MspAdminPassword` | Resets the MSP admin account password |

### Navigation

| Function | Description |
|---|---|
| `Set-ImportLocation` | Navigates to `$env:import` |
| `Set-ScriptsLocation` | Navigates to `$env:ps` |

---

## Install modules

Use the bootstrap script from the repository root:

```powershell
.\scripts\Install-Modules.ps1
```

---

*Part of M365-Scripts — Sjoerd Kanon*
