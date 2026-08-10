# Legacy Utilities

Modernized, house-style equivalents of a batch of scripts from a retired internal
repository. None of these are verbatim copies of the originals — several old
one-off scripts that did slight variations of the same task were consolidated
into a single well-parameterized script, dead modules (MSOnline/AzureAD) were
replaced with Microsoft Graph / Exchange Online equivalents, and every hardcoded
customer name, tenant domain, hostname, password, or secret found in the
originals was generalized into a parameter — none of that data was carried over.

This folder is organized by theme, one subfolder per area:

| Folder | Contents |
|--------|----------|
| [`Exchange/`](Exchange/readme.md) | Mailbox delegate access, bulk shared mailbox/contact creation, contact sync, message trace, mailbox dedup |
| [`Entra/`](Entra/readme.md) | Group membership changes, Conditional Access policy backup |
| [`Teams/`](Teams/readme.md) | Team cloning, Planner plan copying, bulk project Team creation |
| [`Network/`](Network/readme.md) | Azure Files SMB share mounting |
| [`Device/`](Device/readme.md) | Num Lock default, Lock Workstation shortcut |
| [`Workspace365/`](Workspace365/readme.md) | Workspace 365 environment provisioning/deprovisioning |

None of these scripts are wired into [`menu.ps1`](../../menu.ps1) or the
top-level docs yet — that integration is a separate pass.

---

## What was intentionally left out

A large share of the source material was skipped rather than ported — either
because it was already fully covered elsewhere in this repo, because it was a
genuine dead end, or because it only existed as real customer/tenant data that
must never be reproduced. See the porting report for the full breakdown; in
short:

- **Already covered elsewhere in this repo**: Windows Autopilot info gathering
  (`scripts/Intune/Get-Autopilot/`), general disk/log cleanup
  (`scripts/Device/Invoke-WindowsCleanup.ps1`), OEM bloatware removal
  (`scripts/Device/Remove-OemBloatware.ps1`), UniFi network reporting/firmware
  updates (`scripts/Network/UniFi/`), the CSP/GDAP tenant-connect pattern
  (`scripts/Startup/functies.ps1`), calendar folder permissions
  (`scripts/Exchange/Set-Calendar-rights.ps1`), and a set of device/AppDeployment
  one-offs (desktop URL shortcuts, Start Menu shortcut copying, Office
  uninstall, Teams LAN firewall rule) already generalized under
  `scripts/TenantOnboarding/`.
- **Dead ends**: a full PrintNightmare (2021 spooler CVE) mitigation kit built
  around the deprecated `subinacl.exe` tool — the vulnerability has been patched
  for years and the workaround has no ongoing value.
- **Real customer/tenant data, not code**: several old scripts and JSON exports
  contained live tenant IDs, group/user object GUIDs, storage account keys, app
  client secrets, or customer email addresses/domains. These were never
  reproduced anywhere (not even in reports) per the project's data-handling
  rules; only the generic *capability* behind them (e.g. "back up Conditional
  Access policies", "mount an Azure Files share") was ported, with all
  identifying values turned into parameters.
