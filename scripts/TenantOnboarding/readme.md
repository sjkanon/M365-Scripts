# Tenant Onboarding

Scripts for onboarding a new M365 tenant and provisioning/managing its devices — ported and modernized from a retired legacy repository. Every script follows this repo's house style: dry-run by default with an `-Apply` switch for anything that changes state, `[CmdletBinding(SupportsShouldProcess)]`, and no hardcoded credentials, tenant/customer names, or internal endpoints.

Where the legacy source had several near-identical scripts doing slight variations of the same thing (six OneDrive restart scripts, seven hardcoded vendor-installer downloaders, four grant/revoke local-group scripts...), they were consolidated into one well-parameterized script rather than ported one-for-one.

---

## Subfolders

| Folder | Contents |
|--------|----------|
| [`Provisioning/`](Provisioning/readme.md) | Single-tenant bootstrap: break-glass admin account, baseline security groups, Intune baseline policy assignment |
| [`MultiTenant/`](MultiTenant/readme.md) | MSP-wide, cross-customer (GDAP) scripts: license reporting, break-glass password rotation, customer portal index |
| [`AppDeployment/`](AppDeployment/readme.md) | Generic Win32/Chocolatey app installers, default file associations, desktop shortcuts, network printers |
| [`DeviceConfig/`](DeviceConfig/readme.md) | Local group self-elevation, credential storage hardening, kiosk power settings, Office removal, Start Menu layout, Teams firewall rule |
| [`OneDriveManagement/`](OneDriveManagement/readme.md) | OneDrive restart/reset watchdog, per-library sync teardown, Known Folder Move redirection |
| [`UserManagement/`](UserManagement/readme.md) | Dynamic Distribution Group creation by filter, feature-gating group membership |

See each subfolder's readme for the full script list, parameters, and examples.

---

## What was intentionally left out

- **Microsoft Entra ID / macOS Intune sample repository** (`Install New Tenant/MacOS/`): a large vendored collection of shell scripts, `.mobileconfig` profiles, and third-party app installers for macOS — not this MSP's own PowerShell tooling, and out of scope for a PowerShell-house-style port.
- **`Manage Tenant/Compare/Compare-Intune.ps1`**: verified to be the same capability already covered by `scripts/Intune/Compare-IntuneConfig.ps1` (both wrap the `IntuneBackupAndRestore` module's `Compare-IntuneBackupDirectories` to diff a tenant's Intune config against a baseline) — not re-ported.
- Several reporting/config scripts already covered elsewhere in this repo: OneDrive/lockscreen wallpaper (`scripts/Intune/Desktop/Background/`), "pin to Start" (`scripts/Intune/Desktop/Add Lockscreen to start and desktop/`), time sync (`scripts/Device/Time sync/`), and UniFi controller reporting (`scripts/Network/UniFi/`) — the legacy versions were either byte-identical or superseded by a more general version already in this repo.
- Scripts relying on the retired **MSOnline** / **AzureAD** / **AzureADPreview** modules were not ported as-is; their capabilities were reimplemented against Microsoft Graph or Exchange Online (see `MultiTenant/` and `Provisioning/`) or dropped where the underlying technique no longer applies.
- A handful of one-off, single-purpose, or already-obsolete scripts were dropped as genuine dead ends: a build-specific Windows 11 Explorer registry fix scoped to builds below 25211 (long since patched), two duplicate/broken "set default PDF app + import unrelated Start layout" experiments, a trivial local DPAPI credential-file helper, and two third-party (non-MSP-authored) reporting scripts whose functionality is already covered by `scripts/Exchange/Test-MailboxPermissions.ps1` and `scripts/Entra/Test-M365GroupMembership.ps1`.
- The `Setup tenant/` source folder was empty.

No customer names, tenant domains, internal hostnames, IP addresses, or credentials from the source repository were reproduced anywhere in this folder — every script here is written against generic, parameterized inputs.
