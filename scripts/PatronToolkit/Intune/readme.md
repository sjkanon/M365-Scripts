**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [PatronToolkit](../readme.md) › **Intune**

# Patron Toolkit — Intune

Intune policy assignment reporting and Windows Autopilot device inventory via Microsoft
Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-IntunePolicyAssignments.ps1`](Get-IntunePolicyAssignments.ps1) ([docs](#get-intunepolicyassignmentsps1)) | Report which groups are assigned to which Intune profiles/policies/apps |
| [`Get-AutopilotDevices.ps1`](Get-AutopilotDevices.ps1) ([docs](#get-autopilotdevicesps1)) | Report registered Windows Autopilot devices and deployment profiles |

---

### Get-IntunePolicyAssignments.ps1

Enumerates Intune device configuration profiles, compliance policies, settings catalog
policies, and mobile apps, resolving each object's assignments to readable group names
(or "All users"/"All devices"), and whether the assignment is Include or Exclude; app rows
also carry the intent (required/available/uninstall). One CSV row per policy/assignment
pair.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-PolicyType` | No | `DeviceConfiguration`, `CompliancePolicy`, `SettingsCatalog`, `MobileApp` (default: all) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`). Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in with the ClientId and thumbprint for the tenant from `graph.appid.json` |

**Examples**

```powershell
.\Get-IntunePolicyAssignments.ps1

.\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog
```

**Notes**
- Read-only. Bulk assignment changes were intentionally not scripted — review this
  report's output and make assignment changes per-policy in the Intune admin center
- All calls are `Invoke-MgGraphRequest` with `$expand=assignments` and `@odata.nextLink`
  paging. Settings catalog policies only exist in Graph **beta**
  (`/beta/deviceManagement/configurationPolicies`); the earlier version called
  `Get-MgDeviceManagementConfigurationPolicy`, which the Microsoft.Graph v2 SDK does not
  have, and silently reported no settings catalog policies. A type that cannot be read is
  now a visible warning instead of an empty result
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): Microsoft
  Graph, delegated by default (scopes `DeviceManagementConfiguration.Read.All`,
  `DeviceManagementApps.Read.All`, `Group.Read.All` plus an Intune role); app-only with
  `-ClientId` + `-CertificateThumbprint` or `-AppOnly` (the same application permissions)

---

### Get-AutopilotDevices.ps1

Lists every Windows Autopilot device identity registered in the tenant (serial number,
model, manufacturer, group tag, enrollment state, deployment profile assignment status)
plus a summary of the deployment profiles that exist. This is a tenant-side inventory
report — different from
[`Get-Autopilot/Get-WindowsAutoPilotInfo.ps1`](../../Intune/readme.md), which collects a
hardware hash from a physical device to register it.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-GroupTag` | No | Only report devices with this group tag |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`). Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in with the ClientId and thumbprint for the tenant from `graph.appid.json` |

**Examples**

```powershell
.\Get-AutopilotDevices.ps1

.\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"
```

**Notes**
- Read-only. Bulk device import/assign/delete were intentionally not scripted — use
  `Get-WindowsAutoPilotInfo.ps1` (already in this repo) plus the Intune admin center for
  registration, and review this report before any bulk reassignment
- Deployment profiles only exist in Graph **beta**
  (`/beta/deviceManagement/windowsAutopilotDeploymentProfiles`) and are read with
  `Invoke-MgGraphRequest`; the earlier version used a cmdlet the Microsoft.Graph v2 SDK
  does not have and always listed zero profiles. Devices come from Graph v1.0
- The "without an assigned deployment profile" count now treats `assignedInSync`,
  `assignedOutOfSync` and `assignedUnkownSyncState` as assigned — Graph has no plain
  `assigned` value, so every device used to be counted as unassigned
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): Microsoft
  Graph, delegated by default (scope `DeviceManagementServiceConfig.Read.All` plus an
  Intune role); app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly`

**Required modules**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
