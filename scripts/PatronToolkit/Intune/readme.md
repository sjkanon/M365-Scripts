# Patron Toolkit — Intune

Intune policy assignment reporting and Windows Autopilot device inventory via Microsoft
Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-IntunePolicyAssignments.ps1`](#get-intunepolicyassignmentsps1) | Report which groups are assigned to which Intune profiles/policies/apps |
| [`Get-AutopilotDevices.ps1`](#get-autopilotdevicesps1) | Report registered Windows Autopilot devices and deployment profiles |

---

### Get-IntunePolicyAssignments.ps1

Enumerates Intune device configuration profiles, compliance policies, settings catalog
policies, and mobile apps, resolving each object's assignments to readable group names
(or "All users"/"All devices"), and whether the assignment is Include or Exclude. One CSV
row per policy/assignment pair.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-PolicyType` | No | `DeviceConfiguration`, `CompliancePolicy`, `SettingsCatalog`, `MobileApp` (default: all) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-IntunePolicyAssignments.ps1

.\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog
```

**Notes**
- Read-only. Bulk assignment changes were intentionally not scripted — review this
  report's output and make assignment changes per-policy in the Intune admin center
- Required scopes: `DeviceManagementConfiguration.Read.All`,
  `DeviceManagementApps.Read.All`, `Group.Read.All`

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
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-AutopilotDevices.ps1

.\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"
```

**Notes**
- Read-only. Bulk device import/assign/delete were intentionally not scripted — use
  `Get-WindowsAutoPilotInfo.ps1` (already in this repo) plus the Intune admin center for
  registration, and review this report before any bulk reassignment
- Required scope: `DeviceManagementServiceConfig.Read.All`

**Required modules**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
