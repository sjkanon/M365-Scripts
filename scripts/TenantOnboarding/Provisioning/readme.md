# Provisioning

Scripts for bootstrapping a single newly onboarded tenant: break-glass admin account, baseline security groups, and applying baseline Intune policy assignments. Replaces an old interactive menu-driven install script with atomic, parameterized scripts consistent with the rest of this repo — run them in the order below as part of a new-tenant checklist.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-BreakGlassAdminAccount.ps1`](#new-breakglassadminaccountps1) | Create a cloud-only emergency-access Global Administrator account |
| [`New-TenantBaselineGroups.ps1`](#new-tenantbaselinegroupsps1) | Create the CA-exclusion group and any feature-toggle groups a new tenant needs |
| [`Set-IntuneBaselinePolicyAssignment.ps1`](#set-intunebaselinepolicyassignmentps1) | Bulk-assign name-filtered Intune baseline policies to a target group |

**Recommended order for a new tenant:**
1. `New-BreakGlassAdminAccount.ps1` — create the emergency-access account
2. `New-TenantBaselineGroups.ps1` — create the CA-exclusion group (referencing the break-glass account's UPN pattern) and any feature-toggle groups
3. `New-BreakGlassAdminAccount.ps1 -ExcludeFromGroupId <exclusion group id>` (or re-run `Add-UserToFeatureGroup.ps1` from `../UserManagement/`) to make sure the break-glass account is a member of its own exclusion group
4. Import/configure your Conditional Access and Intune baseline policies (e.g. `Import-ConditionalAccessBaseline.ps1` in `scripts/Entra/`)
5. `Set-IntuneBaselinePolicyAssignment.ps1` — assign the baseline Intune policies to the "all users except break glass" group

---

### New-BreakGlassAdminAccount.ps1

Creates a cloud-only Entra ID user as an emergency-access ("break-glass") account, following Microsoft's documented guidance: dedicated cloud-only account, long random password (printed once, never saved to disk), Global Administrator assigned directly. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | Yes | UPN for the new account — use the tenant's `*.onmicrosoft.com` domain |
| `-DisplayName` | No | Display name (default: `Break Glass Admin`) |
| `-PasswordLength` | No | Generated password length (default: `24`) |
| `-AssignGlobalAdmin` | No | Assign Global Administrator (default: on) |
| `-ExcludeFromGroupId` | No | Group to add the account to (e.g. a CA-exclusion group) |
| `-TenantId` | No | Entra ID tenant ID or domain |
| `-Apply` | No | Actually create the account (default: preview only) |

**Examples**

```powershell
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com"
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply
```

**Required modules**
```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### New-TenantBaselineGroups.ps1

Creates a dynamic "all users except break-glass accounts" exclusion group plus any number of static feature-toggle groups you name — for scoping Conditional Access and Intune assignments in a freshly onboarded tenant.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-BreakGlassUpnPattern` | * | Pattern matched against UPNs to build the exclusion group's dynamic rule |
| `-ExclusionGroupName` | No | Display name for the exclusion group |
| `-SkipExclusionGroup` | No | Skip the exclusion group entirely |
| `-AdditionalGroupNames` | No | Names of additional static groups to create |
| `-TenantId` | No | Entra ID tenant ID or domain |
| `-Apply` | No | Actually create the groups (default: preview only) |

*Required unless `-SkipExclusionGroup` is used.

**Examples**

```powershell
.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" `
    -AdditionalGroupNames "SG - Enable Password Manager","SG - Enable Windows 365"

.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" -Apply
```

**Required module**
```powershell
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### Set-IntuneBaselinePolicyAssignment.ps1

Bulk-assigns Intune configuration profiles, compliance policies, admin templates, scripts, and security baselines whose display name matches a filter to a single target group.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TargetGroupId` | Yes | Group to assign matching policies to |
| `-NameFilter` | No | Wildcard filter on policy display name (default: `*Default*`) |
| `-PolicyTypes` | No | Which object types to include (default: all) |
| `-TenantId` | No | Entra ID tenant ID or domain |
| `-Apply` | No | Actually create assignments (default: preview only) |

**Examples**

```powershell
.\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
.\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

**Required module**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

**Notes**
- Uses Microsoft Graph's beta endpoint — Intune policy assignment isn't fully exposed on v1.0 for every policy type.
- All scripts in this folder connect to Microsoft Graph automatically if no session is active, and reuse an existing session if already connected — consistent with `Test-M365GroupMembership.ps1` in `scripts/Entra/`.
