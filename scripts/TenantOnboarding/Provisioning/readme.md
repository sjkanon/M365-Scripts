**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [TenantOnboarding](../readme.md) › **Provisioning**

# Provisioning

Scripts for bootstrapping a single newly onboarded tenant: break-glass admin account, baseline security groups, and applying baseline Intune policy assignments. Replaces an old interactive menu-driven install script with atomic, parameterized scripts consistent with the rest of this repo — run them in the order below as part of a new-tenant checklist.

**Sign-in** (all three scripts) goes through [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1): **delegated by default** — you sign in as an admin of the tenant (device code when `$global:useDeviceCodeAuth` is set; under GDAP the customer tenant comes from `$global:cid` unless `-TenantId` names one). **App-only** is an option with `-ClientId` + `-CertificateThumbprint`, or `-AppOnly` to read them from `graph.appid.json`. An existing Graph session is reused only when it is for the right tenant and already holds the scopes the script needs; the script disconnects only a session it opened itself.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-BreakGlassAdminAccount.ps1`](New-BreakGlassAdminAccount.ps1) ([docs](#new-breakglassadminaccountps1)) | Create a cloud-only emergency-access Global Administrator account |
| [`New-TenantBaselineGroups.ps1`](New-TenantBaselineGroups.ps1) ([docs](#new-tenantbaselinegroupsps1)) | Create the CA-exclusion group and any feature-toggle groups a new tenant needs |
| [`Set-IntuneBaselinePolicyAssignment.ps1`](Set-IntuneBaselinePolicyAssignment.ps1) ([docs](#set-intunebaselinepolicyassignmentps1)) | Bulk-assign name-filtered Intune baseline policies to a target group |

**Recommended order for a new tenant:**
1. `New-BreakGlassAdminAccount.ps1` — create the emergency-access account
2. `New-TenantBaselineGroups.ps1` — create the "all users except break glass" group (its dynamic rule leaves out the break-glass account's UPN pattern) and any feature-toggle groups
3. Only if your Conditional Access policies exclude a *static* group: add the break-glass account to it with `New-BreakGlassAdminAccount.ps1 -ExcludeFromGroupId <group id>` at creation, or later with `Add-UserToFeatureGroup.ps1` from `../UserManagement/`. The dynamic group from step 2 needs no members added (and takes none).
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
| `-ExcludeFromGroupId` | No | A static CA-exclusion group to add the account to (not the dynamic group from `New-TenantBaselineGroups.ps1`) |
| `-TenantId` | No | Entra ID tenant ID or domain (default: GDAP customer, else the tenant you sign in to) |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |
| `-Apply` | No | Actually create the account (default: preview only) |

**Examples**

```powershell
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com"
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply
```

**Notes**
- Scopes: `User.ReadWrite.All`, `RoleManagement.ReadWrite.Directory`, `GroupMember.ReadWrite.All`.
- Global Administrator is assigned with a unified role assignment (`roleManagement/directory/roleAssignments`, role `62e90394-69f5-4237-9190-012177145e10`). The earlier code tried to activate the role with `New-MgDirectoryRoleTemplate -RoleTemplateId`; that cmdlet creates a role template and has no `-RoleTemplateId` parameter, so a tenant where the role had never been activated failed at that step.

**Required modules**
```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.Governance -Scope CurrentUser
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
| `-TenantId` | No | Entra ID tenant ID or domain (default: GDAP customer, else the tenant you sign in to) |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |
| `-Apply` | No | Actually create the groups (default: preview only) |

*Required unless `-SkipExclusionGroup` is used.

**Examples**

```powershell
.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" `
    -AdditionalGroupNames "SG - Enable Password Manager","SG - Enable Windows 365"

.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" -Apply
```

**Notes**
- Scope: `Group.ReadWrite.All`. Dynamic membership needs Entra ID P1.

**Required module**
```powershell
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### Set-IntuneBaselinePolicyAssignment.ps1

Bulk-assigns Intune configuration profiles, compliance policies, admin templates, platform scripts, and security baselines whose display name matches a filter to a single target group. The group is **added** to each policy's existing assignments.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TargetGroupId` | Yes | Group to assign matching policies to |
| `-NameFilter` | No | Wildcard filter on policy display name (default: `*Default*`) |
| `-PolicyTypes` | No | Which object types to include (default: all) |
| `-TenantId` | No | Entra ID tenant ID or domain (default: GDAP customer, else the tenant you sign in to) |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |
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
- Intune's `/assign` action replaces a policy's whole assignment list. The script now reads the current assignments and posts them back together with the new group; before, every policy it touched lost its other assignments. Policies already assigned to the group are skipped.
- Policy lists follow `@odata.nextLink`, so tenants with more policies than one page are fully covered.
- Settings catalog policies (`configurationPolicies`) are not included.
- Scopes: `DeviceManagementConfiguration.ReadWrite.All`, `DeviceManagementServiceConfig.ReadWrite.All`.
