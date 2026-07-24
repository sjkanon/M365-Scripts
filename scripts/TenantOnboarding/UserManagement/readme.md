# UserManagement

Small Entra ID / Exchange Online user and group management helpers used during tenant onboarding and ongoing feature rollout.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-DynamicDistributionGroupByFilter.ps1`](#new-dynamicdistributiongroupbyfilterps1) | Create a Dynamic Distribution Group from a job title or custom recipient filter |
| [`Add-UserToFeatureGroup.ps1`](#add-usertofeaturegroupps1) | Add/remove a user from a named Entra ID group to gate an add-on feature |

---

### New-DynamicDistributionGroupByFilter.ps1

Previews recipients matching a job-title (or custom) Exchange filter, then creates a Dynamic Distribution Group using that filter after confirmation. Requires an active Exchange Online session — Dynamic Distribution Groups are an Exchange feature, not a Graph one.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-JobTitle` | * | Job title to filter on |
| `-RecipientFilter` | * | Custom Exchange recipient filter string |
| `-Name` | No | Group name (default: the job title; required with `-RecipientFilter`) |
| `-PrimarySmtpAddress` | No | SMTP address for the new group |
| `-Apply` | No | Actually create the group (default: preview only) |

*One of `-JobTitle` / `-RecipientFilter` is required.

```powershell
.\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager"
.\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager" -Apply

.\New-DynamicDistributionGroupByFilter.ps1 -Name "Finance Dept" `
    -RecipientFilter "((Department -eq 'Finance') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))" -Apply
```

**Required module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Add-UserToFeatureGroup.ps1

Adds or removes a user as a direct member of an Entra ID security group — the supported way to gate a Conditional Access policy, Intune assignment, or license group to a subset of users. Replaces the legacy technique of tagging a user via a mailbox custom attribute (reporting-only, did not actually scope anything).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserId` | Yes | UPN or object ID |
| `-GroupId` | * | Target group object ID |
| `-GroupName` | * | Target group display name (resolved automatically) |
| `-Remove` | No | Remove instead of add |
| `-TenantId` | No | Entra ID tenant ID or domain |
| `-Apply` | No | Actually change membership (default: preview only) |

*One of `-GroupId` / `-GroupName` is required.

```powershell
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Apply
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Remove -Apply
```

**Required scopes:** `GroupMember.ReadWrite.All`, `User.Read.All`, `Group.Read.All`
**Required modules:** `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`

---

Connects to Microsoft Graph automatically if no session is active; reuses an existing session if already connected — consistent with `Test-M365GroupMembership.ps1` in `scripts/Entra/`.
