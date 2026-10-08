**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [LegacyUtilities](../readme.md) › **Entra**

# Legacy Utilities — Entra

Group membership and Conditional Access utilities via Microsoft Graph. They sign in through [`Connect-M365.ps1`](../../Startup/readme.md): delegated as the admin by default (browser, or device code / GDAP customer per `load.config.ps1`), app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly` (app from `graph.appid.json`). A session for the right tenant that already fits is reused and left connected; only a session the script opened is disconnected. Every script accepts `-TenantId`, `-ClientId`, `-CertificateThumbprint` and `-AppOnly`.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Add-M365GroupMember.ps1`](Add-M365GroupMember.ps1) ([docs](#add-m365groupmemberps1)) | Add or remove group members, single or bulk |
| [`Backup-ConditionalAccessPolicies.ps1`](Backup-ConditionalAccessPolicies.ps1) ([docs](#backup-conditionalaccesspoliciesps1)) | Export every CA policy to individual JSON files |

---

### Add-M365GroupMember.ps1

Adds or removes a single member or a CSV/TXT list of members to/from a group
(by object ID or display name). Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-GroupId` | Yes | Object ID or display name of the group |
| `-Member` | * | Single UPN/object ID |
| `-CsvPath` | * | CSV/TXT list of members |
| `-Action` | No | `Add` (default) or `Remove` |
| `-Apply` | No | Actually change membership (default: preview) |
| `-TenantId` | No | Entra ID tenant ID or domain (default: the GDAP customer when `authMode` is GDAP) |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in with this app registration and certificate |
| `-AppOnly` | No | App-only sign-in with the app from `graph.appid.json` |

```powershell
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply
```

**Notes**
- A group display name with quotes is escaped for the filter, and a name that matches more than one group stops the script (it used to take the first match)
- Delegated scopes: `GroupMember.ReadWrite.All`, `Group.Read.All`, `User.Read.All` (the last one was missing for the member lookup)

---

### Backup-ConditionalAccessPolicies.ps1

Exports every Conditional Access policy to one JSON file per policy (named by
policy ID), plus an `_index.csv` summary. Read-only. Modernized replacement for
an old script that used the retired AzureADPreview module.

```powershell
.\Backup-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\CA-2026-07-24"
```

---

**Required modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Notes**
- Each file holds the policy exactly as Graph returns it (`GET /identity/conditionalAccess/policies`, paged); before, the SDK objects were serialized with `-Depth 10`, which adds SDK wrapper properties and can cut off nested conditions
- Delegated scope: `Policy.Read.All`
