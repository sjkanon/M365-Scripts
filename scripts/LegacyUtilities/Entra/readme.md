# Legacy Utilities — Entra

Group membership and Conditional Access utilities via Microsoft Graph. Connect
automatically if no session is active; reuse an existing session if already
connected.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Add-M365GroupMember.ps1`](#add-m365groupmemberps1) | Add or remove group members, single or bulk |
| [`Backup-ConditionalAccessPolicies.ps1`](#backup-conditionalaccesspoliciesps1) | Export every CA policy to individual JSON files |

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

```powershell
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply
```

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
