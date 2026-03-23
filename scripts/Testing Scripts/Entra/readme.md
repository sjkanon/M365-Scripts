# Testing — Entra ID / Graph

Audit scripts for Microsoft Entra ID (formerly Azure AD) via Microsoft Graph.
Connect to Graph automatically if no session is active; reuse an existing session if already connected.

---

## Scripts

### Test-M365GroupMembership.ps1

Lists all owners and members of Microsoft 365 Groups (including Teams-backed groups).
Results are exported to CSV with one row per owner/member entry.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Group` | No | Display name or Object ID of a single group. If omitted, all M365 groups are audited |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Audit all M365 groups
.\Test-M365GroupMembership.ps1

# Single group by display name
.\Test-M365GroupMembership.ps1 -Group "Team Finance"

# Single group by Object ID
.\Test-M365GroupMembership.ps1 -Group "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

**Required scopes**
- `Group.Read.All`
- `Directory.Read.All`

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
