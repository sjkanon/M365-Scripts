[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Entra**

# Legacy Utilities — Entra

Hulpscripts voor groepslidmaatschap en Conditional Access via Microsoft Graph. Ze maken
automatisch verbinding als er geen sessie actief is, en hergebruiken een bestaande sessie als
je al verbonden bent.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Add-M365GroupMember.ps1`](Add-M365GroupMember.ps1) ([docs](#add-m365groupmemberps1)) | Groepsleden toevoegen of verwijderen, afzonderlijk of in bulk |
| [`Backup-ConditionalAccessPolicies.ps1`](Backup-ConditionalAccessPolicies.ps1) ([docs](#backup-conditionalaccesspoliciesps1)) | Exporteer elk CA-beleid naar een eigen JSON-bestand |

---

### Add-M365GroupMember.ps1

Voegt één lid of een CSV/TXT-lijst met leden toe aan een groep, of verwijdert ze eruit
(op object-ID of weergavenaam). Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-GroupId` | Ja | Object-ID of weergavenaam van de groep |
| `-Member` | * | Eén UPN/object-ID |
| `-CsvPath` | * | CSV/TXT-lijst met leden |
| `-Action` | Nee | `Add` (standaard) of `Remove` |
| `-Apply` | Nee | Wijzig het lidmaatschap echt (standaard: voorbeeldweergave) |

```powershell
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply
```

---

### Backup-ConditionalAccessPolicies.ps1

Exporteert elk Conditional Access-beleid naar één JSON-bestand per beleid (genoemd naar het
beleids-ID), plus een samenvatting in `_index.csv`. Alleen-lezen. Gemoderniseerde vervanging
van een oud script dat de uitgefaseerde AzureADPreview-module gebruikte.

```powershell
.\Backup-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\CA-2026-07-24"
```

---

**Vereiste modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
