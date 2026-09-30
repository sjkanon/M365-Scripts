[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Entra**

# Legacy Utilities — Entra

Utilitaires d'appartenance aux groupes et de Conditional Access via Microsoft Graph. Ils se
connectent automatiquement si aucune session n'est active, et réutilisent la session existante
si vous êtes déjà connecté.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Add-M365GroupMember.ps1`](Add-M365GroupMember.ps1) ([docs](#add-m365groupmemberps1)) | Ajouter ou retirer des membres d'un groupe, individuellement ou en masse |
| [`Backup-ConditionalAccessPolicies.ps1`](Backup-ConditionalAccessPolicies.ps1) ([docs](#backup-conditionalaccesspoliciesps1)) | Exporter chaque stratégie CA dans un fichier JSON distinct |

---

### Add-M365GroupMember.ps1

Ajoute à un groupe, ou en retire, un membre unique ou une liste de membres au format CSV/TXT
(par ID d'objet ou nom d'affichage). Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-GroupId` | Oui | ID d'objet ou nom d'affichage du groupe |
| `-Member` | * | UPN/ID d'objet unique |
| `-CsvPath` | * | Liste de membres CSV/TXT |
| `-Action` | Non | `Add` (par défaut) ou `Remove` |
| `-Apply` | Non | Modifier réellement l'appartenance (par défaut : aperçu) |

```powershell
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply
```

---

### Backup-ConditionalAccessPolicies.ps1

Exporte chaque stratégie Conditional Access dans un fichier JSON par stratégie (nommé d'après
l'ID de la stratégie), plus un récapitulatif `_index.csv`. Lecture seule. Remplacement modernisé
d'un ancien script qui utilisait le module retiré AzureADPreview.

```powershell
.\Backup-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\CA-2026-07-24"
```

---

**Modules requis**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
