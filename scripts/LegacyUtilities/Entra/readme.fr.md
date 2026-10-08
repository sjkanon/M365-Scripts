[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Entra**

# Legacy Utilities — Entra

Utilitaires d'appartenance aux groupes et de Conditional Access via Microsoft Graph. Ils se connectent via [`Connect-M365.ps1`](../../Startup/readme.fr.md) : en délégué en tant qu'administrateur par défaut (navigateur, ou code d'appareil / client GDAP selon `load.config.ps1`), en app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly` (application de `graph.appid.json`). Une session adaptée pour le bon tenant est réutilisée et reste connectée ; seule une session ouverte par le script est fermée. Chaque script accepte `-TenantId`, `-ClientId`, `-CertificateThumbprint` et `-AppOnly`.

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
| `-TenantId` | Non | ID ou domaine du tenant Entra ID (par défaut : le client GDAP si `authMode` vaut GDAP) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion app-only avec cette inscription d'application et ce certificat |
| `-AppOnly` | Non | Connexion app-only avec l'application de `graph.appid.json` |

```powershell
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply
```

**Remarques**
- Un nom de groupe contenant des apostrophes est échappé pour le filtre, et un nom correspondant à plus d'un groupe arrête le script (auparavant, la première correspondance était prise)
- Étendues déléguées : `GroupMember.ReadWrite.All`, `Group.Read.All`, `User.Read.All` (cette dernière manquait pour la recherche des membres)

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

**Remarques**
- Chaque fichier contient la stratégie exactement telle que Graph la renvoie (`GET /identity/conditionalAccess/policies`, paginé) ; auparavant, les objets du SDK étaient sérialisés avec `-Depth 10`, ce qui ajoute des propriétés d'enveloppe du SDK et peut tronquer les conditions imbriquées
- Étendue déléguée : `Policy.Read.All`
