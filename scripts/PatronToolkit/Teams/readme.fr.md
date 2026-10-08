[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [PatronToolkit](../readme.fr.md) › **Teams**

# Patron Toolkit — Teams

Gouvernance du tenant Microsoft Teams et rapports d'inventaire.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-TeamsConfigReport.ps1`](Get-TeamsConfigReport.ps1) ([docs](#get-teamsconfigreportps1)) | Rapport des paramètres de gouvernance du tenant Teams et inventaire des équipes |

---

### Get-TeamsConfigReport.ps1

Rapporte les stratégies Teams à l'échelle du tenant qui comptent le plus lors d'une revue de gouvernance/sécurité :
accès externe (fédération), accès invité, stratégie de réunion globale (participation
anonyme, enregistrement, rôle de présentateur), stratégie de messagerie globale et stratégie de configuration des applications
(chargement de version test/sideloading). Liste aussi chaque équipe avec sa visibilité, son état d'archivage et le nombre de
propriétaires/membres — en signalant toute équipe sans propriétaire.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-IncludeTeamsInventory` | Non | Lister aussi chaque équipe avec le nombre de membres/propriétaires, via Graph (par défaut : activé) |
| `-OutputPath` | Non | Dossier du ou des rapports CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine. Par défaut le client GDAP (`load.config.ps1`) ou votre propre tenant ; obligatoire en app-only |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`). Sans lui, le script se connecte en délégué, en votre nom |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only avec `-ClientId` |
| `-AppOnly` | Non | Connexion app-only avec le ClientId et l'empreinte du tenant depuis `graph.appid.json` |

**Exemples**

```powershell
.\Get-TeamsConfigReport.ps1

.\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false

# App-only, Teams et Graph avec l'application de graph.appid.json
.\Get-TeamsConfigReport.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Remarques**
- Lecture seule
- Les stratégies de tenant `Cs*` restent sur Teams PowerShell (`Connect-M365Teams`) —
  Microsoft Graph n'a pas d'API pour elles. L'inventaire des équipes passe à Graph : `/groups`
  filtré sur le provisionnement Team, `/teams/{id}` (`isArchived`) et `/teams/{id}/members`
  (rôle propriétaire), à la place de `Get-Team` / `Get-TeamUser`
- L'accès invité lit désormais `AllowGuestUser` (`Get-CsTeamsClientConfiguration`) et
  `DisableAnonymousJoin` (`Get-CsTeamsMeetingConfiguration`) ; la version précédente lisait
  `AllowAnonymousUsersToJoinMeeting` dans la configuration des réunions invités, qui n'a pas
  cette propriété, et affichait toujours une valeur vide
- Connexion via [`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1), en délégué
  par défaut (rôle Teams Administrator ou Global Reader ; étendues Graph `Group.Read.All`,
  `TeamMember.Read.All`, `TeamSettings.Read.All`) ; app-only avec `-ClientId` +
  `-CertificateThumbprint` ou `-AppOnly` pour les deux. Avec `-IncludeTeamsInventory:$false`,
  aucune connexion Graph n'est ouverte

**Modules requis**
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```
