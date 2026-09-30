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
| `-IncludeTeamsInventory` | Non | Lister aussi chaque équipe avec le nombre de membres/propriétaires (par défaut : activé) |
| `-OutputPath` | Non | Dossier du ou des rapports CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-TeamsConfigReport.ps1

.\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false
```

**Remarques**
- Lecture seule

**Module requis**
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
```
