[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [PatronToolkit](../readme.fr.md) › **Exchange**

# Patron Toolkit — Exchange

Diagnostics du flux de messagerie via Exchange Online.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MessageTraceReport.ps1`](Get-MessageTraceReport.ps1) ([docs](#get-messagetracereportps1)) | Exporter un rapport de suivi des messages (flux de messagerie), avec en option le détail de remise par message |

---

### Get-MessageTraceReport.ps1

Exécute un suivi des messages sur une fenêtre de temps donnée (par défaut : les 48 dernières heures), éventuellement
filtré par expéditeur, destinataire et état de remise. Exporte un CSV récapitulatif et, avec
`-IncludeDetail`, un CSV de détail de remise par message.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Hours` | Non | Nombre d'heures à remonter à partir de maintenant (par défaut : `48`). Ignoré si `-StartDate` est fourni |
| `-StartDate` | Non | Début explicite de la fenêtre de suivi |
| `-EndDate` | Non | Fin explicite de la fenêtre de suivi (par défaut : maintenant) |
| `-SenderAddress` | Non | Filtrer sur un expéditeur précis |
| `-RecipientAddress` | Non | Filtrer sur un destinataire précis |
| `-Status` | Non | Filtrer par état de remise (p. ex. `Delivered`, `Failed`, `Pending`, `Quarantined`) |
| `-IncludeDetail` | Non | Exporter aussi le détail de remise par message (plus lent sur de gros volumes de résultats) |
| `-OutputPath` | Non | Dossier du ou des rapports CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail
```

**Remarques**
- `Get-MessageTrace` ne couvre que les 10 derniers jours — pour du courrier plus ancien, utilisez plutôt la recherche
  historique du centre d'administration Exchange

**Module requis**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
