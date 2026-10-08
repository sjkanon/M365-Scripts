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
filtré par expéditeur, destinataire et état de remise, avec `Get-MessageTraceV2`. Exporte un CSV
récapitulatif et, avec `-IncludeDetail`, un CSV de détail de remise par message
(`Get-MessageTraceDetailV2`).

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
| `-MaxResults` | Non | Arrêter la pagination après ce nombre de lignes de suivi (par défaut : `50000`) |
| `-OutputPath` | Non | Dossier du ou des rapports CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine. Par défaut le client GDAP (`load.config.ps1`) ou votre propre tenant ; obligatoire en app-only |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`). Sans lui, le script se connecte en délégué, en votre nom |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only avec `-ClientId` |
| `-AppOnly` | Non | Connexion app-only avec le ClientId et l'empreinte du tenant depuis `graph.appid.json` |

**Exemples**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail
```

**Remarques**
- Le suivi des messages V2 remonte à 90 jours ; un `-StartDate` plus ancien est ramené à
  90 jours avec un avertissement. Un appel `Get-MessageTraceV2` couvre au plus 10 jours et
  5000 lignes, le script découpe donc la fenêtre en tranches de 10 jours et pagine chacune
  (`EndDate` + `StartingRecipientAddress` de la dernière ligne) — les versions précédentes
  s'arrêtaient à 5000 lignes sans le signaler
- Les cmdlets retirées `Get-MessageTrace` / `Get-MessageTraceDetail` ne sont plus utilisées ;
  le script s'arrête avec un message clair si la session n'a pas les cmdlets V2 (mettez à
  jour ExchangeOnlineManagement)
- Connexion via [`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1) : Exchange
  Online, en délégué par défaut (code d'appareil et client GDAP via `-DelegatedOrganization`,
  selon `load.config.ps1`) ; app-only avec `-ClientId` + `-CertificateThumbprint` ou
  `-AppOnly` (nécessite `-TenantId` sous forme de domaine). Reste sur Exchange Online car
  Microsoft Graph n'a pas d'API de suivi des messages

**Module requis**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
