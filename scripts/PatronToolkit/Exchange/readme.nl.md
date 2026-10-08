[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **Exchange**

# Patron Toolkit — Exchange

Diagnostiek van de mailflow via Exchange Online.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-MessageTraceReport.ps1`](Get-MessageTraceReport.ps1) ([docs](#get-messagetracereportps1)) | Een message trace-rapport (mailflow) exporteren, optioneel met bezorgdetails per bericht |

---

### Get-MessageTraceReport.ps1

Voert een message trace uit over een opgegeven tijdvenster (standaard: de laatste 48 uur), optioneel
gefilterd op afzender, ontvanger en bezorgstatus, met `Get-MessageTraceV2`. Exporteert een CSV
met een overzicht en, met `-IncludeDetail`, een CSV met bezorgdetails per bericht
(`Get-MessageTraceDetailV2`).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Hours` | Nee | Aantal uren terug vanaf nu voor de trace (standaard: `48`). Genegeerd als `-StartDate` is opgegeven |
| `-StartDate` | Nee | Expliciet begin van het tracevenster |
| `-EndDate` | Nee | Expliciet einde van het tracevenster (standaard: nu) |
| `-SenderAddress` | Nee | Filteren op een specifieke afzender |
| `-RecipientAddress` | Nee | Filteren op een specifieke ontvanger |
| `-Status` | Nee | Filteren op bezorgstatus (bijv. `Delivered`, `Failed`, `Pending`, `Quarantined`) |
| `-IncludeDetail` | Nee | Ook bezorgdetails per bericht exporteren (trager bij grote resultaatsets) |
| `-MaxResults` | Nee | Stoppen met pagineren na dit aantal trace-rijen (standaard: `50000`) |
| `-OutputPath` | Nee | Map voor de CSV-rapport(en) (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail
```

**Opmerkingen**
- Message trace V2 gaat 90 dagen terug; een oudere `-StartDate` wordt met een waarschuwing
  ingekort tot 90 dagen. Eén `Get-MessageTraceV2`-aanroep beslaat hooguit 10 dagen en 5000
  rijen, dus het script deelt het venster op in stukken van 10 dagen en pagineert door elk stuk
  (`EndDate` + `StartingRecipientAddress` uit de laatste rij) — eerdere versies stopten
  zonder melding bij 5000 rijen
- De uitgefaseerde `Get-MessageTrace` / `Get-MessageTraceDetail` worden niet meer gebruikt; het
  script stopt met een duidelijke melding als de sessie geen V2-cmdlets heeft (werk
  ExchangeOnlineManagement bij)
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Exchange
  Online, standaard gedelegeerd (device code en de GDAP-klant via `-DelegatedOrganization`,
  volgens `load.config.ps1`); app-only met `-ClientId` + `-CertificateThumbprint` of
  `-AppOnly` (vereist `-TenantId` als domein). Blijft op Exchange Online omdat Microsoft
  Graph geen API voor message trace heeft

**Vereiste module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
