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
gefilterd op afzender, ontvanger en bezorgstatus. Exporteert een CSV met een overzicht en, met
`-IncludeDetail`, een CSV met bezorgdetails per bericht.

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
| `-OutputPath` | Nee | Map voor de CSV-rapport(en) (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail
```

**Opmerkingen**
- `Get-MessageTrace` gaat maar 10 dagen terug — voor oudere mail gebruik je in plaats daarvan de historische
  zoekfunctie in het Exchange-beheercentrum

**Vereiste module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
