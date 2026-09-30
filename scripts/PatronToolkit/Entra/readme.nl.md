[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **Entra**

# Patron Toolkit — Entra

Rapportage van MFA/SSPR-registratie en back-up van Conditional Access-beleid via Microsoft Graph.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-MfaRegistrationReport.ps1`](Get-MfaRegistrationReport.ps1) ([docs](#get-mfaregistrationreportps1)) | Rapport van de MFA/SSPR-registratiestatus voor alle of geselecteerde gebruikers |
| [`Export-ConditionalAccessPolicies.ps1`](Export-ConditionalAccessPolicies.ps1) ([docs](#export-conditionalaccesspoliciesps1)) | Back-up van al het Conditional Access-beleid en alle benoemde locaties naar JSON/CSV |

---

### Get-MfaRegistrationReport.ps1

Rapporteert de MFA- en SSPR-registratiestatus van gebruikers via het rapport met registratiegegevens
van authenticatiemethoden in Microsoft Graph. Markeert gebruikers die niet voor MFA zijn geregistreerd en noemt
beheerdersaccounts apart, omdat een niet-geregistreerd beheerdersaccount de bevinding met de hoogste
prioriteit is.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserList` | Nee | Een of meer UPN's om over te rapporteren. Zonder deze parameter worden alle gebruikers gerapporteerd |
| `-AdminsOnly` | Nee | Alleen rapporteren over houders van een directoryrol |
| `-NotRegisteredOnly` | Nee | Alleen gebruikers opnemen die niet voor MFA zijn geregistreerd |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
# Volledig tenantrapport
.\Get-MfaRegistrationReport.ps1

# Alleen gebruikers die zich nog moeten registreren
.\Get-MfaRegistrationReport.ps1 -NotRegisteredOnly

# Beheerders die niet voor MFA zijn geregistreerd — de bevinding met de hoogste prioriteit
.\Get-MfaRegistrationReport.ps1 -AdminsOnly -NotRegisteredOnly
```

**Opmerkingen**
- Vereist een Entra ID P1/P2-licentie (het onderliggende rapport is een Premium-functie)
- Vereiste scope: `Reports.Read.All` (of `AuditLog.Read.All`)

---

### Export-ConditionalAccessPolicies.ps1

Maakt een back-up van elk Conditional Access-beleid en elke benoemde locatie die op dat moment in de
tenant is geconfigureerd, naar JSON (één bestand per beleid, plus een gecombineerde momentopname) en een platgeslagen CSV-
overzicht — een hulpmiddel voor een momentopname-back-up en het bijhouden van wijzigingen, los van
[`Import-ConditionalAccessBaseline.ps1`](../../Entra/readme.nl.md), dat een specifieke
community-basislijn importeert. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-OutputPath` | Nee | Back-upmap (standaard: `.\CAPolicyBackup_<timestamp>\` onder `C:\Temp\` / `~/Downloads\`) |
| `-IncludeNamedLocations` | Nee | Ook benoemde locaties exporteren (standaard: aan) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Export-ConditionalAccessPolicies.ps1

.\Export-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\ContosoCA"
```

**Opmerkingen**
- Een generieke herimport van CA-beleid uit JSON is bewust niet gebouwd — behandel de JSON als
  back-up/vergelijkingsmateriaal, niet als importeerbaar formaat; gebruik
  `scripts/Entra/Import-ConditionalAccessBaseline.ps1` voor een onderhouden importflow
- Vereiste scope: `Policy.Read.All`

**Vereiste module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
