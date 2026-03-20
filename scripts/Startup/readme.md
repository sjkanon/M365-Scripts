# Startup Scripts

Bevat de functiebibliotheken en verbindingslogica voor de EOO M365 managementtoolset.
Onderdeel van de [M365-Scripts](../../readme.md) repository.

---

## Bestanden

| Bestand | Status | Omschrijving |
|---|---|---|
| `functies.ps1` | ✅ Actief | Volledige functiebibliotheek — dot-source dit bestand bij opstarten |
| `f-gettenant2.ps1` | ⛔ Deprecated | Vervangen door `Connect-Tenant` in `functies.ps1` — zie opmerking hieronder |

> **Beveiligingsopmerking `f-gettenant2.ps1`:** dit bestand bevat hardcoded `client_id`, `client_secret` en `tenant_id` in plaintext. Als het ooit in git history heeft gestaan, roteer die credentials direct via **Entra ID → App Registrations → Certificates & secrets**.

---

## functies.ps1

Centrale functiebibliotheek voor multi-tenant M365 beheer via Microsoft Graph en Exchange Online.
Volledig herschreven van MSOnline + AzureAD naar Microsoft Graph.

### Vereisten

| Vereiste | Waarde |
|---|---|
| PowerShell | 7.0 of hoger |
| Modules | Zie `#Requires` bovenaan het bestand |
| Verbinding | Wordt automatisch gestart via `Connect-MgGraph` bij dot-sourcing |

### Opstarten

Dot-source het bestand vanuit je PowerShell profiel of startup script:

```powershell
. "$PSScriptRoot\scripts\Startup\functies.ps1"
```

Stel voor het laden de volgende variabelen in:

```powershell
$upn      = "admin@jouwdomein.nl"   # UPN van de beheerder
$realname = "Sjoerd"                # Weergavenaam voor begroeting (optioneel)
```

### Klant selecteren (CSP)

```powershell
Connect-Tenant -Domain "klant.nl"
# Vult $global:cid en $global:connectmsoldomain
# Daarna gebruiken alle functies automatisch de juiste tenant
```

---

## Functies

### Verbinding

| Functie | Omschrijving |
|---|---|
| `Connect-Tenant` | CSP-klant selecteren op domeinnaam, vult `$cid` en `$connectmsoldomain` |
| `Test-ExoConnection` | Controleert/herstelt Exchange Online verbinding voor de geselecteerde klant |
| `Invoke-Menu` | Interactief menu om modules te laden (EXO, Entra, Teams, Intune) |

### Exchange Online

| Functie | Omschrijving |
|---|---|
| `Enable-CopyOfSentItems` | Zet "kopie van verzonden items" aan voor alle mailboxen |
| `Add-SharedMailboxAccess` | Geeft een gebruiker FullAccess + SendAs op een gedeelde mailbox |
| `Set-MailboxLocale` | Stelt taal (NL) en tijdzone in op alle mailboxen |
| `Add-MailboxAlias` | Voegt een alias toe aan een mailbox |
| `Get-MailboxAliases` | Toont alle SMTP-aliassen per mailbox |
| `Export-DistributionGroups` | Exporteert alle distributiegroepen naar CSV (`%TEMP%\ExportDGs.csv`) |
| `Set-AutoReply` | Stelt een out-of-office bericht in (enabled / disabled / scheduled) |

### Microsoft Entra ID / Graph

| Functie | Omschrijving |
|---|---|
| `Get-TenantAdmins` | Toont alle Global Administrators van de klant-tenant |
| `Add-TenantDomain` | Voegt een domein toe, loopt door het verificatieproces |
| `Get-TenantLicenses` | Toont licentieoverzicht met verbruik en beschikbaarheid |
| `Get-TenantUsers` | Toont alle gebruikers met UPN, naam en licenties |
| `Add-TenantAdmin` | Geeft een gebruiker Global Administrator rechten |
| `Get-EntraApplication` | Zoekt een Enterprise App op naam |
| `Reset-UserPassword` | Reset het wachtwoord van een gebruiker in een CSP-klant-tenant |
| `Export-SignInLogs` | Exporteert inloglogboeken naar CSV (standaard 30 dagen, max 30) |

### EOO Beheeraccount

| Functie | Omschrijving |
|---|---|
| `New-EooAdmin` | Maakt het EOO beheeraccount aan als Global Admin in de klant-tenant |
| `Set-EooAsGroupOwner` | Stelt het EOO beheeraccount in als eigenaar van een groep |
| `Reset-EooPassword` | Reset het wachtwoord van het EOO beheeraccount |

### Navigatie

| Functie | Omschrijving |
|---|---|
| `Set-ImportLocation` | Navigeert naar `$env:import` |
| `Set-ScriptsLocation` | Navigeert naar `$env:ps` |

---

## Modules installeren

Gebruik de bootstrap vanuit de root van de repository:

```powershell
.\scripts\Install-Modules.ps1
```

---

*Onderdeel van M365-Scripts — Sjoerd Kanon*
