[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **iOS-Compliance-Updater**

# Intune — iOS Compliance Updater

Houdt de vereiste minimale iOS-versie in een Intune-compliancebeleid automatisch actueel via de Microsoft Graph API. Draait als geplande taak op een Windows-server — geen handwerk nodig.

**Hoe het werkt**

1. Haalt de nieuwste uitgebrachte iOS-versie op uit de RSS-feed van Apple (met de Apple Support-pagina als terugval); beta- en RC-items worden overgeslagen en de hoogste versie wint (de feed vermeldt ook updates voor oudere hoofdversies)
2. Vergelijkt die met de huidige minimumversie in het Intune-compliancebeleid
3. Verhoogt het beleid als er een nieuwere versie beschikbaar is (verlaagt het nooit)
4. Logt alle acties naar een maandelijks logbestand

**Aanmelden** gaat via [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1), dus laat deze map in de structuur van de repository staan. Omdat de updater zonder toezicht draait, is **app-only de standaard**: een certificaat (vingerafdruk in `config.json`, privésleutel in het Windows-certificaatarchief). Een `ClientSecret` in een oudere `config.json` werkt nog. `Setup.ps1` meldt je **gedelegeerd** aan om de app te maken. Alle scripts vereisen PowerShell 7.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Update-iOSCompliancePolicy.ps1`](Update-iOSCompliancePolicy.ps1) ([docs](#gebruik)) | Hoofdscript — handmatig of via een geplande taak uit te voeren |
| [`Setup.ps1`](Setup.ps1) ([docs](#optie-a--automatisch-aanbevolen)) | Eenmalige setup — maakt de App Registration en het certificaat aan en schrijft config.json |
| [`Install-ScheduledTask.ps1`](Install-ScheduledTask.ps1) ([docs](#de-geplande-taak-registreren)) | Registreert de geplande taak in Windows (pwsh.exe) |
| `config.example.json` | Voorbeeldconfiguratiebestand |

---

## Setup

### Optie A — Automatisch (aanbevolen)

Voer `Setup.ps1` één keer **als administrator** uit (zodat het certificaat in `LocalMachine\My` komt, waar de SYSTEM-taak het kan gebruiken). Het regelt alles:

```powershell
.\Setup.ps1
```

Het script:
- Installeert de vereiste PowerShell-modules
- Meldt je aan (gedelegeerd, browser of apparaatcode volgens `load.config.ps1`)
- Maakt de App Registration aan in Entra ID (of hergebruikt die)
- Kent de vereiste API-machtiging toe en verleent beheerderstoestemming
- Maakt een zelfondertekend certificaat aan en uploadt de publieke sleutel naar de app (bestaande sleutels van een hergebruikte app blijven staan)
- Toont de beschikbare iOS-compliancebeleidsregels om uit te kiezen (alle pagina's)
- Schrijft `config.json` automatisch weg

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TenantId` | Nee | Tenant om in te richten (standaard: GDAP-klant, anders de tenant waarbij je je aanmeldt) |
| `-CompliancePolicyName` | Nee | Te gebruiken beleid; zonder deze parameter kies je uit een lijst |
| `-AppName` | Nee | Naam van de App Registration (standaard: `Intune iOS Compliance Updater`) |
| `-CredentialType` | Nee | `Certificate` (standaard) of `Secret` |
| `-CertificateStoreLocation` | Nee | `LocalMachine` (standaard als administrator) of `CurrentUser` |
| `-CertificateExpiryYears` | Nee | Geldigheid van het certificaat (standaard: `2`) |
| `-SecretExpiryYears` | Nee | Geldigheid van het secret bij `-CredentialType Secret` (standaard: `2`) |
| `-ConfigPath` | Nee | Waar `config.json` wordt weggeschreven (standaard: map van het script) |

```powershell
# Geef tenant en beleidsnaam direct op om de keuzeprompt over te slaan
.\Setup.ps1 -TenantId "contoso.onmicrosoft.com" -CompliancePolicyName "iOS - Minimum version compliance"

# Oud gedrag: een client secret in config.json
.\Setup.ps1 -CredentialType Secret
```

> Vereiste rol: **Global Administrator**, of **Application Administrator + Privileged Role Administrator** (beheerderstoestemming) **+ Intune Administrator**

---

### Optie B — Handmatig

**Stap 1 — App Registration aanmaken**

1. Entra ID → **App registrations** → **New registration**
2. Naam: `Intune iOS Compliance Updater`
3. Na het aanmaken → **API permissions** → **Add a permission**
4. Kies **Microsoft Graph** → **Application permissions**
5. Voeg toe: `DeviceManagementConfiguration.ReadWrite.All`
6. Klik op **Grant admin consent**
7. **Certificates & secrets** → **Certificates** → upload de `.cer` van een certificaat waarvan de privésleutel in `LocalMachine\My` op de server staat (of, minder veilig, maak een client secret aan)

**Stap 2 — De ID van het compliancebeleid opzoeken**

1. Intune Admin Center → **Devices** → **Compliance** → open het iOS-beleid
2. Kopieer de GUID uit de URL: `.../deviceCompliancePolicies/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**Stap 3 — config.json aanmaken**

Kopieer `config.example.json` naar `config.json` en vul de waarden in:

```json
{
    "TenantId":              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "CertificateThumbprint": "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "CompliancePolicyId":    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

Gebruik bij een client secret `"ClientSecret": "..."` in plaats van `CertificateThumbprint`.

> Commit `config.json` nooit naar Git. Het staat in `.gitignore`.

---

### De geplande taak registreren

Registreer na de setup (bij beide opties) de taak:

```powershell
# Uitvoeren als Administrator
.\Install-ScheduledTask.ps1
```

De taak draait elke **maandag om 07:00** als SYSTEM, met `pwsh.exe` (PowerShell 7 moet geïnstalleerd zijn; de taak gebruikte eerst `powershell.exe`, dat de PowerShell 7-updater niet kan uitvoeren). Voer het opnieuw uit na het bijwerken naar deze versie.

---

## Gebruik

```powershell
# Handmatig uitvoeren
.\Update-iOSCompliancePolicy.ps1

# Proefdraai — er wordt niets gewijzigd
.\Update-iOSCompliancePolicy.ps1 -WhatIf

# Eigen config- of logpad
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"

# Handmatige gedelegeerde run voor één beleid (geen ClientId in config.json)
.\Update-iOSCompliancePolicy.ps1 -CompliancePolicyId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -WhatIf
```

**Parameters** (`Update-iOSCompliancePolicy.ps1`)

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-ConfigPath` | Nee | Configuratiebestand (standaard: `config.json` naast het script) |
| `-LogPath` | Nee | Logmap (standaard: `logs\`) |
| `-CompliancePolicyId` | Nee | Overschrijft `CompliancePolicyId` uit `config.json` |
| `-TenantId` / `-ClientId` / `-CertificateThumbprint` | Nee | Overschrijven de waarden uit `config.json` |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-WhatIf` | Nee | Proefdraai |

**Opmerkingen**
- `-WhatIf` was zowel als scriptparameter als via `SupportsShouldProcess` gedeclareerd, waardoor PowerShell het script helemaal niet startte ("A parameter with the name 'WhatIf' was defined multiple times"). Het is nu alleen nog de gemeenschappelijke parameter.
- De PATCH stuurt nu `@odata.type` (`#microsoft.graph.iosCompliancePolicy`) mee, en het script weigert een beleids-ID die geen iOS-compliancebeleid is.
- Het opzoeken van de versie werkte nooit: `Invoke-RestMethod` geeft de RSS-items zelf terug, dus `$rss.channel.item.title` was altijd leeg, en de terugvalpagina (`111900`) bevatte geen versie meer. Het script leest de items nu rechtstreeks (gecontroleerd tegen de live feed: 27.0.1, terwijl de feed ook 26.6.2 en 18.7.10 vermeldt) en valt terug op "Over iOS-updates" (`support.apple.com/100100`, "The latest version of iOS and iPadOS is …").

---

## Logging

Logs worden weggeschreven naar `logs\compliance-updater-<yyyy-MM>.log` naast het script:

```
[2026-03-24 07:00:01] [INFO   ] Script started
[2026-03-24 07:00:03] [SUCCESS] Latest iOS version: 18.3.2
[2026-03-24 07:00:04] [INFO   ] Policy 'iOS - Minimum version' — current minimum: 18.3.1
[2026-03-24 07:00:05] [SUCCESS] Compliance policy updated to iOS 18.3.2.
```

---

## Beveiliging

- Geef de voorkeur aan het certificaat: `config.json` bevat dan geen secret en de privésleutel is niet exporteerbaar uit het certificaatarchief van de server.
- Een client secret in `config.json` staat er in platte tekst: iedereen die het bestand kan lezen, kan de Intune-configuratie wijzigen. Beperk de ACL van het bestand tot SYSTEM en beheerders, of stap over op een certificaat door `Setup.ps1` opnieuw uit te voeren.
- Commit `config.json` nooit naar Git (staat al in `.gitignore`)
- De App Registration gebruikt alleen de minimaal vereiste machtiging: `DeviceManagementConfiguration.ReadWrite.All`
