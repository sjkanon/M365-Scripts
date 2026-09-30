[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **iOS-Compliance-Updater**

# Intune — iOS Compliance Updater

Houdt de vereiste minimale iOS-versie in een Intune-compliancebeleid automatisch actueel via de Microsoft Graph API. Draait als geplande taak op een Windows-server — geen handwerk nodig.

**Hoe het werkt**

1. Haalt de nieuwste iOS-versie op uit de RSS-feed van Apple (met de Apple Support-pagina als terugval)
2. Vergelijkt die met de huidige minimumversie in het Intune-compliancebeleid
3. Werkt het beleid bij als er een nieuwere versie beschikbaar is
4. Logt alle acties naar een maandelijks logbestand

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Update-iOSCompliancePolicy.ps1`](Update-iOSCompliancePolicy.ps1) | Hoofdscript — handmatig of via een geplande taak uit te voeren |
| [`Setup.ps1`](Setup.ps1) | Eenmalige setup — maakt de App Registration aan en schrijft config.json |
| [`Install-ScheduledTask.ps1`](Install-ScheduledTask.ps1) | Registreert de geplande taak in Windows |
| `config.example.json` | Voorbeeldconfiguratiebestand |

---

## Setup

### Optie A — Automatisch (aanbevolen)

Voer `Setup.ps1` één keer uit. Het regelt alles:

```powershell
.\Setup.ps1
```

Het script:
- Installeert de vereiste PowerShell-modules
- Opent een browser voor authenticatie
- Maakt de App Registration aan in Entra ID
- Kent de vereiste API-machtiging toe en verleent beheerderstoestemming
- Maakt een Client Secret aan
- Toont de beschikbare iOS-compliancebeleidsregels om uit te kiezen
- Schrijft `config.json` automatisch weg

```powershell
# Geef de beleidsnaam direct op om de keuzeprompt over te slaan
.\Setup.ps1 -CompliancePolicyName "iOS - Minimum version compliance"
```

> Vereiste rol: **Global Administrator** of **Application Administrator + Intune Administrator**

---

### Optie B — Handmatig

**Stap 1 — App Registration aanmaken**

1. Entra ID → **App registrations** → **New registration**
2. Naam: `Intune iOS Compliance Updater`
3. Na het aanmaken → **API permissions** → **Add a permission**
4. Kies **Microsoft Graph** → **Application permissions**
5. Voeg toe: `DeviceManagementConfiguration.ReadWrite.All`
6. Klik op **Grant admin consent**
7. **Certificates & secrets** → **New client secret** — noteer de waarde

**Stap 2 — De ID van het compliancebeleid opzoeken**

1. Intune Admin Center → **Devices** → **Compliance** → open het iOS-beleid
2. Kopieer de GUID uit de URL: `.../deviceCompliancePolicies/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**Stap 3 — config.json aanmaken**

Kopieer `config.example.json` naar `config.json` en vul de waarden in:

```json
{
    "TenantId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret":       "your-client-secret",
    "CompliancePolicyId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

> Commit `config.json` nooit naar Git — het bevat secrets. Het staat in `.gitignore`.

---

### De geplande taak registreren

Registreer na de setup (bij beide opties) de taak:

```powershell
# Uitvoeren als Administrator
.\Install-ScheduledTask.ps1
```

De taak draait elke **maandag om 07:00** als SYSTEM.

---

## Gebruik

```powershell
# Handmatig uitvoeren
.\Update-iOSCompliancePolicy.ps1

# Proefdraai — er wordt niets gewijzigd
.\Update-iOSCompliancePolicy.ps1 -WhatIf

# Eigen config- of logpad
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"
```

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

- Commit `config.json` nooit naar Git (staat al in `.gitignore`)
- Overweeg in productieomgevingen de Client Secret op te slaan in **Windows Credential Manager** of **Azure Key Vault**
- De App Registration gebruikt alleen de minimaal vereiste machtiging: `DeviceManagementConfiguration.ReadWrite.All`
