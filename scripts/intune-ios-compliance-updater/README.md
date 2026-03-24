# Intune iOS Compliance Updater

Automatisch de minimum iOS versie in een Intune compliance policy bijhouden via Microsoft Graph API.

Ontwikkeld door **BraveHub** als onderdeel van het managed device management platform.

---

## Wat doet dit script?

1. Haalt de laatste iOS versie op via Apple's RSS feed
2. Vergelijkt deze met de huidige minimum versie in de Intune compliance policy
3. Updatet de compliance policy automatisch als er een nieuwere versie beschikbaar is
4. Logt alle acties naar een logbestand

Draait als scheduled task op een Windows server — volledig automatisch, geen handmatig werk.

---

## Vereisten

- Windows Server met PowerShell 5.1 of hoger
- Internettoegang naar:
  - `login.microsoftonline.com`
  - `graph.microsoft.com`
  - `developer.apple.com`
- Een **App Registration** in Entra ID met de juiste permissies (zie hieronder)

---

## Installatie

### Optie A: Automatisch (aanbevolen) ⭐

Zet de scripts op de server en voer één commando uit. Het setup script doet alles automatisch:

```powershell
.\Setup.ps1
```

Het script:
- Installeert de benodigde PowerShell modules
- Laat je aanmelden via een browservenster (Entra ID)
- Maakt de App Registration aan
- Wijst API permissies toe en geeft admin consent
- Maakt een Client Secret aan
- Toont een lijst van iOS compliance policies om uit te kiezen
- Schrijft automatisch `config.json`

Optioneel kun je de policy naam meegeven:
```powershell
.\Setup.ps1 -CompliancePolicyName "iOS - Minimum versie compliance"
```

> ⚠️ Vereiste rol: **Global Administrator** of **Application Administrator + Intune Administrator**

---

### Optie B: Handmatig

#### Stap 1: App Registration aanmaken in Entra ID

1. Ga naar [Entra ID](https://entra.microsoft.com) → **App registrations** → **New registration**
2. Naam: `Intune iOS Compliance Updater`
3. Na aanmaken → **API permissions** → **Add a permission**
4. Kies **Microsoft Graph** → **Application permissions**
5. Zoek en voeg toe: `DeviceManagementConfiguration.ReadWrite.All`
6. Klik op **Grant admin consent**
7. Ga naar **Certificates & secrets** → **New client secret**
8. Noteer de volgende waarden:
   - **Tenant ID** (te vinden via Overview van Entra ID)
   - **Client ID** (Application ID van de App Registration)
   - **Client Secret** (de waarde die je net aanmaakte)

#### Stap 2: Compliance Policy ID opzoeken

1. Ga naar [Intune Admin Center](https://intune.microsoft.com)
2. Navigeer naar **Devices** → **Compliance** → klik op de iOS compliance policy
3. Kopieer het GUID uit de URL:
   ```
   .../deviceCompliancePolicies/XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
   ```

#### Stap 3: Config aanmaken

Kopieer `config.example.json` naar `config.json` en vul de waarden in:

```json
{
    "TenantId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret":       "jouw-client-secret",
    "CompliancePolicyId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

> ⚠️ Voeg `config.json` toe aan `.gitignore` — sla nooit secrets op in Git!

---

### Scheduled task registreren (beide opties)

Open PowerShell **als Administrator** en voer uit:

```powershell
.\Register-ScheduledTask.ps1
```

De task draait voortaan elke **maandag om 07:00**.

---

## Gebruik

### Handmatig uitvoeren

```powershell
.\Update-iOSCompliancePolicy.ps1
```

### Dry-run (geen wijzigingen)

```powershell
.\Update-iOSCompliancePolicy.ps1 -WhatIf
```

### Custom config of log pad

```powershell
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"
```

---

## Logging

Logs worden opgeslagen in de `logs\` map naast het script:

```
logs\
└── compliance-updater-2026-03.log
```

Voorbeeld log output:
```
[2026-03-24 07:00:01] [INFO]    Script started
[2026-03-24 07:00:02] [INFO]    Configuration loaded
[2026-03-24 07:00:03] [INFO]    Access token obtained
[2026-03-24 07:00:04] [SUCCESS] Latest iOS version detected: 26.3.1
[2026-03-24 07:00:05] [INFO]    Current minimum iOS version: 26.2.0
[2026-03-24 07:00:06] [SUCCESS] Compliance policy updated to iOS 26.3.1!
```

---

## Bestandsstructuur

```
intune-ios-compliance-updater/
├── Update-iOSCompliancePolicy.ps1   # Hoofdscript
├── Register-ScheduledTask.ps1       # Task Scheduler installatie
├── config.example.json              # Voorbeeld configuratie
├── .gitignore
└── README.md
```

---

## Beveiliging

- Sla `config.json` **nooit** op in Git (staat in `.gitignore`)
- Overweeg de Client Secret op te slaan in **Windows Credential Manager** of **Azure Key Vault** voor productieomgevingen
- De App Registration heeft enkel de minimaal vereiste permissies

---

## Licentie

MIT — vrij te gebruiken en aan te passen.
