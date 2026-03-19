# Migrate-HolidaysCalendar.ps1

> **BraveHub Internal Script**
> Ticket: #0298048 | Klant: Onco3R Therapeutics
> Auteur: Sjoerd Kanon | Datum: 19/03/2026

---

## Achtergrond en probleemstelling

Onco3R wilde een gedeelde kalender waarop medewerkers hun verlof kunnen boeken, zodat iedereen een overzicht heeft van wie wanneer afwezig is. De initiële oplossing gebruikte een **Microsoft 365 Group** als gedeelde kalender. Dit werkte technisch, maar had een groot ongewenst neveneffect: **alle groepsleden ontvingen een e-mailnotificatie bij elke nieuwe afspraak** in de kalender. Bij een bedrijfsbrede kalender betekent dit dat iedereen een mail krijgt telkens iemand verlof boekt.

De oplossing is een **Room/Resource Mailbox** — hetzelfde mechanisme als het boeken van een vergaderzaal in Outlook. Medewerkers voegen de resource toe als attendee bij hun verlofafspraak, de boeking wordt automatisch goedgekeurd, en de afspraak verschijnt op de gedeelde kalender. Geen e-mailnotificaties, geen groepslidmaatschap vereist.

### Vergelijking M365 Group vs Room Mailbox

| | M365 Group | Room Mailbox |
|---|---|---|
| Gedeelde kalender | ✅ | ✅ |
| Zichtbaar voor iedereen | ❌ (alleen leden) | ✅ |
| E-mailnotificaties bij events | ❌ (altijd, niet uit te zetten) | ✅ (geen) |
| Werkt als vergaderzaal boeken | ❌ | ✅ |
| AutoAccept verlof | ❌ | ✅ |
| Overlappende boekingen mogelijk | ❌ | ✅ (instelbaar) |

---

## Wat doet het script

Het script voert de volledige migratie uit in één keer:

1. **Platform detecteren** — kiest de juiste authenticatiemethode (Windows vs macOS/Linux)
2. **Exchange Online verbinden** — voor het aanmaken en configureren van de Room Mailbox
3. **App Registration aanmaken** — maakt automatisch een Entra ID app aan met de juiste application permissions (of hergebruikt een bestaande)
4. **Admin consent verlenen** — geeft automatisch consent voor alle benodigde Graph permissions
5. **Graph verbinden (app auth)** — verbindt met client credentials voor schrijftoegang tot andere mailboxen
6. **Room Mailbox aanmaken** — maakt `holidays-calendar@onco3r.com` aan als Room type
7. **Permissies instellen** — stelt Default op Reviewer zodat iedereen de kalender kan lezen
8. **AutoAccept instellen** — verlofboekingen worden automatisch goedgekeurd, overlappen toegestaan
9. **Graph herverbinden (delegated)** — tijdelijk als delegated gebruiker voor het lezen van de groepskalender (Microsoft beperking: groepskalenders zijn niet leesbaar via app auth)
10. **M365 Group opzoeken** — zoekt de Holidays groep op vier manieren (mail lowercase, mail origineel, displayName, Search)
11. **Afspraken ophalen** — haalt alle afspraken op uit de groepskalender binnen het opgegeven tijdsvenster
12. **Terugschakelen naar app auth** — voor het schrijven naar de Room Mailbox
13. **Afspraken kopieren** — kopieert elke afspraak naar de Room Mailbox kalender met Out of Office status
14. **M365 Group verwijderen** — optioneel, verwijdert de M365 Group na migratie
15. **Samenvatting** — toont resultaten en gebruikersinstructies

### Technische opmerking: dual-auth flow

Het script gebruikt bewust **twee Graph verbindingen** tijdens de uitvoering. Dit is nodig omdat Microsoft twee tegenstrijdige beperkingen heeft:

- **Groepskalender lezen** vereist *delegated* access (als ingelogde gebruiker) — app auth wordt geblokkeerd met 403
- **Room Mailbox schrijven** vereist *application* permissions — delegated access geeft 403 op mailboxen van andere gebruikers

Het script schakelt daarom automatisch tussen beide verbindingen op het juiste moment.

---

## Vereisten

### PowerShell versie

PowerShell 7+ is vereist voor macOS en Linux. Op Windows werkt ook PowerShell 5.1.

```powershell
$PSVersionTable.PSVersion  # controleer versie
```

PowerShell 7 installeren: https://aka.ms/powershell

### Modules installeren

```powershell
Install-Module ExchangeOnlineManagement       -Scope CurrentUser
Install-Module Microsoft.Graph.Applications   -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Calendar       -Scope CurrentUser
Install-Module Microsoft.Graph.Groups         -Scope CurrentUser
Install-Module Microsoft.Graph.Users          -Scope CurrentUser
```

Modules updaten indien al geïnstalleerd:

```powershell
Update-Module ExchangeOnlineManagement
Update-Module Microsoft.Graph
```

### Benodigde rechten

De admin die het script uitvoert heeft het volgende nodig:

| Recht | Waarvoor |
|---|---|
| Exchange Admin of Global Admin | Room Mailbox aanmaken, permissies instellen |
| Global Admin | App Registration aanmaken + admin consent verlenen |
| Lid van de Holidays M365 Group | Groepskalender lezen via delegated access |

> **Belangrijk:** de uitvoerende admin moet lid zijn van de Holidays groep. Voeg de admin toe via M365 Admin Center → Groups → Holidays → Members als dat nog niet het geval is.

---

## Platform support

Het script detecteert automatisch het besturingssysteem:

| Platform | Auth methode | Toelichting |
|---|---|---|
| **Windows** | Interactieve browser | Browser opent automatisch |
| **macOS** | Device code flow | Code + URL verschijnt in terminal |
| **Linux** | Device code flow | Code + URL verschijnt in terminal |

Bij device code flow zie je dit in de terminal:

```
To sign in, use a web browser to open the page https://login.microsoft.com/device
and enter the code XXXXXXXXX to authenticate.
```

Open de URL in je browser, voer de code in, en log in met je admin account. Het script detecteert automatisch wanneer je klaar bent en gaat verder.

---

## Gebruik

### Eerste keer — volledig automatisch (Modus A)

Geen `ClientId` of `ClientSecret` opgeven. Het script maakt zelf een App Registration aan.

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId        "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN        "admin@onco3r.onmicrosoft.com" `
    -SourceGroupMail "holidays@onco3r.com"
```

Het script toont aan het einde de `ClientId` en `ClientSecret`. **Sla deze op in Vaultwarden** — het secret wordt maar één keer getoond.

### Volgende keer — bestaande App Registration (Modus B)

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId        "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN        "admin@onco3r.onmicrosoft.com" `
    -ClientId        "717164c5-7927-4066-bdbc-f4ab9f10be56" `
    -ClientSecret    "u2K8Q~JDCMJ~95MilAYfB4o0YMpRU5L2w3H77b0m" `
    -SourceGroupMail "holidays@onco3r.com"
```

### Dry run (geen wijzigingen)

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId        "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN        "admin@onco3r.onmicrosoft.com" `
    -SourceGroupMail "holidays@onco3r.com" `
    -WhatIf
```

### Met verwijderen van de M365 Group na migratie

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId          "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN          "admin@onco3r.onmicrosoft.com" `
    -ClientId          "717164c5-7927-4066-bdbc-f4ab9f10be56" `
    -ClientSecret      "u2K8Q~JDCMJ~95MilAYfB4o0YMpRU5L2w3H77b0m" `
    -SourceGroupMail   "holidays@onco3r.com" `
    -DeleteSourceGroup $true
```

---

## Parameters

| Parameter | Verplicht | Default | Beschrijving |
|---|---|---|---|
| `TenantId` | Ja | — | Azure AD Tenant ID (Entra ID → Overview) |
| `AdminUPN` | Ja | — | UPN van de uitvoerende admin |
| `ClientId` | Nee | `""` | AppId van bestaande App Registration. Leeg = automatisch aanmaken |
| `ClientSecret` | Nee | `""` | Client Secret. Leeg = automatisch aanmaken |
| `AppName` | Nee | `BraveHub-HolidaysCalendarMigration` | Naam van de App Registration |
| `SourceGroupMail` | Nee | `holidays@onco3r.com` | E-mail van de bestaande M365 Group |
| `SourceGroupDisplayName` | Nee | `Holidays` | DisplayName van de M365 Group (fallback) |
| `RoomDisplayName` | Nee | `Holidays Calendar` | Weergavenaam van de nieuwe Room Mailbox |
| `RoomAlias` | Nee | `holidays-calendar` | Alias (moet uniek zijn in de tenant) |
| `RoomEmail` | Nee | `holidays-calendar@onco3r.com` | SMTP-adres van de nieuwe Room Mailbox |
| `DaysBack` | Nee | `365` | Dagen terug voor afspraken ophalen |
| `DaysForward` | Nee | `730` | Dagen vooruit voor afspraken ophalen |
| `DeleteSourceGroup` | Nee | `$false` | M365 Group verwijderen na migratie |

### Tenant ID opzoeken

```powershell
# Via Graph (als je al verbonden bent)
(Get-MgOrganization).Id
```

Of via Azure Portal: **Entra ID → Overview → Tenant ID**

---

## Authenticatie tijdens uitvoering

Afhankelijk van de modus zie je twee of drie login momenten:

| Login | Wanneer | Waarvoor |
|---|---|---|
| Login 1 (delegated) | Altijd bij Modus A | App Registration aanmaken + admin consent |
| Login 2 (delegated) | Altijd | Groepskalender lezen (Microsoft beperking) |
| Login 3 (automatisch) | Altijd | App auth voor Room Mailbox schrijven — geen interactie nodig |

Bij Modus B (bestaande app) vervalt Login 1 en ga je direct naar Login 2.

---

## Na de migratie

### Verlof boeken (eindgebruikers)

1. Maak een afspraak in Outlook
2. Zet de duur op **All day** en de status op **Out of office**
3. Voeg `holidays-calendar@onco3r.com` toe als **attendee** (net als een vergaderzaal)
4. Sla op — de boeking wordt automatisch goedgekeurd
5. De afspraak verschijnt op de gedeelde Holidays Calendar voor iedereen

### Holidays Calendar toevoegen in Outlook (eenmalig per gebruiker)

1. Outlook → Calendar → **Add calendar**
2. Kies **Add from directory**
3. Zoek op `Holidays Calendar` of `holidays-calendar@onco3r.com`
4. Klik **Add** — de kalender verschijnt onder **People's calendars**

---

## Troubleshooting

### Admin is geen lid van de Holidays groep

```
[FAIL] Groep niet gevonden na 4 pogingen.
```

Als de groep wél bestaat maar niet gevonden wordt via delegated access, is de admin waarschijnlijk geen lid. Voeg de admin toe:

**M365 Admin Center → Groups → Active groups → Holidays → Members → Add members**

### Room Mailbox alias conflict

```
New-Mailbox: The alias 'holidays-calendar' is already in use.
```

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -RoomAlias "verlof-kalender" `
    -RoomEmail "verlof-kalender@onco3r.com"
```

### Graph 403 op groepskalender

Dit is een bekende Microsoft beperking — `Get-MgGroupCalendarEvent` werkt niet met application permissions. Het script lost dit op via de dual-auth flow (automatisch). Als je dit toch ziet, controleer of de admin lid is van de groep (zie hierboven).

Referentie: https://learn.microsoft.com/en-us/graph/known-issues#group-calendar

### Graph 403 op Room Mailbox schrijven

Controleer in Entra ID of admin consent correct is verleend:

**Entra ID → App Registrations → BraveHub-HolidaysCalendarMigration → API Permissions**

Alle permissions moeten de status **Granted for Onco3R** tonen. Zo niet, klik **Grant admin consent for Onco3R**.

### Client credentials auth mislukt

```
ClientSecretCredential authentication failed
```

Het script probeert automatisch een fallback via environment variables. Als beide methoden falen, controleer of het secret nog geldig is (vervaldatum staat in de samenvatting). Maak zo nodig een nieuw secret aan:

**Entra ID → App Registrations → BraveHub-HolidaysCalendarMigration → Certificates & secrets → New client secret**

### Afspraken gedeeltelijk mislukt

Afspraken die niet gekopieerd worden, worden gelogd als `[WARN]` met foutmelding. Het script stopt niet bij een fout op een individuele afspraak maar gaat verder. Controleer de `[WARN]` regels in de output na afloop.

---

## Mapstructuur in de repo

```
scripts/
└── m365/
    └── holidays-calendar-migration/
        ├── Migrate-HolidaysCalendar.ps1
        └── README.md
```

---

## Changelog

| Datum | Versie | Wijziging |
|---|---|---|
| 19/03/2026 | 1.0 | Initiële versie |
| 19/03/2026 | 1.1 | Platform detectie (macOS/Linux device code flow) |
| 19/03/2026 | 1.2 | Fallback groep opzoeken op displayName en Search |
| 19/03/2026 | 1.3 | App Registration setup geïntegreerd in hoofdscript |
| 19/03/2026 | 1.4 | Dual-auth flow: delegated lezen + app auth schrijven |
| 19/03/2026 | 1.5 | Fix read-only `$IsWindows`/`$IsMacOS`/`$IsLinux` variabelen |
| 19/03/2026 | 1.6 | Fix `Get-MgGroupCalendarEvent` 403 via module reload tussen verbindingen |