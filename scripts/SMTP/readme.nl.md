[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **SMTP**

# SMTP-testscripts

> Auteur: Sjoerd Kanon

Scripts om SMTP-connectiviteit en -authenticatie te testen tegen Office 365 (of elke andere SMTP-server). Handig om problemen met mailrelay te diagnosticeren, inloggegevens van gedeelde mailboxen te testen en de configuratie van connectors te controleren.

---

## Bestanden

| Bestand | Omschrijving |
|---|---|
| [`testsmtp.ps1`](testsmtp.ps1) ([docs](#testsmtpps1)) | Eenmalige SMTP-test — vraagt interactief om het wachtwoord |
| [`testsmtp_5min.ps1`](testsmtp_5min.ps1) ([docs](#testsmtp_5minps1)) | Terugkerende test — verstuurt elke 5 minuten met een opgeslagen wachtwoord |

---

## testsmtp.ps1

### Wanneer gebruiken

Snelle eenmalige controle: kan dit account zich via SMTP authenticeren en versturen? Vraagt het wachtwoord tijdens het uitvoeren — er wordt niets op schijf opgeslagen.

### Configuratie

Pas de variabelen bovenaan het script aan:

```powershell
$SMTPServer = "smtp.office365.com"
$SMTPPort   = 587
$From       = "sender@domain.com"
$To         = "recipient@domain.com"
```

### Gebruik

```powershell
.\testsmtp.ps1
# Vraagt: Enter SMTP password for sender@domain.com: ****
```

---

## testsmtp_5min.ps1

### Wanneer gebruiken

Langdurige relaytest: controleren of SMTP over langere tijd blijft werken, of periodieke storingen reproduceren. Draait elke 5 minuten totdat je het stopt met `Ctrl+C`. Gebruikt een opgeslagen versleuteld wachtwoord, zodat het onbeheerd kan draaien.

### Configuratie

Pas de variabelen bovenaan het script aan:

```powershell
$SMTPServer      = "smtp.office365.com"
$SMTPPort        = 587
$From            = "sender@domain.com"
$AuthAs          = "sender@domain.com"   # zet op het authenticatieaccount als dat afwijkt van $From
$To              = "recipient@domain.com"
$IntervalSeconds = 300                   # 5 minuten
$SavedKeyPath    = "$env:USERPROFILE\smtp_test_password.txt"
```

> Gebruik `$AuthAs` als je vanuit een gedeelde mailbox verstuurt: `$From` = adres van de gedeelde mailbox, `$AuthAs` = het gebruikersaccount dat het recht Send As heeft.

### Gedrag per platform

| Platform | Omgang met inloggegevens |
|---|---|
| **Windows** | Laadt uit `$SavedKeyPath` (DPAPI-versleuteld). Eén keer opslaan, draait daarna onbeheerd. |
| **macOS / Linux** | DPAPI is niet beschikbaar — vraagt bij het starten één keer om het wachtwoord en houdt het voor de sessie in het geheugen. |

### Eerste run op Windows — wachtwoord opslaan

Voer dit **één keer** uit om het versleutelde wachtwoord op schijf op te slaan:

```powershell
Read-Host -AsSecureString "Enter SMTP password" | ConvertFrom-SecureString | Set-Content "$env:USERPROFILE\smtp_test_password.txt"
```

> Het opgeslagen bestand gebruikt Windows DPAPI-versleuteling — het kan alleen worden ontsleuteld door dezelfde Windows-gebruiker op dezelfde machine. Op macOS/Linux is deze stap niet nodig.

### Gebruik

```powershell
.\testsmtp_5min.ps1
# Uitvoer:
# Starting recurring SMTP test — sending every 5 minutes. Press Ctrl+C to stop.
# Server : smtp.office365.com:587
# From   : sender@domain.com  →  To: recipient@domain.com
#
# 2026-03-20 14:00:00 - Email sent successfully to recipient@domain.com
# 2026-03-20 14:05:00 - Email sent successfully to recipient@domain.com
```

---

## Gangbare SMTP-instellingen

| Provider | Server | Poort | Opmerkingen |
|---|---|---|---|
| Microsoft 365 | `smtp.office365.com` | `587` | STARTTLS, SMTP AUTH moet voor de mailbox ingeschakeld zijn |
| Gmail | `smtp.gmail.com` | `587` | Vereist een app-wachtwoord als 2FA is ingeschakeld |
| On-premises Exchange | `mail.domain.com` | `587` of `25` | Afhankelijk van de connectorconfiguratie |

### SMTP AUTH inschakelen voor een mailbox in M365

SMTP AUTH is in Microsoft 365 standaard uitgeschakeld. Schakel het per mailbox in:

```powershell
# Maak eerst verbinding met Exchange Online
Set-CASMailbox -Identity "sender@domain.com" -SmtpClientAuthenticationDisabled $false
```

Of via de beheerportal: **Exchange Admin Center → Mailboxes → [mailbox] → Mail flow settings → Authenticated SMTP**

---

## Wijzigingslog

| Datum | Versie | Wijziging |
|---|---|---|
| 2026-03-20 | 2.1 | Cross-platform: Windows gebruikt een opgeslagen DPAPI-bestand, macOS/Linux vraagt één keer bij het starten |
| 2026-03-20 | 2.0 | Herschreven naar het Engels; `System.Net.Mail.SmtpClient` vervangt het verouderde `Send-MailMessage`; hardcoded adressen verwijderd; `$AuthAs`, `param()` en `#Requires -Version 5.1` toegevoegd |
