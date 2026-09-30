[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **ActiveDirectory**

# ActiveDirectory

Monitoring van on-prem Active Directory Domain Services — in tegenstelling tot [`Entra/`](../Entra/readme.nl.md), dat zich via Microsoft Graph op de clouddirectory richt.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Watch-ADAccountLockouts.ps1`](Watch-ADAccountLockouts.ps1) ([docs](#watch-adaccountlockoutsps1)) | AD elke 20 minuten controleren op vergrendelde gebruikersaccounts en alleen nieuwe vergrendelingen loggen |

---

### Watch-ADAccountLockouts.ps1

Vraagt de **PDC Emulator** van het domein om de gebruikersaccounts die op dat moment vergrendeld zijn — bewust niet een willekeurige/lokale DC, want vergrendelingstellers en `lockoutTime` worden per DC bijgehouden en zijn alleen op de PDC Emulator gegarandeerd gezaghebbend totdat de replicatie elders is bijgewerkt.

Houdt een klein statusbestand bij (`state.json`, naast het log) zodat alleen een echt **nieuwe** vergrendeling — voor het eerst gezien, of hetzelfde account dat ontgrendeld en opnieuw vergrendeld is — in het log komt. Een account dat over vele runs vergrendeld blijft, wordt niet bij elke run opnieuw gelogd.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-SearchBase` | Optioneel. Beperk tot één OU (distinguished name). Standaard: het hele domein |
| `-LogPath` | Ander logbestand gebruiken (standaard: `C:\ProgramData\ADLockoutMonitor\lockouts.log`) |
| `-RegisterTask` | Dit script registreren als terugkerende geplande taak in plaats van één controle uit te voeren |
| `-TaskIntervalMinutes` | Herhalingsinterval voor `-RegisterTask` (standaard: `20`) |

**Voorbeelden**

```powershell
# Eenmalige controle — hetzelfde wat de geplande taak elke 20 minuten uitvoert
.\Watch-ADAccountLockouts.ps1

# De terugkerende taak één keer registreren, op de DC/fileserver (draait als SYSTEM)
.\Watch-ADAccountLockouts.ps1 -RegisterTask

# Slechts één OU monitoren en in plaats daarvan elke 5 minuten controleren
.\Watch-ADAccountLockouts.ps1 -RegisterTask -TaskIntervalMinutes 5 -SearchBase "OU=Sales,DC=contoso,DC=com"
```

**Opmerkingen**
- Geen meldingen via e-mail/Teams — dit schrijft alleen naar het logbestand. Richt je monitoring-/RMM-tool op `lockouts.log`, of volg het met tail, om gewaarschuwd te worden.
- Voer `-RegisterTask` **één keer** uit — het registreert een geplande taak met de naam `AD Account Lockout Monitor` die daarna ditzelfde script (met jouw `-LogPath`/`-SearchBase`) elke `-TaskIntervalMinutes` opnieuw aanroept, als SYSTEM.
- Vereist de PowerShell-module `ActiveDirectory` (aanwezig op domeincontrollers; installeer op een fileserver de RSAT-feature `RSAT-AD-PowerShell` — `Install-WindowsFeature RSAT-AD-PowerShell`).
