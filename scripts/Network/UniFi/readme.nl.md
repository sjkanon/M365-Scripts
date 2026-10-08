[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Network](../readme.nl.md) › **UniFi**

# UniFi

Tooling voor een UniFi Network Controller of UniFi OS-console (UDM/UDM-Pro/UDR). Praat rechtstreeks met de API van de controller — geen onderdeel van [`menu.ps1`](../../../menu.ps1), omdat elke run een controller-URL en inloggegevens nodig heeft. Inloggegevens worden altijd gevraagd via `-Credential`/`Get-Credential`, nooit hardcoded.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-UnifiNetworkReport.ps1`](Get-UnifiNetworkReport.ps1) ([docs](#get-unifinetworkreportps1)) | Een HTML-rapport met netwerkdocumentatie genereren (apparaten, firmware, uptime) |
| [`Update-UnifiFirmware.ps1`](Update-UnifiFirmware.ps1) ([docs](#update-unififirmwareps1)) | Firmware-upgrades over sites heen weergeven en optioneel starten |
| [`UnifiApi.ps1`](UnifiApi.ps1) ([docs](#unifiapips1)) | Gedeelde helper voor aanmelding/sessie, automatisch dot-sourced door de twee scripts hierboven — niet bedoeld om rechtstreeks uit te voeren |

---

### Get-UnifiNetworkReport.ps1

Alleen-lezen. Meldt zich aan, somt elke site op (of één met `-Site`), toont de geadopteerde apparaten per site en schrijft een HTML-rapport (naam, model, MAC, IP, firmwareversie, status, uptime) naar `C:\Temp\`.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Controller` | Ja | Basis-URL van de controller, bijv. `https://192.168.1.1` of `https://unifi.contoso.local:8443` |
| `-Credential` | Nee | Admin-inloggegevens — worden via `Get-Credential` gevraagd als je ze weglaat |
| `-Site` | Nee | Beperken tot één site op naam. Weggelaten: alle sites worden meegenomen |
| `-SkipCertificateCheck` | Nee | Zelfondertekende/niet-vertrouwde certificaten accepteren (gebruikelijk bij on-prem controllers) |
| `-OutputPath` | Nee | Map voor het rapport (standaard: `C:\Temp\` / `~/Downloads`) |

**Voorbeelden**

```powershell
.\Get-UnifiNetworkReport.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck
.\Get-UnifiNetworkReport.ps1 -Controller "https://unifi.contoso.local:8443" -Site "Head Office"
```

---

### Update-UnifiFirmware.ps1

Toont per site de apparaten waarvoor een firmware-upgrade beschikbaar is (volgens de eigen `upgradable`-vlag van de controller). Draait standaard als proefdraai — zonder `-Apply` wordt geen upgrade gestart. Upgraden herstart het apparaat, wat een korte onderbreking geeft voor alles wat erachter hangt — bevestiging per apparaat is vereist, tenzij je ook `-Force` meegeeft.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Controller` | Ja | Basis-URL van de controller |
| `-Credential` | Nee | Admin-inloggegevens — worden via `Get-Credential` gevraagd als je ze weglaat |
| `-Site` | Nee | Beperken tot één site op naam. Weggelaten: alle sites worden verwerkt |
| `-SkipCertificateCheck` | Nee | Zelfondertekende/niet-vertrouwde certificaten accepteren |
| `-Apply` | Nee | Upgrades daadwerkelijk starten (standaard: proefdraai) |
| `-Force` | Nee | Bevestiging per apparaat overslaan bij gebruik met `-Apply` |
| `-OutputPath` | Nee | Map voor het rapport (standaard: `C:\Temp\` / `~/Downloads`) |

**Voorbeelden**

```powershell
# Proefdraai over alle sites
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck

# Eén site upgraden, bevestigen per apparaat
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Site "Head Office" -Apply

# Onbeheerd, alle sites — voorzichtig gebruiken, apparaten herstarten
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Apply -Force
```

**Opmerkingen**
- Beide scripts ondersteunen de klassieke zelfgehoste UniFi Network Controller (`/api/login`) en UniFi OS-consoles (`/api/auth/login` + `/proxy/network/...`) — bij het aanmelden wordt automatisch gedetecteerd welke van de twee er is.
- `-SkipCertificateCheck` is geïmplementeerd voor zowel PowerShell 7+ (native parameter) als Windows PowerShell 5.1 (tijdelijke callback voor certificaatvalidatie, direct na het verzoek teruggezet).
- Na elke run wordt altijd een CSV-/HTML-rapport geschreven, ook bij een proefdraai.

---

### UnifiApi.ps1

Gedeelde hulpfuncties voor beide scripts hierboven, automatisch dot-sourced — niet bedoeld om los uit te voeren, en het heeft geen parameters. Het regelt de aanmelding (klassieke zelf-gehoste controller en UniFi OS-consoles zoals UDM/UDM-Pro/UDR, die een ander auth-endpoint, een CSRF-header en het pad `/proxy/network/...` gebruiken), de sessiecookies en de afhandeling van zelfondertekende certificaten voor zowel Windows PowerShell 5.1 als PowerShell 7+. Functies: `Connect-UnifiController`, `Disconnect-UnifiController`, `Invoke-UnifiApi`, `Invoke-UnifiRestMethod`, `Get-UnifiSite`, `Get-UnifiDevice`.

Referenties komen altijd uit `Get-Credential` (interactief of van het aanroepende script) — nooit hardcoded.
