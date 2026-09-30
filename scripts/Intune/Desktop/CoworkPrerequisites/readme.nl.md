[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../readme.nl.md) › [scripts](../../../readme.nl.md) › [Intune](../../readme.nl.md) › [Desktop](../readme.nl.md) › **CoworkPrerequisites**

# CoworkPrerequisites

Machinebrede Intune-uitrol van de vereisten aan de Windows-kant voor [Claude Cowork](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop) — het optionele onderdeel `VirtualMachinePlatform` en het uitschakelen van Windows Snel opstarten — verpakt als **eigen, onafhankelijke** Win32-app, los van [`../ClaudeDesktop/`](../ClaudeDesktop/readme.nl.md).

## Waarom een aparte app in plaats van bundelen in het installatiescript van Claude Desktop

Cowork is optioneel: Claude Desktop werkt prima zonder. Als je de stap voor het Windows-onderdeel in het installatiescript van Claude Desktop zelf opneemt, kan een DISM-hapering die niets met Claude te maken heeft (een drukke servicing stack, Windows Update onbereikbaar als bron voor onderdelen — beide waarschijnlijker op een vers geïmagede machine midden in Autopilot-ESP) de *hele* installatie van Claude Desktop afbreken. Door het af te splitsen:

- Blokkeert een mislukte Cowork-vereiste nooit Claude Desktop zelf.
- Heeft elke app in Intune een eigen, apart zichtbare installatiestatus — je ziet in één oogopslag of het probleem van een apparaat bij "Windows-onderdeel" of bij "Claude Desktop" ligt, in plaats van één ondoorzichtige gecombineerde installatieopdracht.

Er is standaard bewust **geen Intune-"Dependency"** tussen de twee apps ingesteld. Een harde afhankelijkheid zou betekenen dat Claude Desktop niet eens probeert te installeren tot deze app is geslaagd — waarmee precies het probleem hierboven terugkomt. Wijs beide als **Required** toe aan dezelfde groep, los van elkaar.

Behandelt je omgeving Cowork als harde eis in plaats van optioneel, rol dan eerst deze app uit en voer daarna `Deploy-ClaudeDesktopIntune.ps1 -RequireCoworkPrerequisites` uit — zie [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.nl.md#optioneel-cowork-prerequisites-verplicht-stellen) voor wat dat toevoegt en welke afweging daarbij hoort.

## Inhoud

| Script | Rol in Intune |
|---|---|
| [`Deploy-CoworkPrerequisitesIntune.ps1`](Deploy-CoworkPrerequisitesIntune.ps1) | **Het script dat je uitvoert** voor de Win32-app-route. Verpakt de scripts hieronder en maakt de Win32-app aan of werkt die bij. |
| [`Install-CoworkPrerequisites-Intune.ps1`](Install-CoworkPrerequisites-Intune.ps1) | Contentscript voor de installatieopdracht van de Win32-app |
| [`Uninstall-CoworkPrerequisites-Intune.ps1`](Uninstall-CoworkPrerequisites-Intune.ps1) | Contentscript voor de verwijderopdracht van de Win32-app |
| [`Detect-CoworkPrerequisites-Intune.ps1`](Detect-CoworkPrerequisites-Intune.ps1) | Aangepast detectiescript van de Win32-app |
| [`CoworkPrerequisites-PlatformScript.ps1`](CoworkPrerequisites-PlatformScript.ps1) | **Alternatieve**, zelfstandige route — geen Deploy-script, geen verpakking, rechtstreeks geüpload als Intune-"Platform script". Zie "Alternatief: platformscript" hieronder. |

Er is hier geen MSIX — anders dan bij Claude Desktop bestaat de "content" alleen uit deze drie scripts, dus een nieuwe run doet alleen iets als je er echt een hebt aangepast (bijgehouden via een `ScriptsHash` in het veld Notes van de app, hetzelfde patroon als `Deploy-ClaudeDesktopIntune.ps1`).

## Wat het installatiescript doet

1. **VirtualMachinePlatform**: `Enable-WindowsOptionalFeature`, met nieuwe pogingen bij tijdelijke DISM-fouten. Doet niets als het al is ingeschakeld.
2. **Snel opstarten**: zet bij elke run `HiberbootEnabled = 0` onder `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power` (niet alleen wanneer VMP net is ingeschakeld). De eigen Cowork-documentatie van Anthropic waarschuwt expliciet: *"Restart the machine using Restart, not shut down and power on. With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization services uninitialized."* Snel opstarten staat op vrijwel elk Windows-image standaard aan — zonder het uit te schakelen krijgt een gebruiker die gewoon "afsluit" in plaats van "opnieuw opstart" (het gebruikelijke geval) nooit werkende Cowork-services, hoe vaak het apparaat ook uit en aan gaat, terwijl Intune de app als geïnstalleerd toont.
3. Als VMP net is ingeschakeld: sluit af met code **3010** ("soft reboot required"). De app is geconfigureerd met `-RestartBehavior 'basedOnReturnCode'`, dus **Intune zelf** dwingt de herstart af (prompt/deadline/respijtperiode) — dit is niet afhankelijk van een aangemelde gebruiker. Als extraatje stuurt het ook een Engelstalige `msg.exe`-melding naar de actieve consolesessie (herkend aan de onvertaalde `SESSIONNAME` "console", niet aan de taalafhankelijke `STATE`-tekst "Active" — een simpele tekstvergelijking op "Active" zou op een niet-Engelse Windows-weergavetaal stilletjes nooit afgaan), maar de melding is een aardigheidje, niet het mechanisme waarop de herstart echt leunt.

   De melding wordt **niet** verstuurd door `msg.exe` rechtstreeks vanuit dit script in SYSTEM-context aan te roepen — dat levert een dialoogvenster op waarvan de knop OK niet op klikken reageert (een bekende eigenaardigheid van `msg.exe` bij berichten tussen sessies vanaf een niet-interactieve afzender). In plaats daarvan registreert `Show-UserRestartNotification` een kortlevende geplande taak (`LogonType Interactive`, principal = de eigen gebruiker van de consolesessie) die `msg.exe` *binnen de eigen sessie van de gebruiker* uitvoert, en die taak na verzending weer afmeldt — het dialoogvenster hoort dan bij een echt interactief bureaublad en sluit normaal.

## Wat het detectiescript controleert

Meldt alleen "installed" als **zowel** `VirtualMachinePlatform` op `Enabled` staat **als** de onderliggende HCS-services (`vmcompute`, `HNS`, `vfpext`) aanwezig zijn — dat VMP in DISM `Enabled` toont, garandeert niet dat deze services al bestaan (zie de eigen Cowork-probleemoplossing van Anthropic: *"Missing HCS services: HNS, vmcompute, vfpext"*), vooral niet vlak na het inschakelen van VMP maar vóór de vereiste herstart.

Bewust **geen controle op de `Status` van de service** (bijv. `Running`): `vmcompute` is een service met triggerstart en staat naar verwachting op `Stopped` zolang er geen Cowork-sessie actief is — controleren op `Running` zou een volkomen gezond apparaat als "not installed" melden. Alleen een volledig ontbrekende service (`Get-Service` kan hem niet vinden) is een betrouwbaar signaal dat de onderliggende Hyper-V-componenten er nog niet zijn.

## Bekend probleem: `-RequirementRule` op `Set-IntuneWin32App` doet stilletjes niets (opgelost)

Dezelfde bevestigde bug in de module `IntuneWin32App` (1.5.0) als in [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.nl.md#bekend-probleem-installatiefouten-0x80070001--met-een-veel-diepere-bug-in-de-intunewin32app-module-als-oorzaak): de parameter `-RequirementRule` van `Set-IntuneWin32App` eist ten onrechte een eigenschap `@odata.type` die `New-IntuneWin32AppRequirementRule` nooit zet, en de `break` die daarop volgt beëindigt **de hele rest van de functie** — waardoor elke update-aanroep met `-RequirementRule` stilletjes ook `Notes`, `DetectionRule` en `RestartBehavior` oversloeg, niet alleen de architectuurvereiste. `Add-IntuneWin32App` (het aanmaken) heeft deze bug niet.

Op dezelfde manier opgelost: de update-tak geeft `-RequirementRule` niet meer door aan `Set-IntuneWin32App`, en roept in plaats daarvan `Set-Win32AppArchitectureRequirement` aan — een directe Graph PATCH voor `allowedArchitectures`/`minimumSupportedWindowsRelease`, met hergebruik van de sessie die `Connect-MSIntuneGraph` al heeft opgezet.

## Zichtbaarheid in de Bedrijfsportal

`-CompanyPortalFeaturedApp $true` wordt bij elke run gezet (zowel bij het aanmaken als bij updates), zodat de app uitgelicht in de Bedrijfsportal verschijnt in plaats van onzichtbaar te blijven als vereiste die alleen op de achtergrond draait. Er is hier geen MSIX om een logo uit te halen, dus het standaardpictogram voor Win32-apps van Intune wordt gebruikt — stel achteraf in de Intune-portal handmatig `-Icon` in als je een eigen pictogram wilt.

## Maandelijkse run / run wanneer nodig

```powershell
.\Deploy-CoworkPrerequisitesIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

Zelfde gedrag voor bevestiging/`-Force` als `Deploy-ClaudeDesktopIntune.ps1`. Omdat er geen MSIX is waarvan de versie omhooggaat, pusht een nieuwe run alleen een update als je een van de drie contentscripts hebt aangepast — anders doet hij niets.

| Parameter | Standaard | Omschrijving |
|---|---|---|
| `-AssignmentGroupName` | *(verplicht)* | Entra ID-groep die als Required wordt toegewezen. Alleen gebruikt bij het aanmaken. |
| `-WorkingDirectory` | `C:\Temp\CoworkPrereqDeploy` | Staging-map voor de build (`Source/`, `Output/`) |
| `-AppDisplayName` | `Cowork Windows Prerequisites (Machine-wide)` | Wordt gebruikt om de bestaande app bij latere runs te vinden — niet wijzigen zonder de app ook in Intune te hernoemen |
| `-MinimumSupportedWindowsRelease` | `W10_21H2` | Vereistenregel |
| `-TenantId` | automatisch gedetecteerd | Tenant-ID van Entra ID |
| `-IntuneWinAppUtilPath` | automatische download | Gebruik een al gedownloade `IntuneWinAppUtil.exe` |
| `-Force` | uit | Sla de bevestigingsprompt(s) over |

## Verwijderen

`VirtualMachinePlatform` wordt standaard **niet** uitgeschakeld (andere toepassingen — WSL, op Hyper-V gebaseerde tools, andere Cowork-achtige apps — kunnen er ook van afhangen). Geef `-DisableVirtualMachinePlatform` mee aan het verwijderscript als je zeker weet dat niets anders op het apparaat het nodig heeft. Snel opstarten blijft in beide gevallen uitgeschakeld — dat is een onschuldige, apparaatbrede instelling, niet iets dat specifiek bij Cowork hoort.

## Vereisten

- De PowerShell-modules `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups` en `IntuneWin32App` — installeer ze met `.\scripts\Startup\Install-Modules.ps1`
- Uitvoeren vanaf Windows (de verpakkingstool en de DISM-cmdlets werken alleen op Windows)

## Alternatief: platformscript

`CoworkPrerequisites-PlatformScript.ps1` bevat dezelfde logica voor VMP inschakelen en Snel opstarten uitschakelen, aangepast om rechtstreeks als Intune-**Platform script** te worden geüpload (Devices → Scripts and remediations → Platform scripts) in plaats van via de Win32-app-machinerie hierboven. Geen `.intunewin`-verpakking, geen detectie-/vereistenregel, geen `Deploy-*.ps1` — upload gewoon dat ene bestand.

**Waarom dit naast de Win32-app bestaat:** de Win32-app-route liep in de praktijk tegen echte wrijving aan (de `@odata.type`-bug in de module `IntuneWin32App`, GRS-blokkades die het hele apparaat raken, een typfout in `RestartBehavior`) — een gewoon platformscript omzeilt al die Win32-app-specifieke machinerie volledig. De keerzijde is dat je de twee dingen verliest die Win32-apps en Remediations elk boden:

| | Win32-app (`Deploy-CoworkPrerequisitesIntune.ps1` in deze map) | Platformscript (`CoworkPrerequisites-PlatformScript.ps1`) |
|---|---|---|
| Herstart na het inschakelen van VMP | `-RestartBehavior 'basedOnReturnCode'` + exit `3010` → **Intune zelf** dwingt een herstart af (deadline/respijtperiode), geen actie van de gebruiker nodig behalve de prompt | Voor platformscripts bestaat zo'n mechanisme niet — de exitcode is alleen succes(`0`)/mislukt(al het andere), dus dit script sluit altijd af met `0` en stuurt alleen de eenmalige `msg.exe`-melding (via dezelfde truc met een geplande taak in de sessie van de gebruiker als de Win32-app, zodat de knop OK echt werkt); de gebruiker moet volledig op eigen initiatief herstarten |
| Opnieuw controleren na een fout | Detectieregel wordt bij elke check-in opnieuw beoordeeld | Draait standaard één keer per apparaat; een "mislukte" run wordt bij volgende check-ins wel opnieuw geprobeerd, maar een *geslaagde* run (VMP ingeschakeld, herstart nog in afwachting) wordt daarna niet opnieuw gecontroleerd zoals de dagelijkse Detect van een Proactive Remediation dat zou doen |

Proactive Remediations (Detect + Remediate, dagelijks opnieuw uitgevoerd) zouden het beste van beide bieden — maar die functie vereist Windows Enterprise-/Education-licenties of VDA per gebruiker, en dat **zit niet in Business Premium**, dus dat is hier geen optie.

**Uitrollen:**
1. **Devices → Scripts and remediations → Platform scripts → Add → Windows 10 and later**
2. Upload `CoworkPrerequisites-PlatformScript.ps1`
3. Scriptinstellingen: **Run this script using the logged on credentials** = No (SYSTEM) · **Enforce script signature check** = No · **Run script in 64-bit PowerShell Host** = Yes
4. Wijs toe aan dezelfde apparaatgroep als Claude Desktop

Logt naar `%ProgramData%\CoworkPrereqDeploy\platformscript.log` — een andere bestandsnaam dan de `install.log` van de Win32-app-variant, zodat testen van beide op hetzelfde apparaat geen van beide logs overschrijft.
