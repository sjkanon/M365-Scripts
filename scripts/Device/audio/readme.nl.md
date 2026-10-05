[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Device](../readme.nl.md) › **audio**

# Audio — interne microfoon uitschakelen

Scripts om de interne microfoon te detecteren en uit te schakelen op Windows-laptops waarop medewerkers een browsergebaseerde VoIP-applicatie gebruiken.

**Probleem:** Windows kiest soms de interne laptopmicrofoon als audio-invoer in plaats van de aangesloten headset, wat achtergrondgeluid geeft in browsergebaseerde dialers. Teams en andere apps met hun eigen audiobeheer hebben hier geen last van.

**Oplossing:** schakel de interne microfoon uit op PnP-apparaatniveau — de headset wordt nooit aangeraakt.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`detect-audiodevices.ps1`](detect-audiodevices.ps1) ([docs](#detect-audiodevicesps1)) | Fase 1 — inventariseert alle audioapparaten en classificeert ze als intern of headset |
| [`Disable-internalmic.ps1`](Disable-internalmic.ps1) ([docs](#disable-internalmicps1)) | Fase 2 — schakelt de interne microfoon uit op basis van de detectiepatronen |
| [`Rollback-InternalMic.ps1`](Rollback-InternalMic.ps1) ([docs](#rollback-internalmicps1)) | Noodgeval — schakelt de microfoon weer in als er iets misgaat |

---

## detect-audiodevices.ps1

Maakt een volledige inventaris van alle audioapparaten op het endpoint en classificeert ze als interne microfoon of headset. De uitvoer kan naar een custom field in de RMM worden geschreven of op het scherm worden getoond.

**Classificatielogica**

Microfoons worden geclassificeerd op basis van naampatronen:

| Type | Patronen |
|------|----------|
| Intern (uitschakelen) | `Array`, `Intel`, `Realtek`, `Synaptics`, `High Definition Audio`, `Camera`, `Webcam`, `Integrated`, `Microfoonmatrix`, `Microphone (` |
| Headset (nooit aanraken) | `EPOS`, `Jabra`, `Yealink`, `Plantronics`, `Poly`, `Logitech`, `Headset` |

**Uitrol**

Draai als SYSTEM met administratorrechten. Werkt met NinjaOne, Datto RMM of handmatige uitvoering.

```powershell
.\detect-audiodevices.ps1
```

---

## Disable-internalmic.ps1

Schakelt interne microfoons uit op basis van de patronen uit de detectieanalyse. De headset-safelist wordt altijd eerst gecontroleerd — headsets worden nooit uitgeschakeld.

**Logica per microfoon**

```
1. Is the name in the headset safelist?   → YES: skip
2. Does the name match an internal pattern? → NO: skip (unknown device)
3. Is the device already disabled?          → YES: skip
4. Disable via Disable-PnpDevice -InstanceId
```

**Interne patronen die worden uitgeschakeld**

```
*Microfoonmatrix*, *Microphone Array*, *Microphone*Realtek*,
*Microphone*High Definition Audio*, *Microfoon*Realtek*,
*Microfoon*Synaptics*, *Microphone (2-*
```

**Headset-safelist (nooit aangeraakt)**

```
EPOS, Jabra, Plantronics, Yealink, Poly, Logitech,
Headset, Blackwire, Voyager, Sennheiser, Astro, SteelSeries
```

**Exitcodes**

| Code | Betekenis |
|------|---------|
| `0` | Alles gelukt |
| `1` | Een of meer apparaten konden niet worden uitgeschakeld |

> Rol gefaseerd uit — test eerst op 5–10 apparaten voordat je naar alle endpoints uitrolt.

---

## Rollback-InternalMic.ps1

Schakelt de interne microfoon weer in op een specifiek apparaat. Rol dit direct uit als een gebruiker meldt dat zijn audio niet meer werkt na het uitschakelscript.

```powershell
.\Rollback-InternalMic.ps1
```

> Rol de rollback alleen uit op het getroffen apparaat, niet op alle apparaten tegelijk.

---

## Nieuwe laptops onboarden

Als er een nieuw laptopmodel aan de omgeving wordt toegevoegd:

1. Draai `detect-audiodevices.ps1` om de naam van de interne microfoon te achterhalen
2. Controleer of de naam overeenkomt met een bestaand patroon in `Disable-internalmic.ps1`
3. Zo ja: rol het uitschakelscript uit
4. Zo nee: voeg het nieuwe patroon toe aan `$internalPatterns` in het uitschakelscript
