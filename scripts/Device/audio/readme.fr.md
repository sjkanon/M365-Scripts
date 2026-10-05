[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Device](../readme.fr.md) › **audio**

# Audio — désactiver le microphone interne

Scripts pour détecter et désactiver le microphone interne des ordinateurs portables Windows sur lesquels les collaborateurs utilisent une application VoIP dans le navigateur.

**Problème :** Windows sélectionne parfois le microphone interne du portable comme entrée audio au lieu du casque connecté, ce qui provoque du bruit de fond dans les numéroteurs web. Teams et les autres applications qui gèrent leur propre audio ne sont pas concernés.

**Solution :** désactiver le microphone interne au niveau du périphérique PnP — le casque n'est jamais touché.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`detect-audiodevices.ps1`](detect-audiodevices.ps1) ([docs](#detect-audiodevicesps1)) | Phase 1 — inventorie tous les périphériques audio et les classe comme internes ou casques |
| [`Disable-internalmic.ps1`](Disable-internalmic.ps1) ([docs](#disable-internalmicps1)) | Phase 2 — désactive le microphone interne d'après les motifs de détection |
| [`Rollback-InternalMic.ps1`](Rollback-InternalMic.ps1) ([docs](#rollback-internalmicps1)) | Urgence — réactive le microphone en cas de problème |

---

## detect-audiodevices.ps1

Établit un inventaire complet de tous les périphériques audio du poste et les classe comme microphone interne ou casque. La sortie peut être écrite dans un champ personnalisé du RMM ou affichée à l'écran.

**Logique de classification**

Les microphones sont classés d'après des motifs de nom :

| Type | Motifs |
|------|----------|
| Interne (à désactiver) | `Array`, `Intel`, `Realtek`, `Synaptics`, `High Definition Audio`, `Camera`, `Webcam`, `Integrated`, `Microfoonmatrix`, `Microphone (` |
| Casque (jamais touché) | `EPOS`, `Jabra`, `Yealink`, `Plantronics`, `Poly`, `Logitech`, `Headset` |

**Déploiement**

À exécuter en tant que SYSTEM avec les droits administrateur. Compatible avec NinjaOne, Datto RMM ou une exécution manuelle.

```powershell
.\detect-audiodevices.ps1
```

---

## Disable-internalmic.ps1

Désactive les microphones internes d'après les motifs issus de l'analyse de détection. La liste blanche des casques est toujours vérifiée en premier — les casques ne sont jamais désactivés.

**Logique par microphone**

```
1. Is the name in the headset safelist?   → YES: skip
2. Does the name match an internal pattern? → NO: skip (unknown device)
3. Is the device already disabled?          → YES: skip
4. Disable via Disable-PnpDevice -InstanceId
```

**Motifs internes désactivés**

```
*Microfoonmatrix*, *Microphone Array*, *Microphone*Realtek*,
*Microphone*High Definition Audio*, *Microfoon*Realtek*,
*Microfoon*Synaptics*, *Microphone (2-*
```

**Liste blanche des casques (jamais touchés)**

```
EPOS, Jabra, Plantronics, Yealink, Poly, Logitech,
Headset, Blackwire, Voyager, Sennheiser, Astro, SteelSeries
```

**Codes de sortie**

| Code | Signification |
|------|---------|
| `0` | Tout a réussi |
| `1` | Un ou plusieurs périphériques n'ont pas pu être désactivés |

> Déployez par phases — testez d'abord sur 5 à 10 appareils avant de déployer sur tous les postes.

---

## Rollback-InternalMic.ps1

Réactive le microphone interne sur un appareil précis. À déployer immédiatement si un utilisateur signale que son audio ne fonctionne plus après le script de désactivation.

```powershell
.\Rollback-InternalMic.ps1
```

> Ne déployez le rollback que sur l'appareil concerné, pas sur tous les appareils à la fois.

---

## Intégration de nouveaux portables

Lorsqu'un nouveau modèle de portable arrive dans l'environnement :

1. Exécutez `detect-audiodevices.ps1` pour identifier le nom du microphone interne
2. Vérifiez si ce nom correspond à un motif existant dans `Disable-internalmic.ps1`
3. Si oui : déployez le script de désactivation
4. Si non : ajoutez le nouveau motif à `$internalPatterns` dans le script de désactivation
