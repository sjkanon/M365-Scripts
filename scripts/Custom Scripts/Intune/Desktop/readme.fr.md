[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../readme.fr.md) › [scripts](../../../readme.fr.md) › [Custom Scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **Desktop**

# Desktop (thème Office)

Déploiement du thème et de la palette de couleurs Office via Intune. Conservé volontairement à cet emplacement — les deux scripts ont leur URL de téléchargement codée en dur vers cet emplacement exact du dépôt (branche `main`) ; les déplacer casserait donc le téléchargement jusqu'à ce que les scripts soient mis à jour et redéployés dans Intune.

Pour le déploiement du fond d'écran, de l'écran de verrouillage et du raccourci de la barre des tâches, voir [`scripts/Intune/Desktop/`](../../../Intune/Desktop/readme.fr.md).

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Office Themes/`](Office%20Themes/readme.fr.md) | `Deploy-Officecolors.ps1` — installe uniquement le jeu de couleurs |

## Scripts

| Script | Description |
|--------|-------------|
| [`Deploy-OfficeTheme.ps1`](Deploy-OfficeTheme.ps1) ([docs](#deploy-officethemeps1)) | Installe le thème Office `.thmx` complet de VIAS Institute |
| `2026 Vias institute colours (2).thmx` | Le fichier de thème Office téléchargé par `Deploy-OfficeTheme.ps1` |

---

### Deploy-OfficeTheme.ps1

Télécharge `2026 Vias institute colours (2).thmx` depuis la branche `main` de ce dépôt sur GitHub et le copie dans `%APPDATA%\Microsoft\Templates\Document Themes\`, afin qu'il apparaisse dans le sélecteur **Création > Thèmes** (**Design > Themes**) d'Office.

```powershell
.\Deploy-OfficeTheme.ps1
```

> Aucun paramètre — l'URL source et le nom de fichier sont codés en dur en haut du script. Déployez-le via Intune en tant qu'utilisateur connecté (il écrit dans `%APPDATA%`).
