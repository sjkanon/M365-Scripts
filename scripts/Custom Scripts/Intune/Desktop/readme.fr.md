[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../readme.fr.md) › [scripts](../../../readme.fr.md) › [Custom Scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **Desktop**

# Desktop (thème Office)

Déploiement du thème et de la palette de couleurs Office via Intune. Les deux scripts reçoivent l'URL de téléchargement du thème en paramètre, si bien qu'un seul script sert pour chaque client ; les fichiers de thème eux-mêmes ne sont pas conservés dans ce dépôt.

Pour le déploiement du fond d'écran, de l'écran de verrouillage et des raccourcis de la barre des tâches, voir [`scripts/Intune/Desktop/`](../../../Intune/Desktop/readme.fr.md).

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Office Themes/`](Office%20Themes/readme.fr.md) | `Deploy-Officecolors.ps1` — installe uniquement un jeu de couleurs |

## Scripts

| Script | Description |
|--------|-------------|
| [`Deploy-OfficeTheme.ps1`](Deploy-OfficeTheme.ps1) ([docs](#deploy-officethemeps1)) | Télécharge un thème Office `.thmx` depuis une URL et l'installe pour l'utilisateur connecté |

---

### Deploy-OfficeTheme.ps1

Télécharge le `.thmx` depuis `-ThemeUrl` dans `%ProgramData%\OfficeThemes` et le copie dans `%APPDATA%\Microsoft\Templates\Document Themes\`, afin qu'il apparaisse dans le sélecteur **Création > Thèmes** (**Design > Themes**) d'Office.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-ThemeUrl` | URL de téléchargement direct du fichier `.thmx`. Obligatoire |
| `-ThemeName` | Nom de fichier sous lequel l'enregistrer, se terminant par `.thmx` — le nom affiché par Office. Par défaut : le dernier segment de l'URL |

**Exemples**

```powershell
.\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://contoso.blob.core.windows.net/branding/Contoso.thmx'

# Commande d'installation d'une application Win32
powershell.exe -ExecutionPolicy Bypass -File .\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://example.com/theme.thmx' -ThemeName 'Contoso 2026.thmx'
```

**Remarques**

- À exécuter en tant qu'utilisateur connecté (écrit dans `%APPDATA%`).
- Les scripts de plateforme Intune ne peuvent pas transmettre de paramètres : déployez-le comme application Win32 avec les paramètres dans la commande d'installation, ou chargez une copie dont les valeurs par défaut sont renseignées.
- Sans `-ThemeUrl`, ou avec un nom qui ne se termine pas par `.thmx`, il s'arrête avec le code de sortie 1 au lieu de demander — sous Intune, personne ne répondrait.
