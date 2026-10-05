[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../../readme.fr.md) › [scripts](../../../../readme.fr.md) › [Custom Scripts](../../../readme.fr.md) › [Intune](../../readme.fr.md) › [Desktop](../readme.fr.md) › **Office Themes**

# Couleurs du thème Office

Déploie une palette de couleurs Office (uniquement les couleurs du thème, pas un thème `.thmx` complet — pour cela, voir [`Deploy-OfficeTheme.ps1`](../readme.fr.md#deploy-officethemeps1) dans le dossier parent).

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Deploy-Officecolors.ps1`](Deploy-Officecolors.ps1) ([docs](#deploy-officecolorsps1)) | Télécharge un XML de jeu de couleurs depuis une URL et l'installe dans le dossier Theme Colors d'Office |

---

### Deploy-Officecolors.ps1

Télécharge la définition du jeu de couleurs (`<a:clrScheme>`) depuis `-ColorsUrl` dans `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\`, en créant le dossier si nécessaire. Une fois déployé, le jeu apparaît dans le sélecteur **Création > Couleurs** (**Design > Colors**) d'Office.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-ColorsUrl` | URL de téléchargement direct du jeu de couleurs `.xml`. Obligatoire |
| `-ColorsName` | Nom de fichier sous lequel l'enregistrer, se terminant par `.xml`. Par défaut : le dernier segment de l'URL |

**Exemples**

```powershell
.\Deploy-Officecolors.ps1 -ColorsUrl 'https://contoso.blob.core.windows.net/branding/Contoso.xml'
```

**Remarques**

- À exécuter en tant qu'utilisateur connecté (écrit dans `%APPDATA%`). Sous Intune, déployez-le comme application Win32 avec les paramètres dans la commande d'installation, ou chargez une copie dont les valeurs par défaut sont renseignées.
- Sans `-ColorsUrl`, ou avec un nom qui ne se termine pas par `.xml`, il s'arrête avec le code de sortie 1.
