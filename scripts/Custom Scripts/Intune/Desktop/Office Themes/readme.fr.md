[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../../readme.fr.md) › [scripts](../../../../readme.fr.md) › [Custom Scripts](../../../readme.fr.md) › [Intune](../../readme.fr.md) › [Desktop](../readme.fr.md) › **Office Themes**

# Couleurs du thème Office

Déploie la palette de couleurs Office de VIAS Institute (uniquement les couleurs du thème, pas le thème `.thmx` complet — pour cela, voir [`Deploy-OfficeTheme.ps1`](../readme.fr.md#deploy-officethemeps1) dans le dossier parent).

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Deploy-Officecolors.ps1`](Deploy-Officecolors.ps1) ([docs](#deploy-officecolorsps1)) | Télécharge le XML du jeu de couleurs et l'installe dans le dossier Theme Colors d'Office |
| `Test VIAS.xml` | La définition du jeu de couleurs (`<a:clrScheme>`) — couleurs sombres/claires/d'accentuation du thème VIAS Institute |

---

### Deploy-Officecolors.ps1

Télécharge `Test VIAS.xml` depuis la branche `main` de ce dépôt sur GitHub et le copie dans `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\`, en créant le dossier si nécessaire. Une fois déployé, « Test VIAS » apparaît comme jeu de couleurs sélectionnable dans le sélecteur **Création > Couleurs** (**Design > Colors**) d'Office.

```powershell
.\Deploy-Officecolors.ps1
```

> Aucun paramètre — l'URL source et le nom de fichier sont codés en dur en haut du script. Déployez-le via Intune en tant qu'utilisateur connecté (il écrit dans `%APPDATA%`).
