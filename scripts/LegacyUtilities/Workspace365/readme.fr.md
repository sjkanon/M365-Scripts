[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Workspace365**

# Legacy Utilities — Workspace 365

Scripts de provisionnement pour les environnements de poste de travail numérique
[Workspace 365](https://workspace365.net). Réécriture modernisée d'un ancien script interactif
de plus de 600 lignes : la connexion par code d'appareil via MSAL.PS et les appels Graph bruts
avec jeton bearer ont été remplacés par `Connect-MgGraph` / `Invoke-MgGraphRequest` (en
réutilisant la session existante si vous êtes déjà connecté), et les invites systématiquement
interactives ont été transformées en paramètres selon le modèle de ce dépôt : essai à blanc par
défaut, `-Apply` pour exécuter réellement. La clé de provisionnement et le nom d'hôte sont
toujours fournis par l'appelant ; l'original les enregistrait dans un fichier `.cfg` local en
texte clair à côté du script, et cette persistance a été volontairement supprimée.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-Workspace365Environment.ps1`](New-Workspace365Environment.ps1) ([docs](#new-workspace365environmentps1)) | Provisionner un nouvel environnement, son inscription d'application SSO et les liens Exchange/SharePoint par défaut |
| [`Remove-Workspace365Environment.ps1`](Remove-Workspace365Environment.ps1) ([docs](#remove-workspace365environmentps1)) | Supprimer un environnement via la Provisioning API |

---

### New-Workspace365Environment.ps1

Crée une App Registration Entra ID pour le SSO Workspace 365 (étendues déléguées Graph/Power BI
+ un secret client), provisionne l'environnement via la Workspace 365 Provisioning API, associe
l'application comme fournisseur d'identité SSO de l'environnement et fait pointer ses URL
Exchange/SharePoint par défaut vers ce tenant. Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-WorkspaceHostname` | Oui | par ex. `https://yourcompany.workspace365.net` |
| `-ProvisioningKey` | Oui | Clé de provisionnement Workspace 365 (GUID) ; ne jamais la coder en dur, la passer à l'appel |
| `-EnvironmentName` | Oui | Nom d'environnement alphanumérique en minuscules |
| `-RequestingUserUpn` | Non | Par défaut, l'utilisateur connecté |
| `-Apply` | Non | Provisionner réellement (par défaut : aperçu) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

```powershell
.\New-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

---

### Remove-Workspace365Environment.ps1

Supprime un environnement via la Provisioning API. Ne supprime **pas** l'App Registration
associée : nettoyez-la séparément (centre d'administration Entra ou `Remove-MgApplication`) si
elle n'est plus nécessaire. Essai à blanc par défaut.

```powershell
.\Remove-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

---

**Modules requis**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Remarques**
- En utilisant ces scripts, vous acceptez les
  [conditions générales de Workspace 365](https://workspace365.net/en/term-and-conditions).
