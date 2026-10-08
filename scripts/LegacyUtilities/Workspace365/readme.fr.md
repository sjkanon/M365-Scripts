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
| `-TenantId` | Non | ID ou domaine du tenant Entra ID (par défaut : le client GDAP si `authMode` vaut GDAP) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion app-only avec cette inscription d'application et ce certificat |
| `-AppOnly` | Non | Connexion app-only avec l'application de `graph.appid.json` |

```powershell
.\New-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

**Remarques**
- Se connecte à Graph via [`Connect-M365.ps1`](../../Startup/readme.fr.md) : en délégué par défaut, en app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly` (`-RequestingUserUpn` est alors obligatoire — il n'y a pas d'utilisateur connecté). Une session adaptée est réutilisée ; seule une session ouverte par le script est fermée
- Étendues déléguées : `Application.ReadWrite.All`, `User.Read`, `User.ReadBasic.All` (ajoutée : `-RequestingUserUpn` peut désigner un autre administrateur), `Organization.Read.All`
- L'API de provisionnement elle-même est appelée avec la clé de provisionnement, pas via Graph

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
