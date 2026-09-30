[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [PatronToolkit](../readme.fr.md) › **Entra**

# Patron Toolkit — Entra

Rapports d'inscription MFA/SSPR et sauvegarde des stratégies Conditional Access via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MfaRegistrationReport.ps1`](Get-MfaRegistrationReport.ps1) ([docs](#get-mfaregistrationreportps1)) | Rapport de l'état d'inscription MFA/SSPR pour tous les utilisateurs ou une sélection |
| [`Export-ConditionalAccessPolicies.ps1`](Export-ConditionalAccessPolicies.ps1) ([docs](#export-conditionalaccesspoliciesps1)) | Sauvegarder toutes les stratégies Conditional Access et les emplacements nommés en JSON/CSV |

---

### Get-MfaRegistrationReport.ps1

Rapporte l'état d'inscription MFA et SSPR des utilisateurs via le rapport des détails d'inscription
aux méthodes d'authentification de Microsoft Graph. Signale les utilisateurs non inscrits à la MFA, en mettant
à part les comptes administrateurs, car un compte administrateur non inscrit est la constatation la plus
prioritaire.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserList` | Non | Un ou plusieurs UPN sur lesquels produire le rapport. Si omis, tous les utilisateurs sont inclus |
| `-AdminsOnly` | Non | Ne rapporter que les titulaires d'un rôle d'annuaire |
| `-NotRegisteredOnly` | Non | N'inclure que les utilisateurs non inscrits à la MFA |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
# Rapport complet du tenant
.\Get-MfaRegistrationReport.ps1

# Uniquement les utilisateurs qui doivent encore s'inscrire
.\Get-MfaRegistrationReport.ps1 -NotRegisteredOnly

# Administrateurs non inscrits à la MFA — la constatation la plus prioritaire
.\Get-MfaRegistrationReport.ps1 -AdminsOnly -NotRegisteredOnly
```

**Remarques**
- Nécessite une licence Entra ID P1/P2 (le rapport sous-jacent est une fonctionnalité Premium)
- Étendue requise : `Reports.Read.All` (ou `AuditLog.Read.All`)

---

### Export-ConditionalAccessPolicies.ps1

Sauvegarde chaque stratégie Conditional Access et chaque emplacement nommé actuellement configurés dans le
tenant en JSON (un fichier par stratégie, plus un instantané combiné) et dans un récapitulatif CSV
aplati — un outil de sauvegarde ponctuelle et de suivi des modifications, distinct de
[`Import-ConditionalAccessBaseline.ps1`](../../Entra/readme.fr.md), qui importe une base de référence
communautaire précise. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-OutputPath` | Non | Dossier de sauvegarde (par défaut : `.\CAPolicyBackup_<timestamp>\` sous `C:\Temp\` / `~/Downloads\`) |
| `-IncludeNamedLocations` | Non | Exporter aussi les emplacements nommés (par défaut : activé) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Export-ConditionalAccessPolicies.ps1

.\Export-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\ContosoCA"
```

**Remarques**
- La réimportation générique de stratégies CA depuis le JSON n'a volontairement pas été développée — considérez le JSON comme
  un artefact de sauvegarde/comparaison, pas comme un format importable ; utilisez
  `scripts/Entra/Import-ConditionalAccessBaseline.ps1` pour un flux d'import maintenu
- Étendue requise : `Policy.Read.All`

**Module requis**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
