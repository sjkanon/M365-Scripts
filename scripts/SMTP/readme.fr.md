[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **SMTP**

# Scripts de test SMTP

> Auteur : Sjoerd Kanon

Scripts de test de la connectivité et de l'authentification SMTP vers Office 365 (ou tout autre serveur SMTP). Utiles pour diagnostiquer les problèmes de relais de messagerie, tester les identifiants de boîtes aux lettres partagées et vérifier la configuration des connecteurs.

---

## Fichiers

| Fichier | Description |
|---|---|
| [`testsmtp.ps1`](testsmtp.ps1) | Test SMTP ponctuel — demande le mot de passe de manière interactive |
| [`testsmtp_5min.ps1`](testsmtp_5min.ps1) | Test récurrent — envoie un message toutes les 5 minutes à l'aide d'un mot de passe enregistré |

---

## testsmtp.ps1

### Quand l'utiliser

Vérification rapide et ponctuelle : ce compte peut-il s'authentifier et envoyer via SMTP ? Demande le mot de passe à l'exécution — rien n'est stocké sur le disque.

### Configuration

Modifiez les variables en haut du script :

```powershell
$SMTPServer = "smtp.office365.com"
$SMTPPort   = 587
$From       = "sender@domain.com"
$To         = "recipient@domain.com"
```

### Utilisation

```powershell
.\testsmtp.ps1
# Demande : Enter SMTP password for sender@domain.com: ****
```

---

## testsmtp_5min.ps1

### Quand l'utiliser

Test de relais prolongé : vérifier que SMTP reste fonctionnel dans la durée, ou reproduire des échecs intermittents. S'exécute toutes les 5 minutes jusqu'à l'arrêt par `Ctrl+C`. Utilise un mot de passe chiffré enregistré, ce qui lui permet de tourner sans surveillance.

### Configuration

Modifiez les variables en haut du script :

```powershell
$SMTPServer      = "smtp.office365.com"
$SMTPPort        = 587
$From            = "sender@domain.com"
$AuthAs          = "sender@domain.com"   # à définir sur le compte d'authentification s'il diffère de $From
$To              = "recipient@domain.com"
$IntervalSeconds = 300                   # 5 minutes
$SavedKeyPath    = "$env:USERPROFILE\smtp_test_password.txt"
```

> Utilisez `$AuthAs` lors d'un envoi depuis une boîte aux lettres partagée : `$From` = adresse de la boîte partagée, `$AuthAs` = le compte utilisateur disposant de l'autorisation Send As.

### Comportement selon la plateforme

| Plateforme | Gestion des identifiants |
|---|---|
| **Windows** | Chargé depuis `$SavedKeyPath` (chiffré par DPAPI). À enregistrer une fois, puis s'exécute sans surveillance. |
| **macOS / Linux** | DPAPI n'est pas disponible — demande le mot de passe une fois au démarrage et le conserve en mémoire pour la session. |

### Première exécution sous Windows — enregistrer le mot de passe

Exécutez ceci **une seule fois** pour enregistrer le mot de passe chiffré sur le disque :

```powershell
Read-Host -AsSecureString "Enter SMTP password" | ConvertFrom-SecureString | Set-Content "$env:USERPROFILE\smtp_test_password.txt"
```

> Le fichier enregistré utilise le chiffrement Windows DPAPI — il ne peut être déchiffré que par le même utilisateur Windows sur la même machine. Sous macOS/Linux, cette étape n'est pas nécessaire.

### Utilisation

```powershell
.\testsmtp_5min.ps1
# Sortie :
# Starting recurring SMTP test — sending every 5 minutes. Press Ctrl+C to stop.
# Server : smtp.office365.com:587
# From   : sender@domain.com  →  To: recipient@domain.com
#
# 2026-03-20 14:00:00 - Email sent successfully to recipient@domain.com
# 2026-03-20 14:05:00 - Email sent successfully to recipient@domain.com
```

---

## Paramètres SMTP courants

| Fournisseur | Serveur | Port | Remarques |
|---|---|---|---|
| Microsoft 365 | `smtp.office365.com` | `587` | STARTTLS, SMTP AUTH doit être activé pour la boîte aux lettres |
| Gmail | `smtp.gmail.com` | `587` | Nécessite un mot de passe d'application si la 2FA est activée |
| Exchange on-premise | `mail.domain.com` | `587` ou `25` | Dépend de la configuration du connecteur |

### Activer SMTP AUTH pour une boîte aux lettres dans M365

SMTP AUTH est désactivé par défaut dans Microsoft 365. Activez-le boîte par boîte :

```powershell
# Connectez-vous d'abord à Exchange Online
Set-CASMailbox -Identity "sender@domain.com" -SmtpClientAuthenticationDisabled $false
```

Ou via le portail d'administration : **Exchange Admin Center → Mailboxes → [mailbox] → Mail flow settings → Authenticated SMTP**

---

## Journal des modifications

| Date | Version | Modification |
|---|---|---|
| 2026-03-20 | 2.1 | Multiplateforme : Windows utilise un fichier DPAPI enregistré, macOS/Linux demande une fois au démarrage |
| 2026-03-20 | 2.0 | Réécrit en anglais ; `System.Net.Mail.SmtpClient` remplace `Send-MailMessage`, obsolète ; adresses codées en dur supprimées ; ajout de `$AuthAs`, `param()`, `#Requires -Version 5.1` |
