# SMTP Test Scripts

> **BraveHub Internal Scripts**
> Author: Sjoerd Kanon

Scripts for testing SMTP connectivity and authentication against Office 365 (or any SMTP server). Useful for diagnosing mail relay issues, testing shared mailbox credentials, and verifying connector configuration.

---

## Files

| File | Description |
|---|---|
| `testsmtp.ps1` | One-time SMTP test — prompts for password interactively |
| `testsmtp_5min.ps1` | Recurring test — sends every 5 minutes using a saved password |

---

## testsmtp.ps1

### When to use

Quick one-off check: can this account authenticate and send via SMTP? Prompts for the password at runtime — nothing is stored on disk.

### Configuration

Edit the variables at the top of the script:

```powershell
$SMTPServer = "smtp.office365.com"
$SMTPPort   = 587
$From       = "sender@domain.com"
$To         = "recipient@domain.com"
```

### Usage

```powershell
.\testsmtp.ps1
# Prompts: Enter SMTP password for sender@domain.com: ****
```

---

## testsmtp_5min.ps1

### When to use

Sustained relay test: verify that SMTP stays functional over time, or reproduce intermittent failures. Runs every 5 minutes until stopped with `Ctrl+C`. Uses a saved encrypted password so it can run unattended.

### Configuration

Edit the variables at the top of the script:

```powershell
$SMTPServer      = "smtp.office365.com"
$SMTPPort        = 587
$From            = "sender@domain.com"
$AuthAs          = "sender@domain.com"   # set to auth account if different from $From
$To              = "recipient@domain.com"
$IntervalSeconds = 300                   # 5 minutes
```

> Use `$AuthAs` when sending from a shared mailbox: `$From` = shared mailbox address, `$AuthAs` = the user account that has Send As permission.

### First run — save password

Run this **once** to save the encrypted password to disk, then remove the line:

```powershell
Read-Host -AsSecureString "Enter SMTP password" | ConvertFrom-SecureString | Set-Content "$env:USERPROFILE\smtp_test_password.txt"
```

> The saved file uses Windows DPAPI encryption — it can only be decrypted by the same Windows user on the same machine.

### Usage

```powershell
.\testsmtp_5min.ps1
# Output:
# Starting recurring SMTP test — sending every 5 minutes. Press Ctrl+C to stop.
# Server : smtp.office365.com:587
# From   : sender@domain.com  →  To: recipient@domain.com
#
# 2026-03-20 14:00:00 - Email sent successfully to recipient@domain.com
# 2026-03-20 14:05:00 - Email sent successfully to recipient@domain.com
```

---

## Common SMTP settings

| Provider | Server | Port | Notes |
|---|---|---|---|
| Microsoft 365 | `smtp.office365.com` | `587` | STARTTLS, SMTP AUTH must be enabled for the mailbox |
| Gmail | `smtp.gmail.com` | `587` | Requires App Password if 2FA is enabled |
| On-premises Exchange | `mail.domain.com` | `587` or `25` | Depends on connector config |

### Enable SMTP AUTH for a mailbox in M365

SMTP AUTH is disabled by default in Microsoft 365. Enable it per mailbox:

```powershell
# Connect to Exchange Online first
Set-CASMailbox -Identity "sender@domain.com" -SmtpClientAuthenticationDisabled $false
```

Or via the admin portal: **Exchange Admin Center → Mailboxes → [mailbox] → Mail flow settings → Authenticated SMTP**

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 2026-03-20 | 2.0 | Rewritten to English; `System.Net.Mail.SmtpClient` replaces deprecated `Send-MailMessage`; removed hardcoded addresses; added `$AuthAs` for shared mailbox support; `#Requires -Version 5.1` |
