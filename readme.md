# M365-Scripts

> A collection of PowerShell scripts for managing Microsoft 365 environments, maintained by [Sjoerd Kanon].

---

## Overview

This repository contains production-ready PowerShell scripts used by BraveHub engineers to automate, manage, and report on Microsoft 365 tenants. Scripts are organized by workload and are actively maintained and expanded over time.

---

## Categories

### 📧 Exchange Online
Scripts for mailbox management, permissions, shared mailboxes, distribution groups, and mail flow configuration.

- Mailbox permission management (Full Access, Send As, Send on Behalf)
- Shared mailbox provisioning and access delegation
- Distribution group and mail contact management
- Message trace and mail flow troubleshooting
- Calendar permissions

### 👤 Entra ID (Azure AD)
Scripts for user and group lifecycle management, dynamic groups, and identity governance.

- User provisioning and offboarding
- Dynamic group creation and rule management
- Guest access and B2B management
- Conditional Access policy reporting
- MFA status reporting and enforcement

### 📱 Intune / Autopilot
Scripts for device management, Autopilot enrollment, and compliance reporting.

- Autopilot device registration and profile assignment
- Device compliance status reporting
- App deployment status
- Stale device cleanup

### 📊 Licensing & Reporting
Scripts for license management, usage reporting, and cost optimization.

- License assignment and reconciliation (Pax8 / Ingram Micro)
- Unlicensed user detection
- Service plan assignment per user
- License usage overview per tenant

---

## Requirements

- **PowerShell** 5.1 or later (Windows) / PowerShell 7+ (cross-platform)
- **Required modules** (install as needed per script):

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module AzureAD -Scope CurrentUser          # Legacy, use Graph where possible
Install-Module Microsoft.Graph.Intune -Scope CurrentUser
```

- Appropriate **Microsoft 365 admin permissions** for the target workload
- Script execution must be enabled:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Usage

Each script is self-contained and includes a comment block at the top describing:

- **Purpose** – what the script does
- **Parameters** – required and optional inputs
- **Prerequisites** – required modules and permissions
- **Example** – sample usage

Run any script with `-Help` or check the header comment for usage instructions.

---

## Repository Structure

```
M365-Scripts/
├── ExchangeOnline/
│   ├── Set-MailboxPermissions.ps1
│   └── ...
├── EntraID/
│   ├── New-DynamicGroup.ps1
│   └── ...
├── Intune/
│   ├── Get-DeviceComplianceReport.ps1
│   └── ...
└── Licensing/
    ├── Get-LicenseOverview.ps1
    └── ...
```

---

## Contributing

This repository is actively maintained and expanded. When adding new scripts:

1. Follow the existing naming convention (`Verb-Noun.ps1`)
2. Include a header comment block with Purpose, Parameters, Prerequisites, and Example
3. Test against a non-production tenant before committing
4. Place the script in the appropriate workload folder

---

## Disclaimer

These scripts are provided as-is. Always test in a non-production environment before running against live tenants. BraveHub accepts no liability for unintended changes resulting from misuse or misconfiguration.

---

## Maintainer

**Sjoerd Kanon** — Security-minded | Team- & Projectgericht | Microsoft 365 & Infrastructure