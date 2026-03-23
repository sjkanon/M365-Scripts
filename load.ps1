#Requires -Version 5.1
<#
.SYNOPSIS
    Startup loader for M365-Scripts.
    Sets admin credentials and launches the interactive menu.

.DESCRIPTION
    Edit the Configuration section below with your admin UPN and display name,
    then run this file instead of menu.ps1.

    You can also dot-source this file from your PowerShell profile:
        . C:\path\to\M365-Scripts\load.ps1
#>

# ── Configuration ─────────────────────────────────────────────────────────────
$global:upn      = "admin@contoso.com"   # Your admin UPN
$global:realname = "Sjoerd"              # Your display name (optional)
# ─────────────────────────────────────────────────────────────────────────────

& "$PSScriptRoot\menu.ps1"
