#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Disable Windows Credential Manager's local storage of passwords and credentials.

.DESCRIPTION
    Stops and disables the Credential Manager (VaultSvc) service, preventing
    applications and browsers from persisting saved passwords/credentials to the
    local Windows Credential Vault. Common security-hardening baseline setting for
    managed endpoints.

.PARAMETER Apply
    Actually stop and disable the service. Without this switch, the script only
    reports the current state.

.EXAMPLE
    .\Disable-CredentialManagerVault.ps1

.EXAMPLE
    .\Disable-CredentialManagerVault.ps1 -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $Apply
)

$service = Get-Service -Name VaultSvc -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "  Disable-CredentialManagerVault" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $service) {
    Write-Host "  [WARN] VaultSvc service not found on this system." -ForegroundColor Yellow
    exit 0
}

Write-Host "  Current status     : $($service.Status)"
Write-Host "  Current start type : $($service.StartType)"

if (-not $Apply) {
    Write-Host ""
    Write-Host "  Would stop VaultSvc and set its start type to Disabled." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform this change." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess('VaultSvc', "Stop and disable")) { exit 0 }

Stop-Service -Name VaultSvc -Force -ErrorAction Stop
Set-Service -Name VaultSvc -StartupType Disabled -ErrorAction Stop

Write-Host ""
Write-Host "  [OK]   VaultSvc stopped and disabled." -ForegroundColor Green
Write-Host ""
