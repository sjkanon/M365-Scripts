#Requires -Version 5.1
<#
.SYNOPSIS
    Delete a Workspace 365 environment via the Provisioning API.

.DESCRIPTION
    Modernized rewrite of an old always-interactive delete script: the
    provisioning key and environment name are now parameters (never hardcoded),
    and it follows the repo's dry-run-by-default / -Apply convention instead of
    a one-off y/n prompt.

    This only calls the Workspace 365 Provisioning API — it does not remove the
    associated Entra ID App Registration created by
    New-Workspace365Environment.ps1; remove that separately (via Entra admin
    center or Remove-MgApplication) if it's no longer needed.

.PARAMETER WorkspaceHostname
    Base URL of your Workspace 365 tenant, e.g. "https://yourcompany.workspace365.net".

.PARAMETER ProvisioningKey
    Workspace 365 provisioning key (a GUID). Never hardcode this — pass it in at
    call time.

.PARAMETER EnvironmentName
    Name of the environment to delete.

.PARAMETER Apply
    Actually delete the environment. Without this switch, the script only
    reports what it would do.

.EXAMPLE
    # Preview
    .\Remove-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso"

.EXAMPLE
    .\Remove-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [string] $WorkspaceHostname,

    [Parameter(Mandatory)]
    [guid] $ProvisioningKey,

    [Parameter(Mandatory)]
    [string] $EnvironmentName,

    [switch] $Apply
)

$WorkspaceHostname = $WorkspaceHostname.TrimEnd('/')

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Remove-Workspace365Environment" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Environment : $WorkspaceHostname/$EnvironmentName"
Write-Host ("  Mode        : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would permanently delete environment '$EnvironmentName'." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to delete it." -ForegroundColor Yellow
    exit 0
}

if (-not $PSCmdlet.ShouldProcess("$WorkspaceHostname/$EnvironmentName", "Delete Workspace 365 environment")) { exit 0 }

try {
    Invoke-RestMethod -Method DELETE -Uri "$WorkspaceHostname/Provisioning/Environment/$EnvironmentName/" `
        -ContentType 'application/json' -Headers @{ ProvisioningKey = $ProvisioningKey.ToString() } -ErrorAction Stop | Out-Null
    Write-Host "  [OK]   Environment '$EnvironmentName' deleted." -ForegroundColor Green
} catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "  [ERROR] Deletion failed (HTTP $statusCode): $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
