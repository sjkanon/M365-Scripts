#Requires -Version 5.1
<#
.SYNOPSIS
    Sets (or clears) the temporary opt-out from Entra ID's automatic passkey
    enablement and Registration Campaign rollout scheduled for Sept 1, 2026.

.DESCRIPTION
    Patches the tenant authentication methods policy:

        PATCH /beta/policies/authenticationmethodspolicy
        { "optOutSettings": { "passkeyDynamicMigration": true } }

    Ref: https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sms-voice-retirement

    NOTE: This only defers the Sept 1, 2026 -> Feb 1, 2027 behavior. There is no
    opt-out from the Feb 1, 2027 enforcement, when Microsoft-provided SMS/voice
    is retired and users with no other method get a blocking passkey prompt.

    Accepts a list of tenants so a GDAP partner can walk every customer tenant in
    one run. Each tenant is handled independently — a tenant that fails to
    connect, read or patch is reported and the run continues with the next one.

.PARAMETER TenantId
    One or more tenant IDs (or domain names). Accepts an array so you can walk a
    list of GDAP customer tenants in one run. Omit to use the currently connected
    context / your default tenant.

.PARAMETER Revert
    Sets passkeyDynamicMigration back to $false (re-opting the tenant IN to the
    automatic migration).

.PARAMETER ReportOnly
    Reads and displays current optOutSettings without changing anything.

.EXAMPLE
    # Show the current setting for the connected tenant
    .\Set-EntraPasskeyMigrationOptOut.ps1 -ReportOnly

.EXAMPLE
    # Preview the change without writing it
    .\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId contoso.onmicrosoft.com -WhatIf

.EXAMPLE
    # Opt out every tenant listed in a text file, no per-tenant confirmation
    .\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId (Get-Content .\tenants.txt) -Confirm:$false

.EXAMPLE
    # Re-opt a tenant back IN to the automatic migration
    .\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId contoso.onmicrosoft.com -Revert

.NOTES
    Requires : Microsoft.Graph.Authentication
    Scope    : Policy.ReadWrite.AuthenticationMethod (Policy.Read.All for -ReportOnly)
    Role     : Authentication Policy Administrator (or Global Administrator)
    Endpoint : /beta  -- this property is beta-only as of Aug 2026.
#>

[CmdletBinding(DefaultParameterSetName = 'Set', SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]] $TenantId,

    [Parameter(ParameterSetName = 'Set')]
    [switch]   $Revert,

    [Parameter(Mandatory = $true, ParameterSetName = 'Report')]
    [switch]   $ReportOnly
)

$ErrorActionPreference = 'Stop'
$PolicyUri  = 'https://graph.microsoft.com/beta/policies/authenticationmethodspolicy'
$TargetVal  = -not $Revert.IsPresent   # $true normally, $false when reverting

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Microsoft.Graph.Authentication is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$Scopes = if ($ReportOnly) { 'Policy.Read.All' } else { 'Policy.ReadWrite.AuthenticationMethod' }

function Get-PasskeyMigrationOptOut {
    <#
        Reads optOutSettings.passkeyDynamicMigration out of a policy response.
        Invoke-MgGraphRequest returns hashtables by default but can be switched
        to PSObject, and the property is absent on tenants that never set it —
        so handle both shapes and return $null for "not set".
    #>
    param($Policy)

    $settings = $null
    if ($Policy -is [System.Collections.IDictionary]) {
        if ($Policy.Contains('optOutSettings')) { $settings = $Policy['optOutSettings'] }
    }
    elseif ($Policy -and $Policy.PSObject.Properties['optOutSettings']) {
        $settings = $Policy.optOutSettings
    }
    if ($null -eq $settings) { return $null }

    if ($settings -is [System.Collections.IDictionary]) {
        if ($settings.Contains('passkeyDynamicMigration') -and
            $null -ne $settings['passkeyDynamicMigration']) {
            return [bool] $settings['passkeyDynamicMigration']
        }
        return $null
    }

    $prop = $settings.PSObject.Properties['passkeyDynamicMigration']
    if ($prop -and $null -ne $prop.Value) { return [bool] $prop.Value }
    return $null
}

function Invoke-TenantOptOut {
    param([string] $Tenant)

    $result = [pscustomobject]@{
        Tenant  = $Tenant
        Before  = $null
        After   = $null
        Status  = 'Unknown'
        Message = ''
    }

    # ---- Connect ----------------------------------------------------------
    # Reuse an existing session when no explicit tenant was asked for:
    # Connect-MgGraph would otherwise force a fresh interactive prompt.
    $ctx = try { Get-MgContext } catch { $null }
    $needConnect = $true
    if (-not $Tenant -and $ctx -and $ctx.Scopes -contains $Scopes) { $needConnect = $false }

    if ($needConnect) {
        $connectArgs = @{ Scopes = $Scopes; NoWelcome = $true }
        if ($Tenant) { $connectArgs['TenantId'] = $Tenant }

        try {
            Connect-MgGraph @connectArgs
            $ctx = Get-MgContext
        }
        catch {
            Write-Warning "[$Tenant] Connect-MgGraph failed: $($_.Exception.Message)"
            $result.Status  = 'ConnectFailed'
            $result.Message = $_.Exception.Message
            return $result
        }
    }

    $label = if ($ctx -and $ctx.TenantId) { $ctx.TenantId } else { $Tenant }
    $result.Tenant = $label

    # ---- Read current state ----------------------------------------------
    try {
        $policy = Invoke-MgGraphRequest -Method GET -Uri $PolicyUri
    }
    catch {
        Write-Warning "[$label] Could not read authentication methods policy: $($_.Exception.Message)"
        $result.Status  = 'ReadFailed'
        $result.Message = $_.Exception.Message
        return $result
    }

    $current       = Get-PasskeyMigrationOptOut -Policy $policy
    $result.Before = $current
    $result.After  = $current

    $shown = if ($null -eq $current) { '<not set>' } else { $current }
    Write-Host "[$label] current passkeyDynamicMigration = $shown" -ForegroundColor Cyan

    if ($ReportOnly) {
        $result.Status = 'Reported'
        return $result
    }

    if ($current -eq $TargetVal) {
        Write-Host "[$label] Already set to $TargetVal - nothing to do." -ForegroundColor DarkGray
        $result.Status = 'AlreadySet'
        return $result
    }

    # ---- Patch ------------------------------------------------------------
    $action = if ($TargetVal) { 'opt OUT of automatic passkey enablement' }
              else            { 'opt back IN to automatic passkey enablement' }

    if (-not $PSCmdlet.ShouldProcess($label, $action)) {
        $result.Status = 'Skipped'
        return $result
    }

    $body = @{ optOutSettings = @{ passkeyDynamicMigration = $TargetVal } } |
            ConvertTo-Json -Depth 5

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $PolicyUri `
                              -Body $body -ContentType 'application/json' | Out-Null
    }
    catch {
        Write-Warning "[$label] PATCH failed: $($_.Exception.Message)"
        $result.Status  = 'PatchFailed'
        $result.Message = $_.Exception.Message
        return $result
    }

    # ---- Verify -----------------------------------------------------------
    Start-Sleep -Seconds 2
    try {
        $verify = Invoke-MgGraphRequest -Method GET -Uri $PolicyUri
        $after  = Get-PasskeyMigrationOptOut -Policy $verify
    }
    catch {
        Write-Warning "[$label] PATCH succeeded but read-back failed: $($_.Exception.Message)"
        $result.Status  = 'PatchedUnverified'
        $result.Message = $_.Exception.Message
        return $result
    }

    $result.After = $after
    if ($after -eq $TargetVal) {
        Write-Host "[$label] Verified: passkeyDynamicMigration = $after" -ForegroundColor Green
        $result.Status = 'Patched'
    }
    else {
        Write-Warning "[$label] PATCH returned success but read-back shows '$after'. Re-check manually."
        $result.Status  = 'PatchedUnverified'
        $result.Message = "Read-back returned '$after'"
    }

    return $result
}

# ---------------------------------------------------------------------------
$results = @()

if ($TenantId) {
    foreach ($t in $TenantId) {
        # One bad tenant must not abort the rest of the list.
        try {
            $results += Invoke-TenantOptOut -Tenant $t
        }
        catch {
            Write-Warning "[$t] Unhandled error: $($_.Exception.Message)"
            $results += [pscustomobject]@{
                Tenant = $t; Before = $null; After = $null
                Status = 'Failed'; Message = $_.Exception.Message
            }
        }
        finally {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
}
else {
    $results += Invoke-TenantOptOut
}

if ($results.Count -gt 1) {
    Write-Host "`nSummary" -ForegroundColor Cyan
    $results | Format-Table Tenant, Before, After, Status, Message -AutoSize | Out-Host
}

Write-Host "`nReminder: this defers only the Sept 1, 2026 auto-enablement." -ForegroundColor Yellow
Write-Host "Feb 1, 2027 SMS/voice retirement is enforced regardless of this setting." -ForegroundColor Yellow

$results
