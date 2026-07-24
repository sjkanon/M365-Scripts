#Requires -Version 5.1
<#
.SYNOPSIS
    Create a break-glass (emergency access) Global Administrator account in a new tenant.

.DESCRIPTION
    Creates a cloud-only Entra ID user intended as an emergency access / "break-glass"
    account for a newly onboarded tenant, following Microsoft's documented break-glass
    guidance: a dedicated cloud-only account, a long random password stored securely
    outside the tenant (never written to disk by this script), and Global Administrator
    assigned directly (not via a role-assignable group), so it works even if Conditional
    Access or PIM is misconfigured.

    The generated password is written to the console once so it can be captured into a
    password manager / sealed envelope process — it is never saved to a file.

    Default behavior is a dry run — pass -Apply to actually create the account and
    assign the role. Connects to Microsoft Graph automatically if no session is active;
    reuses an existing session if already connected.

.PARAMETER UserPrincipalName
    UPN for the new break-glass account, e.g. "breakglass-admin@contoso.onmicrosoft.com".
    Using the tenant's *.onmicrosoft.com domain (rather than a custom domain) is the
    Microsoft-recommended practice, since it keeps working even if federation/custom
    domain DNS breaks.

.PARAMETER DisplayName
    Display name for the account. Default: "Break Glass Admin".

.PARAMETER PasswordLength
    Length of the generated random password. Default: 24.

.PARAMETER AssignGlobalAdmin
    Assign the Global Administrator directory role to the new account. Default: on.
    Use -AssignGlobalAdmin:$false to create the account without a role assignment.

.PARAMETER ExcludeFromGroupId
    Object ID of a group (e.g. a Conditional Access exclusion group created by
    New-TenantBaselineGroups.ps1) to add the break-glass account to, so it can be
    excluded from Conditional Access policies and MFA enforcement per Microsoft's
    break-glass guidance.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected, or resolvable from a
    GDAP customer tenant context ($global:cid / $env:M365_CUSTOMER_TENANTID).

.PARAMETER Apply
    Actually create the account and assign roles. Without this switch, the script only
    reports what it would do.

.EXAMPLE
    # Preview only
    .\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com"

.EXAMPLE
    .\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

.EXAMPLE
    .\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" `
        -ExcludeFromGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply

.NOTES
    Required scopes : User.ReadWrite.All, RoleManagement.ReadWrite.Directory, GroupMember.ReadWrite.All
    Required module : Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Groups

    Microsoft recommends at least two break-glass accounts per tenant, permanently
    excluded from Conditional Access, with MFA disabled on the account and strong,
    physically/securely stored passwords, monitored for sign-in activity.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $UserPrincipalName,

    [string] $DisplayName = 'Break Glass Admin',
    [int] $PasswordLength = 24,
    [switch] $AssignGlobalAdmin = $true,
    [string] $ExcludeFromGroupId,
    [string] $TenantId,
    [switch] $Apply
)

function New-RandomPassword {
    param([int] $Length = 24)
    $lower   = 'abcdefghijkmnpqrstuvwxyz'
    $upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $digits  = '23456789'
    $symbols = '!@#$%^&*-_=+'
    $all     = $lower + $upper + $digits + $symbols
    $bytes   = [System.Security.Cryptography.RandomNumberGenerator]::GetBytes($Length)

    $chars = [System.Collections.Generic.List[char]]::new()
    $chars.Add($lower[$bytes[0] % $lower.Length])
    $chars.Add($upper[$bytes[1] % $upper.Length])
    $chars.Add($digits[$bytes[2] % $digits.Length])
    $chars.Add($symbols[$bytes[3] % $symbols.Length])
    for ($i = 4; $i -lt $Length; $i++) { $chars.Add($all[$bytes[$i] % $all.Length]) }

    # Shuffle
    $shuffled = $chars | Sort-Object { Get-Random }
    -join $shuffled
}

# ── Tenant resolution (GDAP-aware) ──────────────────────────────────────────────
$effectiveTenantId = $TenantId
if (-not $effectiveTenantId) {
    try {
        if ($global:authMode -eq 'GDAP' -and $global:cid) { $effectiveTenantId = [string]$global:cid }
        elseif ($env:M365_CUSTOMER_TENANTID) { $effectiveTenantId = [string]$env:M365_CUSTOMER_TENANTID }
    } catch {}
}

# ── Connection (reuse existing session if present) ──────────────────────────────
$script:ConnectedHere = $false
try {
    if (-not (Get-MgContext)) {
        $connectParams = @{ Scopes = @('User.ReadWrite.All', 'RoleManagement.ReadWrite.Directory', 'GroupMember.ReadWrite.All') }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams -NoWelcome -ErrorAction Stop
        $script:ConnectedHere = $true
    }
} catch {
    Write-Host "  [ERROR] Could not connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New Break-Glass Admin Account" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  UPN         : $UserPrincipalName"
Write-Host "  DisplayName : $DisplayName"
Write-Host "  Mode        : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$existing = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "  [WARN] A user with UPN '$UserPrincipalName' already exists (Id: $($existing.Id)). Nothing to create." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 0
}

$password = New-RandomPassword -Length $PasswordLength

if (-not $Apply) {
    Write-Host "  Would create cloud-only user '$UserPrincipalName' with a random $PasswordLength-character password." -ForegroundColor Yellow
    if ($AssignGlobalAdmin) { Write-Host "  Would assign directory role: Global Administrator." -ForegroundColor Yellow }
    if ($ExcludeFromGroupId) { Write-Host "  Would add the account to group $ExcludeFromGroupId (e.g. a CA-exclusion group)." -ForegroundColor Yellow }
    Write-Host "  Re-run with -Apply to perform these actions." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 0
}

if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create break-glass admin account")) {
    $passwordProfile = @{
        Password                             = $password
        ForceChangePasswordNextSignIn        = $false
        ForceChangePasswordNextSignInWithMfa = $false
    }

    $newUser = New-MgUser -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName `
        -MailNickname ($UserPrincipalName.Split('@')[0]) -AccountEnabled `
        -PasswordProfile $passwordProfile -ErrorAction Stop

    Write-Host "  [OK]   Account created (Id: $($newUser.Id))." -ForegroundColor Green

    if ($AssignGlobalAdmin) {
        $role = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'" -ErrorAction SilentlyContinue
        if (-not $role) {
            $template = Get-MgDirectoryRoleTemplate -ErrorAction Stop | Where-Object { $_.DisplayName -eq 'Global Administrator' }
            $role = New-MgDirectoryRoleTemplate -RoleTemplateId $template.Id -ErrorAction Stop
        }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)"
        } -ErrorAction Stop
        Write-Host "  [OK]   Global Administrator role assigned." -ForegroundColor Green
    }

    if ($ExcludeFromGroupId) {
        New-MgGroupMemberByRef -GroupId $ExcludeFromGroupId -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)"
        } -ErrorAction Stop
        Write-Host "  [OK]   Added to exclusion group $ExcludeFromGroupId." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   PASSWORD (shown once — store it securely now)" -ForegroundColor Yellow
    Write-Host "   $password" -ForegroundColor Yellow
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Next steps: store the password in your password manager / sealed" -ForegroundColor DarkGray
    Write-Host "  envelope process, disable MFA on this account, and monitor sign-ins." -ForegroundColor DarkGray
}

if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
