#Requires -Version 7.0
<#
.SYNOPSIS
    Add or remove a user from a named Entra ID group, to gate access to an add-on feature.

.DESCRIPTION
    Modern Microsoft Graph replacement for tagging a user by setting a mailbox custom
    attribute (the legacy approach, which only worked for reporting and did not
    actually drive Conditional Access/Intune scoping). Adds (or removes) a user as a
    direct member of an Entra ID security group — the supported way to gate a
    Conditional Access policy, Intune app assignment, or license group ("enable this
    add-on for these users") to a subset of the tenant.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated by default (you
    sign in as an admin; device code when $global:useDeviceCodeAuth is set, the GDAP
    customer tenant from $global:cid), app-only with -ClientId and
    -CertificateThumbprint, or -AppOnly. An existing Graph session is reused only when
    it is for the right tenant and already holds the scopes below.

.PARAMETER UserId
    UPN or object ID of the user.

.PARAMETER GroupId
    Object ID of the target group. Use -GroupName instead to resolve by display name.

.PARAMETER GroupName
    Display name of the target group (resolved to an ID automatically; errors if
    ambiguous). Use instead of -GroupId.

.PARAMETER Remove
    Remove the user from the group instead of adding them.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer tenant ($global:cid /
    $env:M365_CUSTOMER_TENANTID), else the tenant you sign in to.

.PARAMETER ClientId
    App registration for app-only sign-in, with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER Apply
    Actually change group membership. Without this switch, the script only reports
    what it would do.

.EXAMPLE
    .\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365"

.EXAMPLE
    .\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Apply

.EXAMPLE
    .\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Remove -Apply

.NOTES
    Required scopes : GroupMember.ReadWrite.All, User.Read.All, Group.Read.All
    Required module : Microsoft.Graph.Users, Microsoft.Graph.Groups
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $UserId,

    [string] $GroupId,
    [string] $GroupName,
    [switch] $Remove,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [switch] $Apply
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $GroupId -and -not $GroupName) { throw "Provide either -GroupId or -GroupName." }

try {
    $graph = Connect-M365Graph -Scopes 'GroupMember.ReadWrite.All', 'User.Read.All', 'Group.Read.All' -TenantId $TenantId `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
} catch {
    Write-Host "  [ERROR] Could not connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

if ($GroupName) {
    $groups = @(Get-MgGroup -Filter "displayName eq '$($GroupName -replace "'", "''")'" -ErrorAction Stop)
    if ($groups.Count -eq 0) { Write-Host "  [ERROR] No group found named '$GroupName'." -ForegroundColor Red; Disconnect-M365Graph $graph; exit 1 }
    if ($groups.Count -gt 1) { Write-Host "  [ERROR] Multiple groups named '$GroupName' — use -GroupId instead." -ForegroundColor Red; Disconnect-M365Graph $graph; exit 1 }
    $GroupId = $groups[0].Id
}

$user = Get-MgUser -UserId $UserId -ErrorAction Stop

Write-Host ""
Write-Host "  Add-UserToFeatureGroup : $($user.UserPrincipalName)" -ForegroundColor Cyan
Write-Host "  Group  : $GroupId"
Write-Host "  Action : $(if ($Remove) { 'Remove' } else { 'Add' })"
Write-Host "  Mode   : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would $(if ($Remove) { 'remove' } else { 'add' }) '$($user.UserPrincipalName)' $(if ($Remove) { 'from' } else { 'to' }) group $GroupId." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform this change." -ForegroundColor Yellow
    Disconnect-M365Graph $graph
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, "$(if ($Remove) { 'Remove from' } else { 'Add to' }) group $GroupId")) {
    Disconnect-M365Graph $graph
    exit 0
}

try {
    if ($Remove) {
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $user.Id -ErrorAction Stop
        Write-Host "  [OK]   Removed." -ForegroundColor Green
    } else {
        New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" } -ErrorAction Stop
        Write-Host "  [OK]   Added." -ForegroundColor Green
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Disconnect-M365Graph $graph
