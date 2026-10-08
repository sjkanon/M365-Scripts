#Requires -Version 7.0
<#
.SYNOPSIS
    One way to sign in to Microsoft 365 for every script in this repo: Graph first,
    delegated by default, app-only on request.

.DESCRIPTION
    Dot-source this file at the top of a script:

        . (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')     # adjust the depth

    and connect with one call:

        $graph = Connect-M365Graph -Scopes 'User.Read.All' -TenantId $TenantId
        ...
        Disconnect-M365Graph $graph       # only disconnects what this call connected

    Every Connect-M365* function follows the same rules:

    Delegated (the default)
        You sign in as yourself. Device code when $global:useDeviceCodeAuth is set in
        load.config.ps1 (or -DeviceCode is passed), otherwise an interactive browser.
        Under GDAP ($global:authMode = 'GDAP') the customer tenant comes from
        $global:cid / $global:connectmsoldomain (set by Connect-Tenant) or from
        $env:M365_CUSTOMER_TENANTID, unless -TenantId names one.

    App-only (on request)
        -ClientId with -CertificateThumbprint (or -ClientSecret for Graph), or -AppOnly
        to take ClientId and CertificateThumbprint for the tenant from graph.appid.json
        in the repo root (gitignored; shape: { "<tenant>": { ClientId, CertificateThumbprint } }).
        The app must be consented in that tenant - GDAP does not grant app-only access.

    An existing session is reused when it is the right kind (delegated/app-only), for
    the right tenant, and - delegated - already holds every requested scope. Otherwise
    the functions connect, and say so in the object they return, so the script only
    disconnects what it opened itself and never the caller's session.

    Graph is the standard. Connect-M365Exchange, Connect-M365Teams and Connect-M365PnP
    exist for the work Graph has no API for (mailbox and SendAs permissions, message
    trace, DKIM, EOP policies, Teams Cs* policies, SharePoint role assignments, ...).

.NOTES
    Functions:
        Resolve-M365TenantId     The tenant to use: -TenantId, else the GDAP customer, else none.
        Connect-M365Graph        Microsoft Graph (Microsoft.Graph.Authentication).
        Disconnect-M365Graph     Disconnects only when Connect-M365Graph connected.
        Connect-M365Exchange     Exchange Online (ExchangeOnlineManagement 3.x).
        Disconnect-M365Exchange  Disconnects only when Connect-M365Exchange connected.
        Connect-M365Teams        Microsoft Teams (MicrosoftTeams 5.x).
        Connect-M365PnP          PnP.PowerShell; returns the connection object.
#>

#region Tenant and settings

function Test-M365Gdap {
    $mode = if ($global:authMode) { [string]$global:authMode } elseif ($env:M365_AUTH_MODE) { [string]$env:M365_AUTH_MODE } else { '' }
    return $mode.ToUpperInvariant() -eq 'GDAP'
}

function Test-M365DeviceCode {
    param([switch] $DeviceCode, [switch] $Interactive)
    if ($Interactive) { return $false }
    if ($DeviceCode)  { return $true }
    return [bool]$global:useDeviceCodeAuth
}

function Resolve-M365TenantId {
    <#
    .SYNOPSIS
        The tenant to connect to: -TenantId when given, else the GDAP customer, else $null
        (sign in to your own tenant).
    #>
    param([string] $TenantId)
    if ($TenantId) { return $TenantId }
    if (Test-M365Gdap) {
        if ($global:cid)               { return [string]$global:cid }
        if ($global:connectmsoldomain) { return [string]$global:connectmsoldomain }
    }
    if ($env:M365_CUSTOMER_TENANTID) { return [string]$env:M365_CUSTOMER_TENANTID }
    return $null
}

function Resolve-M365CustomerDomain {
    # Exchange's -DelegatedOrganization and app-only -Organization want a domain; a GUID
    # works for -DelegatedOrganization too, so fall back to it.
    param([string] $TenantId)
    if ($TenantId -and $TenantId -notmatch '^[0-9a-fA-F-]{36}$') { return $TenantId }
    if ((Test-M365Gdap) -and $global:connectmsoldomain -and (-not $TenantId -or $TenantId -eq [string]$global:cid)) {
        return [string]$global:connectmsoldomain
    }
    return $TenantId
}

function Get-M365AppRegistration {
    <#
    .SYNOPSIS
        ClientId and CertificateThumbprint for a tenant from graph.appid.json (repo root).
    #>
    param([string] $TenantId)
    $file = Join-Path $PSScriptRoot '..\..\graph.appid.json'
    if (-not (Test-Path $file)) { throw "-AppOnly needs graph.appid.json in the repo root, or pass -ClientId and -CertificateThumbprint." }
    $map = Get-Content $file -Raw | ConvertFrom-Json -AsHashtable
    $keys = @($map.Keys)
    $key = if ($TenantId) { $keys | Where-Object { $_ -eq $TenantId } | Select-Object -First 1 }
           elseif ($keys.Count -eq 1) { $keys[0] }
    if (-not $key) {
        throw "graph.appid.json has no entry for tenant '$TenantId' (it has: $($keys -join ', ')). Pass -TenantId with one of those, or -ClientId and -CertificateThumbprint."
    }
    [pscustomobject]@{ Tenant = $key; ClientId = $map[$key].ClientId; CertificateThumbprint = $map[$key].CertificateThumbprint }
}

#endregion

#region Microsoft Graph

function Connect-M365Graph {
    <#
    .SYNOPSIS
        Connect to Microsoft Graph: delegated by default, app-only with -ClientId or -AppOnly.

    .PARAMETER Scopes
        Delegated permissions the script needs. Ignored for app-only (the app's consented
        application permissions apply).

    .PARAMETER TenantId
        Tenant id or domain. Defaults to the GDAP customer when authMode is GDAP.

    .PARAMETER ClientId
        App registration for app-only sign-in, with -CertificateThumbprint or -ClientSecret.

    .PARAMETER AppOnly
        App-only with ClientId and CertificateThumbprint from graph.appid.json.

    .PARAMETER DeviceCode
        Delegated sign-in with a device code, whatever load.config.ps1 says.

    .PARAMETER Interactive
        Delegated sign-in in the browser, whatever load.config.ps1 says.

    .OUTPUTS
        [pscustomobject] ConnectedHere, AuthType (Delegated/AppOnly), TenantId, Account.
    #>
    [CmdletBinding()]
    param(
        [string[]]     $Scopes = @(),
        [string]       $TenantId,
        [string]       $ClientId,
        [string]       $CertificateThumbprint,
        [securestring] $ClientSecret,
        [switch]       $AppOnly,
        [switch]       $DeviceCode,
        [switch]       $Interactive
    )

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw 'Microsoft.Graph.Authentication is not installed. Run scripts\Startup\Install-Modules.ps1.'
    }

    $tenant = Resolve-M365TenantId -TenantId $TenantId

    if ($AppOnly -and -not $ClientId) {
        $reg = Get-M365AppRegistration -TenantId $tenant
        $ClientId = $reg.ClientId
        $CertificateThumbprint = $reg.CertificateThumbprint
        if (-not $tenant) { $tenant = $reg.Tenant }
    }
    $wantAppOnly = [bool]$ClientId
    if ($wantAppOnly -and -not $tenant) { throw 'App-only sign-in needs -TenantId.' }
    if ($wantAppOnly -and -not $CertificateThumbprint -and -not $ClientSecret) {
        throw '-ClientId needs -CertificateThumbprint or -ClientSecret.'
    }

    # Reuse what is there when it fits.
    $ctx = Get-MgContext
    if ($ctx) {
        $tenantOk = -not $tenant -or $ctx.TenantId -eq $tenant -or
                    ($tenant -notmatch '^[0-9a-fA-F-]{36}$' -and $ctx.Account -and $ctx.Account -like "*@$tenant")
        if ($wantAppOnly) {
            $fits = $ctx.AuthType -eq 'AppOnly' -and $ctx.ClientId -eq $ClientId -and $tenantOk
        } else {
            $missing = @($Scopes | Where-Object { $_ -notin $ctx.Scopes })
            $fits = $ctx.AuthType -eq 'Delegated' -and $tenantOk -and $missing.Count -eq 0
        }
        if ($fits) {
            return [pscustomobject]@{ ConnectedHere = $false; AuthType = [string]$ctx.AuthType; TenantId = $ctx.TenantId; Account = $ctx.Account }
        }
    }

    $p = @{ NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
    if ($tenant) { $p['TenantId'] = $tenant }

    if ($wantAppOnly) {
        $p['ClientId'] = $ClientId
        if ($CertificateThumbprint) {
            $p['CertificateThumbprint'] = $CertificateThumbprint
        } else {
            $p.Remove('ClientId')
            $p['ClientSecretCredential'] = [pscredential]::new($ClientId, $ClientSecret)
        }
        Write-Host "  Connecting to Microsoft Graph (app-only, $ClientId)..." -ForegroundColor DarkGray
    } else {
        # Keep the scopes an existing delegated session already had, so a second script
        # in the same process does not lose what the first one asked for.
        $all = @($Scopes)
        if ($ctx -and $ctx.AuthType -eq 'Delegated') { $all = @($ctx.Scopes) + $all }
        $all = @($all | Where-Object { $_ } | Select-Object -Unique)
        if ($all.Count) { $p['Scopes'] = $all }
        $device = Test-M365DeviceCode -DeviceCode:$DeviceCode -Interactive:$Interactive
        if ($device) { $p['UseDeviceCode'] = $true }
        elseif ($global:upn) { $p['LoginHint'] = [string]$global:upn }
        Write-Host "  Connecting to Microsoft Graph (delegated$(if ($device) { ', device code' })$(if ($tenant) { ", tenant $tenant" }))..." -ForegroundColor DarkGray
    }

    Connect-MgGraph @p | Out-Null
    $ctx = Get-MgContext
    [pscustomobject]@{ ConnectedHere = $true; AuthType = [string]$ctx.AuthType; TenantId = $ctx.TenantId; Account = $ctx.Account }
}

function Disconnect-M365Graph {
    param([Parameter(Position = 0)] $Connection)
    if ($Connection -and $Connection.ConnectedHere) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

#endregion

#region Exchange Online

function Connect-M365Exchange {
    <#
    .SYNOPSIS
        Connect to Exchange Online, for the work Graph has no API for.
        Delegated by default (GDAP via -DelegatedOrganization), app-only with -ClientId
        and -CertificateThumbprint, or -AppOnly from graph.appid.json.

    .PARAMETER IncludeCompliance
        Also connect to Security & Compliance PowerShell (Connect-IPPSSession).

    .OUTPUTS
        [pscustomobject] ConnectedHere, AuthType, Organization.
    #>
    [CmdletBinding()]
    param(
        [string] $TenantId,
        [string] $ClientId,
        [string] $CertificateThumbprint,
        [switch] $AppOnly,
        [switch] $DeviceCode,
        [switch] $Interactive,
        [switch] $IncludeCompliance
    )

    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        throw 'ExchangeOnlineManagement is not installed. Run scripts\Startup\Install-Modules.ps1.'
    }

    $tenant = Resolve-M365TenantId -TenantId $TenantId
    if ($AppOnly -and -not $ClientId) {
        $reg = Get-M365AppRegistration -TenantId $tenant
        $ClientId = $reg.ClientId; $CertificateThumbprint = $reg.CertificateThumbprint
        if (-not $tenant) { $tenant = $reg.Tenant }
    }
    $wantAppOnly = [bool]$ClientId
    $org = Resolve-M365CustomerDomain -TenantId $tenant

    $existing = @(Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and -not $_.IsEopSession })
    $fits = $existing | Where-Object {
        (-not $org -or $_.TenantID -eq $tenant -or $_.Organization -eq $org -or $_.DelegatedOrganization -eq $org) -and
        ((-not $wantAppOnly) -or $_.AppId -eq $ClientId)
    } | Select-Object -First 1

    $connected = $false
    if (-not $fits) {
        $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($wantAppOnly) {
            if (-not $org -or $org -match '^[0-9a-fA-F-]{36}$') { throw 'App-only Exchange sign-in needs -TenantId as a domain (contoso.onmicrosoft.com).' }
            if (-not $CertificateThumbprint) { throw '-ClientId needs -CertificateThumbprint for Exchange Online.' }
            $p['AppId'] = $ClientId; $p['CertificateThumbprint'] = $CertificateThumbprint; $p['Organization'] = $org
            Write-Host "  Connecting to Exchange Online (app-only, $org)..." -ForegroundColor DarkGray
        } else {
            # -Organization only applies to app-only sign-in; a partner reaches a customer
            # with -DelegatedOrganization. Without GDAP (or -TenantId) you land in your own tenant.
            if ($org -and ($TenantId -or (Test-M365Gdap))) { $p['DelegatedOrganization'] = $org }
            if ($global:upn) { $p['UserPrincipalName'] = [string]$global:upn }
            $device = Test-M365DeviceCode -DeviceCode:$DeviceCode -Interactive:$Interactive
            if ($device) { $p['Device'] = $true; $p.Remove('UserPrincipalName') }
            Write-Host "  Connecting to Exchange Online (delegated$(if ($device) { ', device code' })$(if ($p['DelegatedOrganization']) { ", $org" }))..." -ForegroundColor DarkGray
        }
        Connect-ExchangeOnline @p
        $connected = $true
    }

    if ($IncludeCompliance) {
        $ipps = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' -and $_.IsEopSession })
        if (-not $ipps) {
            $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
            if ($wantAppOnly) {
                $p['AppId'] = $ClientId; $p['CertificateThumbprint'] = $CertificateThumbprint; $p['Organization'] = $org
            } else {
                if ($org -and ($TenantId -or (Test-M365Gdap))) { $p['DelegatedOrganization'] = $org }
                if ($global:upn) { $p['UserPrincipalName'] = [string]$global:upn }
            }
            Write-Host '  Connecting to Security & Compliance PowerShell...' -ForegroundColor DarkGray
            Connect-IPPSSession @p
            $connected = $true
        }
    }

    [pscustomobject]@{ ConnectedHere = $connected; AuthType = $(if ($wantAppOnly) { 'AppOnly' } else { 'Delegated' }); Organization = $org }
}

function Disconnect-M365Exchange {
    param([Parameter(Position = 0)] $Connection)
    if ($Connection -and $Connection.ConnectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
}

#endregion

#region Microsoft Teams

function Connect-M365Teams {
    <#
    .SYNOPSIS
        Connect to Microsoft Teams PowerShell, for the Cs* policies Graph has no API for.
        Delegated by default, app-only with -ClientId and -CertificateThumbprint.
    #>
    [CmdletBinding()]
    param(
        [string] $TenantId,
        [string] $ClientId,
        [string] $CertificateThumbprint,
        [switch] $AppOnly,
        [switch] $DeviceCode,
        [switch] $Interactive
    )

    if (-not (Get-Command Connect-MicrosoftTeams -ErrorAction SilentlyContinue)) {
        throw 'MicrosoftTeams is not installed. Run scripts\Startup\Install-Modules.ps1.'
    }
    $tenant = Resolve-M365TenantId -TenantId $TenantId
    if ($AppOnly -and -not $ClientId) {
        $reg = Get-M365AppRegistration -TenantId $tenant
        $ClientId = $reg.ClientId; $CertificateThumbprint = $reg.CertificateThumbprint
        if (-not $tenant) { $tenant = $reg.Tenant }
    }

    try {
        $current = Get-CsTenant -ErrorAction Stop
        if (-not $tenant -or $current.TenantId -eq $tenant) {
            return [pscustomobject]@{ ConnectedHere = $false; TenantId = [string]$current.TenantId }
        }
    } catch { }

    $p = @{ ErrorAction = 'Stop' }
    if ($tenant) { $p['TenantId'] = $tenant }
    if ($ClientId) {
        if (-not $tenant -or -not $CertificateThumbprint) { throw 'App-only Teams sign-in needs -TenantId and -CertificateThumbprint.' }
        $p['ApplicationId'] = $ClientId; $p['CertificateThumbprint'] = $CertificateThumbprint
    } elseif (Test-M365DeviceCode -DeviceCode:$DeviceCode -Interactive:$Interactive) {
        $p['UseDeviceAuthentication'] = $true
    }
    Write-Host '  Connecting to Microsoft Teams...' -ForegroundColor DarkGray
    $r = Connect-MicrosoftTeams @p
    [pscustomobject]@{ ConnectedHere = $true; TenantId = [string]$r.TenantId }
}

#endregion

#region PnP.PowerShell

function Connect-M365PnP {
    <#
    .SYNOPSIS
        Connect PnP.PowerShell to a SharePoint URL, for SharePoint work Graph has no API
        for (role assignments, SharePoint groups, recycle bin, provisioning).
        PnP needs an app registration of your own since September 2024: -ClientId, else
        the tenant's entry in pnp.appid.json (repo root). Delegated by default
        (device code per load.config.ps1), app-only with -CertificateThumbprint.

    .OUTPUTS
        The PnP connection (pass it on with -Connection).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Url,
        [string] $TenantId,
        [string] $ClientId,
        [string] $CertificateThumbprint,
        [switch] $DeviceCode,
        [switch] $Interactive
    )

    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        throw 'PnP.PowerShell is not installed (needs PowerShell 7.4). Run scripts\Startup\Install-Modules.ps1.'
    }
    $tenant = Resolve-M365TenantId -TenantId $TenantId
    if (-not $tenant -and $Url -match '^https://([^./]+?)(-admin)?\.sharepoint\.com') { $tenant = "$($Matches[1]).onmicrosoft.com" }

    if (-not $ClientId) {
        $file = Join-Path $PSScriptRoot '..\..\pnp.appid.json'
        if (Test-Path $file) {
            $map = Get-Content $file -Raw | ConvertFrom-Json -AsHashtable
            if ($tenant -and $map.ContainsKey($tenant)) { $ClientId = [string]$map[$tenant] }
        }
    }
    if (-not $ClientId) {
        throw "PnP needs -ClientId (an app registration in tenant '$tenant'), or an entry for that tenant in pnp.appid.json."
    }

    $p = @{ Url = $Url; ClientId = $ClientId; ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($CertificateThumbprint) {
        if (-not $tenant) { throw 'App-only PnP sign-in needs -TenantId.' }
        $p['Tenant'] = $tenant; $p['Thumbprint'] = $CertificateThumbprint
    } elseif (Test-M365DeviceCode -DeviceCode:$DeviceCode -Interactive:$Interactive) {
        $p['DeviceLogin'] = $true
        if ($tenant) { $p['Tenant'] = $tenant }
    } else {
        $p['Interactive'] = $true
    }
    Connect-PnPOnline @p
}

#endregion
