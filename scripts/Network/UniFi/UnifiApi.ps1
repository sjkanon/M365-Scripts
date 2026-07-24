#Requires -Version 5.1
<#
.SYNOPSIS
    Shared helper functions for talking to a UniFi Network Controller / UniFi OS console.

.DESCRIPTION
    Dot-sourced by Get-UnifiNetworkReport.ps1 and Update-UnifiFirmware.ps1 — not meant to be
    run directly. Handles login (classic self-hosted controller and UniFi OS consoles such as
    UDM/UDM-Pro/UDR, which proxy the network application under a different path and use a
    different auth endpoint + CSRF header), session cookies, and self-signed certificate
    handling for both Windows PowerShell 5.1 and PowerShell 7+.

    Credentials are always supplied via Get-Credential (interactively or by the caller) —
    never hardcoded in these scripts.
#>

function Invoke-UnifiRestMethod {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $Uri,
        [string] $Method = 'GET',
        [object] $Body,
        [Microsoft.PowerShell.Commands.WebRequestSession] $WebSession,
        [hashtable] $Headers,
        [switch] $SkipCertificateCheck
    )

    $params = @{
        Uri             = $Uri
        Method          = $Method
        WebSession      = $WebSession
        ContentType     = 'application/json'
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($Body)    { $params['Body'] = ($Body | ConvertTo-Json -Depth 10) }
    if ($Headers) { $params['Headers'] = $Headers }

    if ($PSVersionTable.PSVersion.Major -ge 6) {
        if ($SkipCertificateCheck) { $params['SkipCertificateCheck'] = $true }
        return Invoke-RestMethod @params
    }

    # Windows PowerShell 5.1 has no -SkipCertificateCheck; fall back to a callback.
    if ($SkipCertificateCheck) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
    try {
        return Invoke-RestMethod @params
    } finally {
        if ($SkipCertificateCheck) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        }
    }
}

function Connect-UnifiController {
    <#
    .SYNOPSIS
        Log in to a UniFi Network Controller or UniFi OS console and return a session object.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $Controller,
        [Parameter(Mandatory)] [pscredential] $Credential,
        [switch] $SkipCertificateCheck
    )

    $base = $Controller.TrimEnd('/')
    $body = @{
        username = $Credential.UserName
        password = $Credential.GetNetworkCredential().Password
    }

    $isModernPS = $PSVersionTable.PSVersion.Major -ge 6
    $legacyCertBypass = $SkipCertificateCheck -and -not $isModernPS
    if ($legacyCertBypass) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }

    try {
        # Try UniFi OS console login first (UDM/UDM-Pro/UDR/Cloud Gateway), fall back to classic controller.
        try {
            $resp = Invoke-WebRequest -Uri "$base/api/auth/login" -Method Post -Body ($body | ConvertTo-Json) `
                -ContentType 'application/json' -SessionVariable webSessionVar -UseBasicParsing `
                -SkipCertificateCheck:($isModernPS -and $SkipCertificateCheck) -ErrorAction Stop
            $csrf = $resp.Headers['X-CSRF-Token']
            return [PSCustomObject]@{
                BaseUri       = $base
                ApiPrefix     = '/proxy/network'
                WebSession    = $webSessionVar
                CsrfToken     = $csrf
                IsUnifiOs     = $true
                SkipCertCheck = [bool]$SkipCertificateCheck
            }
        } catch {
            Write-Verbose "UniFi OS login failed ($($_.Exception.Message)), falling back to classic controller login."
        }

        # Classic self-hosted controller (port 8443 typically).
        $null = Invoke-WebRequest -Uri "$base/api/login" -Method Post -Body ($body | ConvertTo-Json) `
            -ContentType 'application/json' -SessionVariable webSessionVar -UseBasicParsing `
            -SkipCertificateCheck:($isModernPS -and $SkipCertificateCheck) -ErrorAction Stop

        return [PSCustomObject]@{
            BaseUri       = $base
            ApiPrefix     = ''
            WebSession    = $webSessionVar
            CsrfToken     = $null
            IsUnifiOs     = $false
            SkipCertCheck = [bool]$SkipCertificateCheck
        }
    } finally {
        if ($legacyCertBypass) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        }
    }
}

function Disconnect-UnifiController {
    [CmdletBinding()]
    param ([Parameter(Mandatory)] [PSCustomObject] $Session)

    $loginPath = if ($Session.IsUnifiOs) { '/api/auth/logout' } else { '/api/logout' }
    try {
        Invoke-UnifiRestMethod -Uri "$($Session.BaseUri)$loginPath" -Method Post `
            -WebSession $Session.WebSession -SkipCertificateCheck:$Session.SkipCertCheck | Out-Null
    } catch {
        Write-Verbose "Logout request failed (session likely already expired): $($_.Exception.Message)"
    }
}

function Invoke-UnifiApi {
    <#
    .SYNOPSIS
        Call a UniFi API endpoint relative to a site, e.g. Invoke-UnifiApi -Session $s -Site default -Path stat/device
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Session,
        [Parameter(Mandatory)] [string] $Path,
        [string] $Site = 'default',
        [string] $Method = 'GET',
        [object] $Body
    )

    $headers = @{}
    if ($Session.CsrfToken) { $headers['X-CSRF-Token'] = $Session.CsrfToken }

    $uri = "$($Session.BaseUri)$($Session.ApiPrefix)/api/s/$Site/$Path"
    $result = Invoke-UnifiRestMethod -Uri $uri -Method $Method -Body $Body -Headers $headers `
        -WebSession $Session.WebSession -SkipCertificateCheck:$Session.SkipCertCheck
    return $result.data
}

function Get-UnifiSite {
    [CmdletBinding()]
    param ([Parameter(Mandatory)] [PSCustomObject] $Session)

    $headers = @{}
    if ($Session.CsrfToken) { $headers['X-CSRF-Token'] = $Session.CsrfToken }
    $uri = "$($Session.BaseUri)$($Session.ApiPrefix)/api/self/sites"
    $result = Invoke-UnifiRestMethod -Uri $uri -Headers $headers -WebSession $Session.WebSession `
        -SkipCertificateCheck:$Session.SkipCertCheck
    return $result.data
}

function Get-UnifiDevice {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Session,
        [Parameter(Mandatory)] [string] $Site
    )
    return Invoke-UnifiApi -Session $Session -Site $Site -Path 'stat/device'
}
