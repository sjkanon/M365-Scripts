#Requires -Version 7.0
<#
.SYNOPSIS
    Put a shortcut to one SharePoint folder in the OneDrive of every listed user
    ("Add shortcut to My files"). Reports by default, creates with -Apply.

.DESCRIPTION
    1. Resolves the target folder from its SharePoint URL (or -DriveId/-ItemId).
    2. Collects users from -UserPrincipalName, -CsvPath and/or -GroupId (transitive members,
       enabled accounts only).
    3. Per user: looks in the root of their OneDrive for a shortcut that already points at the
       folder. Present -> AlreadyThere. Missing -> WouldCreate, or Created with -Apply.

    Sign-in works like the other SharePoint scripts in this repo: you sign in delegated as an
    admin (Global Administrator or Application Administrator; GDAP customer per
    load.config.ps1), the script creates a temporary App Registration with exactly the
    application permissions it needs, works app-only with that app, and deletes it at the end
    (also on an error or Ctrl+C). A delegated admin cannot write into someone else's OneDrive
    itself, which is why the app is needed.

        Report     Files.Read.All, Sites.Read.All, User.Read.All
        -Apply     Files.ReadWrite.All instead of Files.Read.All
        -GroupId   + GroupMember.Read.All

    -AppOnly uses the certificate app for the tenant from graph.appid.json instead, and
    -ClientId/-CertificateThumbprint an app of your own; both skip the temporary app. That app
    needs the permissions above.

    Mind:
      - The user needs access to the folder themselves; otherwise the shortcut is created but
        leads nowhere. The script does not grant access.
      - The OneDrive must exist (the user has signed in at least once). If not: NoOneDrive.

.PARAMETER FolderUrl
    URL of the folder in SharePoint, e.g.
    https://contoso.sharepoint.com/sites/Team/Shared Documents/Projects
    A library view URL from the address bar (.../Forms/AllItems.aspx?id=/sites/...) and a
    "Copy link" sharing link work as well.

.PARAMETER DriveId
    Instead of -FolderUrl: drive id of the document library.

.PARAMETER ItemId
    Instead of -FolderUrl: item id of the folder.

.PARAMETER ShortcutName
    Name of the shortcut in OneDrive. Defaults to the folder name.

.PARAMETER UserPrincipalName
    One or more UPNs.

.PARAMETER CsvPath
    CSV with a column UserPrincipalName.

.PARAMETER GroupId
    One or more Entra groups, by object id or mail address; every enabled user in them (nested
    included) gets the shortcut.

.PARAMETER TenantId
    Tenant id or domain. Defaults to the GDAP customer (load.config.ps1), else the tenant you
    sign in to.

.PARAMETER AppOnly
    Use the certificate app for the tenant from graph.appid.json instead of a temporary app.

.PARAMETER ClientId
    Use this app registration instead of a temporary app, with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Certificate for -ClientId.

.PARAMETER Apply
    Create the shortcuts. Without it nothing changes in any OneDrive.

.PARAMETER OutputPath
    CSV with one row per user. Defaults to C:\Temp\OneDriveShortcut_<timestamp>.csv.

.EXAMPLE
    .\Add-OneDriveShortcut.ps1 -TenantId contoso.onmicrosoft.com `
        -FolderUrl 'https://contoso.sharepoint.com/sites/Team/Shared Documents/Projects' `
        -GroupId 00000000-0000-0000-0000-000000000000

    Report: who already has the shortcut and who would get it.

.EXAMPLE
    .\Add-OneDriveShortcut.ps1 -TenantId contoso.onmicrosoft.com -FolderUrl $url `
        -UserPrincipalName jan@contoso.com -Apply

    Create it for one user first, and check it in their OneDrive before doing the rest.
#>
[CmdletBinding(DefaultParameterSetName = 'Url')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Url')]
    [string] $FolderUrl,

    [Parameter(Mandatory, ParameterSetName = 'Ids')]
    [string] $DriveId,

    [Parameter(Mandatory, ParameterSetName = 'Ids')]
    [string] $ItemId,

    [string]   $ShortcutName,
    [string[]] $UserPrincipalName,
    [string]   $CsvPath,
    [string[]] $GroupId,

    [string] $TenantId,
    [switch] $AppOnly,
    [string] $ClientId,
    [string] $CertificateThumbprint,

    [switch] $Apply,
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

if (-not $UserPrincipalName -and -not $CsvPath -and -not $GroupId) {
    throw 'No users given. Use -UserPrincipalName, -CsvPath and/or -GroupId.'
}

$script:TempApp     = $null   # [pscustomobject] ObjectId, AppId, TokenBody, TenantId
$script:Headers     = $null   # app-only bearer header when the temporary app is used
$script:TokenExpiry = [datetime]::MinValue
$script:Warmup      = $false  # true until the first Graph call with a fresh temp app succeeds

# ── Temporary app ────────────────────────────────────────────────────────────

function New-TempApp {
    # One App Registration + service principal with the given Graph application roles and a
    # one-day secret. Deleted again by Remove-TempApp.
    param([Parameter(Mandatory)][string[]] $Roles, [Parameter(Mandatory)][string] $Tenant)

    $name = "OneDrive-Shortcut-Temp-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Write-Host "  Creating temporary App Registration '$name'..." -ForegroundColor Cyan
    $app = New-MgApplication -DisplayName $name -SignInAudience 'AzureADMyOrg'
    $script:TempApp = [pscustomobject]@{ ObjectId = $app.Id; AppId = $app.AppId; TokenBody = $null; TenantId = $Tenant }

    $sp = New-MgServicePrincipal -AppId $app.AppId
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
    foreach ($value in $Roles) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $value -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $role) { throw "Could not resolve Graph app role '$value'." }
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id `
            -ResourceId $graphSp.Id -AppRoleId $role.Id | Out-Null
        Write-Host "  [OK]   $value granted." -ForegroundColor DarkGray
    }

    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
        displayName = 'temp'
        endDateTime = (Get-Date).AddDays(1)
    }
    $script:TempApp.TokenBody = @{
        grant_type    = 'client_credentials'
        scope         = 'https://graph.microsoft.com/.default'
        client_id     = $app.AppId
        client_secret = $secret.SecretText
    }
    Update-AppToken
}

function Update-AppToken {
    # A new app needs a few seconds before Entra accepts its secret.
    if (-not $script:TempApp) { return }
    Write-Host '  Obtaining app-only token...' -ForegroundColor Cyan
    for ($i = 1; $i -le 8; $i++) {
        try {
            $resp = Invoke-RestMethod -Method POST -Body $script:TempApp.TokenBody `
                -Uri "https://login.microsoftonline.com/$($script:TempApp.TenantId)/oauth2/v2.0/token"
            $script:Headers = @{ Authorization = "Bearer $($resp.access_token)" }
            $script:TokenExpiry = (Get-Date).AddSeconds($resp.expires_in - 300)
            return
        } catch {
            if ($i -eq 8) { throw "Could not obtain an app-only token for the temporary app: $($_.Exception.Message)" }
            Write-Host "  [INFO] Waiting for app registration propagation (attempt $i/8)..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 5
        }
    }
}

function Remove-TempApp {
    # The delegated session is still open here, so the app can delete itself through it.
    if (-not $script:TempApp) { return }
    $id = $script:TempApp.ObjectId
    $script:TempApp = $null
    try {
        Remove-MgApplication -ApplicationId $id
        Write-Host '  [OK]   Temporary App Registration removed.' -ForegroundColor DarkGray
    } catch {
        Write-Host "  [WARN] Could not remove the temporary App Registration (object id $id)." -ForegroundColor Yellow
        Write-Host "  [WARN] Remove it by hand: Remove-MgApplication -ApplicationId $id" -ForegroundColor Yellow
    }
}

# ── Graph calls ──────────────────────────────────────────────────────────────

function Invoke-Graph {
    # Through the temporary app's token when there is one, else through the Graph session
    # (-AppOnly / -ClientId). Retries 401/403 a few times right after the app was created:
    # fresh role assignments take a moment before Graph honours them.
    param([Parameter(Mandatory)][string] $Uri, [string] $Method = 'GET', $Body)

    $full = if ($Uri -match '^https://') { $Uri } else { "https://graph.microsoft.com/v1.0/$($Uri.TrimStart('/'))" }
    if (-not $script:Headers) {
        $p = @{ Method = $Method; Uri = $full; OutputType = 'PSObject' }
        if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 10); $p.ContentType = 'application/json' }
        return Invoke-MgGraphRequest @p
    }

    for ($attempt = 1; ; $attempt++) {
        if ((Get-Date) -ge $script:TokenExpiry) { Update-AppToken }
        $p = @{ Method = $Method; Uri = $full; Headers = $script:Headers }
        if ($null -ne $Body) { $p.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10)); $p.ContentType = 'application/json; charset=utf-8' }
        try {
            return Invoke-RestMethod @p
        } catch {
            $status = [int]($_.Exception.Response.StatusCode)
            $warming   = ($status -in 401, 403) -and $script:Warmup -and ($attempt -le 6)
            $throttled = (($status -eq 429) -or ($status -ge 500)) -and ($attempt -le 4)
            $retry = $warming -or $throttled
            if (-not $retry) {
                $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
                throw "Graph $Method $Uri failed ($status): $detail"
            }
            Start-Sleep -Seconds ([Math]::Min(30, 5 * $attempt))
        }
    }
}

function Get-GraphAll {
    param([Parameter(Mandatory)][string] $Uri)
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph -Uri $next
        if ($page.PSObject.Properties['value']) { $page.value } else { $page }
        $next = if ($page.PSObject.Properties['@odata.nextLink']) { $page.'@odata.nextLink' }
    }
}

# ── Main ─────────────────────────────────────────────────────────────────────

$graph = $null
try {
    if ($AppOnly -or $ClientId) {
        $p = @{ TenantId = $TenantId }
        if ($ClientId) { $p.ClientId = $ClientId; $p.CertificateThumbprint = $CertificateThumbprint } else { $p.AppOnly = $true }
        $graph = Connect-M365Graph @p
    } else {
        Write-Host '  Connecting (delegated) to create the temporary app...' -ForegroundColor Cyan
        Write-Host '  Required role: Global Administrator or Application Administrator' -ForegroundColor DarkGray
        $graph = Connect-M365Graph -TenantId $TenantId -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
        $tenant = if (Resolve-M365TenantId -TenantId $TenantId) { Resolve-M365TenantId -TenantId $TenantId } else { (Get-MgContext).TenantId }

        $roles = @(
            $(if ($Apply) { 'Files.ReadWrite.All' } else { 'Files.Read.All' })
            'Sites.Read.All'
            'User.Read.All'
        )
        if ($GroupId) { $roles += 'GroupMember.Read.All' }
        New-TempApp -Roles $roles -Tenant $tenant
        $script:Warmup = $true
    }

    # ── Target folder ────────────────────────────────────────────────────────
    $select = '$select=id,name,folder,parentReference,webUrl'
    if ($PSCmdlet.ParameterSetName -eq 'Url') {
        $url = $FolderUrl.Trim()
        # A library view from the address bar (.../Forms/AllItems.aspx?id=/sites/x/Lib/Folder&viewid=...)
        # names the folder in its id parameter; turn it into the folder's own URL.
        $uri = [uri]$url
        if ($uri.Query -match '(?:^\?|&)id=([^&]+)') {
            $path = [uri]::UnescapeDataString($Matches[1])
            $url = "https://$($uri.Host)" + (($path.Split('/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
        }
        Write-Verbose "Folder URL: $url"
        # Sharing-URL encoding, see https://learn.microsoft.com/graph/api/shares-get
        $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($url))
        $shareId = 'u!' + $b64.TrimEnd('=').Replace('/', '_').Replace('+', '-')
        $folder = Invoke-Graph -Uri "shares/$shareId/driveItem?$select"
    } else {
        $folder = Invoke-Graph -Uri "drives/$DriveId/items/$($ItemId)?$select"
    }
    $script:Warmup = $false
    if (-not $folder.folder) { throw "'$($folder.name)' is not a folder." }

    $targetDriveId = $folder.parentReference.driveId
    $targetItemId  = $folder.id
    if (-not $ShortcutName) { $ShortcutName = $folder.name }

    Write-Host ''
    Write-Host "  Folder   : $($folder.webUrl)" -ForegroundColor Cyan
    Write-Host "  Shortcut : $ShortcutName" -ForegroundColor Cyan

    # ── Users ────────────────────────────────────────────────────────────────
    $users = [System.Collections.Generic.List[string]]::new()
    foreach ($u in @($UserPrincipalName)) { if ($u) { $users.Add($u) } }
    if ($CsvPath) {
        foreach ($row in Import-Csv -LiteralPath $CsvPath) { if ($row.UserPrincipalName) { $users.Add($row.UserPrincipalName) } }
    }
    foreach ($g in @($GroupId | Where-Object { $_ })) {
        $gid = $g
        if ($g -notmatch '^[0-9a-fA-F-]{36}$') {
            $found = @(Get-GraphAll -Uri "groups?`$filter=mail eq '$($g.Replace("'", "''"))'&`$select=id,displayName")
            if ($found.Count -ne 1) {
                throw "No (single) group with mail address '$g' found. A shared mailbox is not a group: pass its users with -UserPrincipalName or -CsvPath."
            }
            $gid = $found[0].id
            Write-Host "  Group    : $g -> $($found[0].displayName)" -ForegroundColor Cyan
        }
        Get-GraphAll -Uri "groups/$gid/transitiveMembers/microsoft.graph.user?`$select=userPrincipalName,accountEnabled&`$top=999" |
            Where-Object accountEnabled |
            ForEach-Object { $users.Add($_.userPrincipalName) }
    }
    $users = @($users | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    Write-Host "  Users    : $($users.Count)" -ForegroundColor Cyan
    Write-Host "  Mode     : $(if ($Apply) { 'APPLY - shortcuts will be created' } else { 'report only (use -Apply to create)' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
    Write-Host ''

    # ── Per user ─────────────────────────────────────────────────────────────
    $results = foreach ($upn in $users) {
        $row = [ordered]@{ UserPrincipalName = $upn; Status = $null; Name = $null; Message = $null }
        try {
            $u = [uri]::EscapeDataString($upn)
            $exists = $true
            try { $null = Invoke-Graph -Uri "users/$($u)?`$select=id" } catch { $exists = $false }
            $hasDrive = $exists
            if ($exists) {
                try { $null = Invoke-Graph -Uri "users/$u/drive?`$select=id" } catch { $hasDrive = $false }
            }

            if (-not $exists) {
                $row.Status = 'UserNotFound'
                $row.Message = 'No user with this UPN in the tenant.'
            } elseif (-not $hasDrive) {
                $row.Status = 'NoOneDrive'
                $row.Message = 'No OneDrive (never signed in, or no license).'
            } else {
                $existing = Get-GraphAll -Uri "users/$u/drive/root/children?`$select=id,name,remoteItem&`$top=999" |
                    Where-Object { $_.remoteItem -and $_.remoteItem.id -eq $targetItemId } |
                    Select-Object -First 1

                if ($existing) {
                    $row.Status = 'AlreadyThere'
                    $row.Name = $existing.name
                } elseif (-not $Apply) {
                    $row.Status = 'WouldCreate'
                    $row.Name = $ShortcutName
                } else {
                    $created = Invoke-Graph -Uri "users/$u/drive/root/children" -Method POST -Body @{
                        name                                = $ShortcutName
                        remoteItem                          = @{ id = $targetItemId; parentReference = @{ driveId = $targetDriveId } }
                        '@microsoft.graph.conflictBehavior' = 'rename'
                    }
                    $row.Status = 'Created'
                    $row.Name = $created.name
                }
            }
        } catch {
            $row.Status = 'Failed'
            $row.Message = $_.Exception.Message
        }

        $color = switch ($row.Status) { 'Created' { 'Green' } 'WouldCreate' { 'Yellow' } 'AlreadyThere' { 'DarkGray' } default { 'Red' } }
        Write-Host ("  {0,-13} {1}{2}" -f $row.Status, $upn, $(if ($row.Message) { " - $($row.Message)" })) -ForegroundColor $color
        [pscustomobject]$row
    }

    # ── Output ───────────────────────────────────────────────────────────────
    if (-not $OutputPath) {
        $outDir = if ($IsWindows) { 'C:\Temp' } else { "$HOME/Downloads" }
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
        $OutputPath = Join-Path $outDir "OneDriveShortcut_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    }
    @($results) | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host ''
    @($results) | Group-Object Status | ForEach-Object { Write-Host ("  {0,-13} {1}" -f $_.Name, $_.Count) }
    Write-Host "  CSV: $((Resolve-Path $OutputPath).Path)" -ForegroundColor DarkGray
} finally {
    Remove-TempApp
    Disconnect-M365Graph $graph
}
