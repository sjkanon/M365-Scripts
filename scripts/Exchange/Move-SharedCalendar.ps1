#Requires -Version 5.1
<#
.SYNOPSIS
    All in one: find a shared calendar by keyword, move it into a resource mailbox
    with every item and permission, and list who has to switch. One sign-in.

.DESCRIPTION
    Chains the two calendar scripts in this folder:

      1. Find   Get-CalendarMappings.ps1 -Search <keyword>: where the calendar
                lives, who has it in their calendar list, who has rights on it.
      2. Pick   The matching calendar that can be moved: one that lives in
                somebody's mailbox next to their main calendar. With several
                matches you pick one, or narrow it down with -Owner.
      3. Move   Convert-SharedCalendarToResource.ps1 on that calendar. Without
                -Apply it shows the preview and then asks whether to go ahead.
      4. Tell   Who had the old calendar in Outlook and who only had rights: the
                people who have to switch to the new one.

    Signing in happens once. Unless an app-only Graph session or -ClientId with
    -ClientSecret is supplied, a temporary App Registration with everything both
    scripts need (Calendars.ReadWrite, User.Read.All, Group.Read.All,
    MailboxSettings.ReadWrite) is created, handed to both, and removed at the end -
    also when something fails. Exchange Online is connected once as well.

    A mailbox whose MAIN calendar matches (a balie@ account that is itself the
    shared calendar) cannot be moved out. The script says so and names the
    in-place alternative, which keeps everything: Set-Mailbox -Type Room.

    The resource options (-ResourceType, -ResourceName, -ResourceAddress,
    -SourceOwnerRights) are passed to Convert-SharedCalendarToResource.ps1 only
    when given, so its defaults apply otherwise: a Room mailbox named after the
    calendar, the original owner as Owner.

    A run that changes something is logged to SharedCalendarMove_<timestamp>.log
    next to the mapping report and the backup.

.PARAMETER Search
    Keyword to find the calendar by: owner name or address, or calendar name.
    Alias: -Keyword.

.PARAMETER Owner
    Narrows several matches down to the calendar of this owner (part of the
    name or address is enough).

.PARAMETER ResourceType
    Room or Equipment. Default (of the convert script): Room.

.PARAMETER ResourceName
    Display name of the new mailbox. Default: the calendar's name.

.PARAMETER ResourceAddress
    SMTP address of the new mailbox. Default: the name as alias at the owner's
    domain.

.PARAMETER SourceOwnerRights
    Rights of the original owner on the new calendar: Owner (default),
    PublishingEditor, Editor, Reviewer or None. Use None for an archived owner.

.PARAMETER SendSharingInvitation
    Send users a sharing invitation for the new calendar. Asked at the prompt
    when not given.

.PARAMETER Apply
    Go ahead without the preview-and-ask round.

.PARAMETER RemoveSourceCalendar
    Remove the original calendar after a clean verification. Asked at the
    prompt when not given.

.PARAMETER Force
    Skip the typed confirmation before removal - needed to remove unattended.

.PARAMETER OutputPath
    Folder for the mapping report, the backup and the log. Default: C:\Temp
    (~/Downloads on macOS/Linux).

.PARAMETER TenantId
    Tenant ID or domain. Default: the tenant of the Exchange session.

.PARAMETER ClientId
    Your own App Registration (with the four permissions above) instead of a
    temporary one. Needs -ClientSecret.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.EXAMPLE
    # Find "balie", preview the move, answer the questions
    .\Move-SharedCalendar.ps1 -Search balie

.EXAMPLE
    # Straight through for one owner's calendar, as an Equipment mailbox,
    # with invitations, keeping the original for now
    .\Move-SharedCalendar.ps1 -Search "balie planning" -Owner peggy -ResourceType Equipment `
        -SourceOwnerRights None -SendSharingInvitation -Apply

.NOTES
    Author: Sjoerd Kanon
    Needs Get-CalendarMappings.ps1 and Convert-SharedCalendarToResource.ps1 in the
    same folder.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [Alias('Keyword')]
    [string] $Search,
    [string] $Owner,
    [ValidateSet('Room', 'Equipment')]
    [string] $ResourceType,
    [string] $ResourceName,
    [string] $ResourceAddress,
    [ValidateSet('Owner', 'PublishingEditor', 'Editor', 'Reviewer', 'None')]
    [string] $SourceOwnerRights,
    [switch] $SendSharingInvitation,
    [switch] $Apply,
    [switch] $RemoveSourceCalendar,
    [switch] $Force,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret
)

$mappingScript = Join-Path $PSScriptRoot 'Get-CalendarMappings.ps1'
$convertScript = Join-Path $PSScriptRoot 'Convert-SharedCalendarToResource.ps1'
foreach ($s in @($mappingScript, $convertScript)) {
    if (-not (Test-Path $s)) { throw "$(Split-Path $s -Leaf) was not found next to this script - the calendar scripts have to sit in the same folder." }
}

# -- Output -------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ""; Write-Host "  == $Message" -ForegroundColor Magenta }
function Write-Ok   { param([string] $Message) Write-Host "  [OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Info { param([string] $Message) Write-Host "         $Message" -ForegroundColor DarkGray }

function Test-CanPrompt {
    # Read-Host needs a person: not in a service, not under -NonInteractive.
    if (-not [Environment]::UserInteractive) { return $false }
    return -not ([Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^-noni' })
}

$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Move Shared Calendar" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Search    : '$Search'$(if ($Owner) { " (owner: $Owner)" })" -ForegroundColor DarkGray
Write-Host "  Mode      : $(if ($Apply) { 'APPLY' } else { 'preview first, then ask' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ==============================================================================
#  One temporary App Registration for both scripts
# ==============================================================================
$AllRoles                = @('Calendars.ReadWrite', 'User.Read.All', 'Group.Read.All', 'MailboxSettings.ReadWrite')
$script:GraphCliClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$script:AdminHeaders     = $null
$script:TempAppObjectId  = $null

function Invoke-GraphAdmin {
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, $Body)
    $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AdminHeaders; ErrorAction = 'Stop' }
    if ($Body) {
        $params['Body']        = ($Body | ConvertTo-Json -Depth 6)
        $params['ContentType'] = 'application/json'
    }
    return Invoke-RestMethod @params
}

function Remove-TempApp {
    if (-not $script:TempAppObjectId) { return }
    Write-Host "  Removing temporary App Registration..." -ForegroundColor DarkGray
    try {
        Invoke-GraphAdmin -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$($script:TempAppObjectId)" | Out-Null
        Write-Host "  [OK]   Temporary App Registration removed." -ForegroundColor DarkGray
    } catch {
        Write-Warning "Could not remove the temporary App Registration (object ID $($script:TempAppObjectId)). Remove it by hand in Entra ID > App registrations."
    }
    $script:TempAppObjectId = $null
}

function Get-TokenClaim {
    param([string] $Jwt)
    try {
        $payload = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) {
            2 { $payload += '==' }
            3 { $payload += '=' }
        }
        return ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload)) | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Get-DelegatedTokenByDeviceCode {
    param([string] $Tenant, [string[]] $Scopes)

    $scopeString = ((@($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" })) + 'offline_access') -join ' '
    $dc = Invoke-RestMethod -Method POST -ErrorAction Stop `
            -Uri  "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
            -Body @{ client_id = $script:GraphCliClientId; scope = $scopeString }

    Write-Host ""
    Write-Host "  ------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "   $($dc.message)" -ForegroundColor Yellow
    Write-Host "  ------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Waiting for sign-in..." -ForegroundColor DarkGray

    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $interval = [Math]::Max(5, [int]$dc.interval)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method POST -ErrorAction Stop `
                    -Uri  "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
                    -Body @{
                        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                        client_id   = $script:GraphCliClientId
                        device_code = $dc.device_code
                    }
            return $tok.access_token
        } catch {
            $code = ''
            try { $code = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}
            if ($code -eq 'authorization_pending')  { continue }
            if ($code -eq 'slow_down')              { $interval += 5; continue }
            if ($code -eq 'expired_token')          { throw "Device code expired before sign-in completed." }
            if ($code -eq 'authorization_declined') { throw "Sign-in was declined." }
            throw
        }
    }
    throw "Device code sign-in timed out."
}

function New-TemporaryApp {
    <#
        Creates the App Registration both scripts will use, and only hands it
        over once a fresh token actually carries all four roles - role
        assignments take a while to reach the token service, and each script
        mints its own token.
    #>
    param([string] $Tenant)

    Write-Host "  Setting up one temporary App Registration for the whole run." -ForegroundColor Cyan
    Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time)" -ForegroundColor DarkGray

    $adminToken = Get-DelegatedTokenByDeviceCode -Tenant $Tenant -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
    $script:AdminHeaders = @{ Authorization = "Bearer $adminToken" }
    Write-Host "  [OK]   Signed in." -ForegroundColor DarkGray

    $appName = "CalendarMove-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')"
    $app = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' -Body @{ displayName = $appName; signInAudience = 'AzureADMyOrg' }
    $script:TempAppObjectId = $app.id
    $sp = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{ appId = $app.appId }
    $graphSp = @((Invoke-GraphAdmin -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value)[0]
    if (-not $graphSp) { throw "Could not resolve the Microsoft Graph service principal." }

    foreach ($roleName in $AllRoles) {
        $appRole = @($graphSp.appRoles | Where-Object { $_.value -eq $roleName -and $_.allowedMemberTypes -contains 'Application' })[0]
        if (-not $appRole) { throw "Could not resolve the $roleName application role." }
        Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" `
            -Body @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $appRole.id } | Out-Null
    }
    Write-Host "  [OK]   '$appName' created with $($AllRoles -join ', ')." -ForegroundColor DarkGray

    # Long enough for a big calendar; the app is deleted at the end anyway.
    $secret = Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
                -Body @{ passwordCredential = @{ displayName = 'temp'; endDateTime = (Get-Date).AddHours(8).ToString('o') } }

    $deadline = (Get-Date).AddMinutes(4)
    $reported = $false
    while ($true) {
        $missing = $AllRoles
        try {
            $tok = (Invoke-RestMethod -Method POST -ErrorAction Stop -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body @{
                grant_type = 'client_credentials'; scope = 'https://graph.microsoft.com/.default'; client_id = $app.appId; client_secret = $secret.secretText
            }).access_token
            $roles   = @((Get-TokenClaim -Jwt $tok).roles)
            $missing = @($AllRoles | Where-Object { $roles -notcontains $_ })
        } catch { }
        if ($missing.Count -eq 0) { break }
        if ((Get-Date) -ge $deadline) { throw "The temporary app's token still lacks $($missing -join ', ') after 4 minutes." }
        if (-not $reported) { Write-Host "  Waiting for the app and its permissions to propagate..." -ForegroundColor DarkGray; $reported = $true }
        Start-Sleep -Seconds 10
    }
    Write-Host "  [OK]   App-only token carries every permission." -ForegroundColor DarkGray
    return [PSCustomObject]@{ AppId = $app.appId; Secret = $secret.secretText }
}

# ==============================================================================
#  Run
# ==============================================================================
$script:ConnectedExo = $false
$transcribing = $false
$log = Join-Path $outputDir "SharedCalendarMove_$stamp.log"
try {
    # -- Exchange Online, once -------------------------------------------------------
    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        throw "The ExchangeOnlineManagement module is required: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    try {
        $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
    } catch {
        $connectParams = @{ ShowBanner = $false }
        if ($TenantId) { $connectParams['Organization'] = $TenantId }
        Connect-ExchangeOnline @connectParams
        $script:ConnectedExo = $true
    }

    $tenant = $TenantId
    if (-not $tenant) {
        try {
            $gdap = ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP') -or
                    ($env:M365_AUTH_MODE -and ([string]$env:M365_AUTH_MODE).ToUpperInvariant() -eq 'GDAP')
            if ($gdap -and $global:cid)          { $tenant = [string]$global:cid }
            elseif ($env:M365_CUSTOMER_TENANTID) { $tenant = [string]$env:M365_CUSTOMER_TENANTID }
        } catch {}
    }
    if (-not $tenant) {
        try { $tenant = [string](@(Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })[0].TenantID) } catch {}
    }

    # -- Graph, once -------------------------------------------------------------
    $graph = @{}
    $ctx = $null
    if (-not $ClientId) { try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {} }
    if ($ClientId) {
        if (-not $ClientSecret) { throw "-ClientId needs -ClientSecret here: both scripts take their token over REST." }
        if (-not $tenant)       { throw "-ClientId needs -TenantId." }
        $graph = @{ TenantId = $tenant; ClientId = $ClientId; ClientSecret = $ClientSecret }
    } elseif ($ctx -and $ctx.AuthType -eq 'AppOnly' -and @($ctx.Scopes) -contains 'Calendars.ReadWrite' -and @($ctx.Scopes) -contains 'User.Read.All') {
        Write-Ok "Using the existing app-only Graph session for both steps."
    } else {
        if (-not $tenant) { throw "No tenant known for the temporary App Registration. Pass -TenantId." }
        $app   = New-TemporaryApp -Tenant $tenant
        $graph = @{ TenantId = $tenant; ClientId = $app.AppId; ClientSecret = $app.Secret }
    }

    # -- 1. Find ---------------------------------------------------------------
    Write-Step "1. Find '$Search'"
    $csv = Join-Path $outputDir "CalendarMappings_$stamp.csv"
    & $mappingScript -Search $Search -OutputPath $csv @graph
    if (-not (Test-Path $csv)) {
        Write-Warn "No calendar or owner matches '$Search' - nothing to move."
        return
    }
    $rows = @(Import-Csv -Path $csv -Encoding UTF8)

    # -- 2. Pick ---------------------------------------------------------------
    Write-Step '2. Pick the calendar'
    $sources = @($rows | Where-Object { $_.Status -eq 'Source' })
    foreach ($m in @($sources | Where-Object { $_.Calendar -eq 'Main' })) {
        Write-Info "$($m.Owner) <$($m.OwnerAddress)> matches as a whole mailbox. Its main calendar cannot be moved out;"
        Write-Info "  convert that mailbox in place instead, which keeps everything: Set-Mailbox $($m.OwnerAddress) -Type Room"
    }
    $movable = @($sources | Where-Object { $_.Calendar -ne 'Main' })
    if ($Owner) { $movable = @($movable | Where-Object { $_.OwnerAddress -like "*$Owner*" -or $_.Owner -like "*$Owner*" }) }
    if ($movable.Count -eq 0) {
        Write-Warn "No calendar that can be moved matches '$Search'$(if ($Owner) { " for owner '$Owner'" })."
        return
    }

    function Get-CalendarRow {
        # Every row of the mapping report about one calendar, except its Source row.
        param($Source)
        return @($rows | Where-Object { $_.OwnerAddress -eq $Source.OwnerAddress -and $_.Calendar -eq $Source.Calendar -and $_.Status -ne 'Source' })
    }
    function Format-Candidate {
        param($Source)
        $r = Get-CalendarRow -Source $Source
        $inOutlook = @($r | Where-Object { $_.Mapped -eq 'Yes' }).Count
        $withRight = @($r | Where-Object { $_.Rights -and $_.UserType -eq 'Mailbox' }).Count
        return "'$($Source.Calendar)' in $($Source.Owner) <$($Source.OwnerAddress)> - $inOutlook in Outlook, $withRight with rights"
    }

    if ($movable.Count -eq 1) {
        $pick = $movable[0]
    } elseif (-not (Test-CanPrompt)) {
        $list = ($movable | ForEach-Object { "'$($_.Calendar)' ($($_.OwnerAddress))" }) -join ', '
        throw "'$Search' matches $($movable.Count) calendars that can be moved: $list. Narrow it down with -Owner or a more specific -Search."
    } else {
        for ($i = 0; $i -lt $movable.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), (Format-Candidate -Source $movable[$i])) }
        $answer = Read-Host "  Which calendar? [1-$($movable.Count)]"
        $n = 0
        if (-not [int]::TryParse($answer, [ref]$n) -or $n -lt 1 -or $n -gt $movable.Count) {
            Write-Warn "No valid choice - nothing was changed."
            return
        }
        $pick = $movable[$n - 1]
    }
    Write-Ok (Format-Candidate -Source $pick)

    # -- 3. Move ---------------------------------------------------------------
    $convert = @{ Mailbox = $pick.OwnerAddress; Calendar = $pick.Calendar; PassThru = $true } + $graph
    foreach ($name in @('ResourceType', 'ResourceName', 'ResourceAddress', 'SourceOwnerRights')) {
        if ($PSBoundParameters.ContainsKey($name)) { $convert[$name] = $PSBoundParameters[$name] }
    }
    if ($OutputPath) { $convert['BackupPath'] = Join-Path $OutputPath "CalendarConvert_$stamp" }

    $doApply = [bool]$Apply
    $invite  = [bool]$SendSharingInvitation
    $remove  = [bool]$RemoveSourceCalendar
    if (-not $doApply) {
        Write-Step '3. Move - preview'
        $null = & $convertScript @convert
        if (Test-CanPrompt) {
            $doApply = (Read-Host "  Create the resource mailbox and copy everything now? [y/N]") -match '^[Yy]'
            if ($doApply -and -not $PSBoundParameters.ContainsKey('SendSharingInvitation')) {
                $invite = (Read-Host "  Send users a sharing invitation for the new calendar? [Y/n]") -notmatch '^[Nn]'
            }
            if ($doApply -and -not $PSBoundParameters.ContainsKey('RemoveSourceCalendar')) {
                $remove = (Read-Host "  Remove the original calendar once every item is verified? [y/N]") -match '^[Yy]'
            }
        }
    }

    $result = $null
    if ($doApply) {
        try { Start-Transcript -Path $log -Append | Out-Null; $transcribing = $true } catch { Write-Warn "No log: $($_.Exception.Message)" }
        Write-Step '3. Move'
        $convert['Apply'] = $true
        if ($invite) { $convert['SendSharingInvitation'] = $true }
        if ($remove) { $convert['RemoveSourceCalendar'] = $true }
        if ($Force)  { $convert['Force'] = $true }
        $result = @(& $convertScript @convert) | Where-Object { $_ -and $_.PSObject.Properties['SourceRemoved'] } | Select-Object -Last 1
    }

    # -- 4. Tell ---------------------------------------------------------------
    Write-Step '4. Who has to switch'
    $calendarRows = Get-CalendarRow -Source $pick
    $inOutlook  = @($calendarRows | Where-Object { $_.Mapped -eq 'Yes' -and $_.UserType -eq 'Mailbox' })
    $rightsOnly = @($calendarRows | Where-Object { $_.Mapped -ne 'Yes' -and $_.Rights -and $_.UserType -eq 'Mailbox' })
    if ($inOutlook.Count -gt 0)  { Write-Info "Had it in Outlook ($($inOutlook.Count)): $(($inOutlook | ForEach-Object { $_.User }) -join ', ')" }
    if ($rightsOnly.Count -gt 0) { Write-Info "Had rights, not in Outlook ($($rightsOnly.Count)): $(($rightsOnly | ForEach-Object { $_.User }) -join ', ')" }

    if (-not $doApply) {
        Write-Info "Nothing was changed. Run again with -Apply, or answer yes at the question."
    } elseif ($result) {
        Write-Ok "New calendar: $($result.ResourceName) <$($result.ResourceAddress)> ($($result.ResourceType) mailbox, $($result.Items) item(s))"
        if ($invite) { Write-Info "They received a sharing invitation - accepting it adds the new calendar." }
        else         { Write-Info "They add it via Calendar > Add calendar > From directory > '$($result.ResourceName)'." }
        if ($result.SourceRemoved) {
            Write-Info "The original calendar is removed: its entry in their list no longer works and can be deleted."
        } elseif ($result.Verified) {
            Write-Info "The original calendar is still there. Once everyone has switched, run the same command with -Apply -RemoveSourceCalendar."
        } else {
            Write-Warn "The copy is not complete - see 'Needs attention' above. Run the same command again; finished items are skipped."
        }
    }
    Write-Host ""
} finally {
    Remove-TempApp
    if ($transcribing) {
        try { Stop-Transcript | Out-Null } catch {}
        Write-Host "  Log: $log" -ForegroundColor DarkGray
    }
    if ($script:ConnectedExo) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
}
