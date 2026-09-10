#Requires -Version 7.0
<#
.SYNOPSIS
    Shared helpers for the SharePoint structure scripts in this folder. Dot-sourced,
    never run on its own.

.DESCRIPTION
    The four scripts here (metadata, libraries/permissions, share-status audit and
    drift check) all speak the same config file and need the same three things: a
    validated configuration, a PnP connection that works both interactively and
    app-only, and a handful of CSOM primitives PnP does not expose directly
    (field-link Required, breaking inheritance on a securable object).

    Putting that in one file is a deliberate exception to the "every script stands
    alone" rule in this repo: these four form one set around one config schema, and
    three copies of the permission code would drift apart within a month.

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x, and Microsoft.Graph.Groups for -EnsureGroups
#>

Set-StrictMode -Version Latest

# -- Output --------------------------------------------------------------------
# Same vocabulary in all four scripts, so a run reads the same whichever one you
# started: [ OK ] nothing to do, [ >> ] something changed, [DIFF] drift found.

function Write-Head   { param([string] $Message) Write-Host ''; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Step   { param([string] $Message) Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok     { param([string] $Message) Write-Host "    [ OK ] $Message" -ForegroundColor Green }
function Write-Change { param([string] $Message) Write-Host "    [ >> ] $Message" -ForegroundColor Magenta }
function Write-Skip   { param([string] $Message) Write-Host "    [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn   { param([string] $Message) Write-Host "    [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad    { param([string] $Message) Write-Host "    [FAIL] $Message" -ForegroundColor Red }
function Write-Diff   { param([string] $Message) Write-Host "    [DIFF] $Message" -ForegroundColor Yellow }

function Get-ConfigValue {
    <#
        Property access on a ConvertFrom-Json object that returns $null instead of
        tripping Set-StrictMode when the key is simply absent from the file.
    #>
    param($Object, [string] $Name, $Default = $null)

    if ($null -eq $Object) { return $Default }
    if ($Object -isnot [psobject]) { return $Default }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $Default }
    $value = $Object.$Name
    if ($null -eq $value) { return $Default }
    return $value
}

# -- Configuration -------------------------------------------------------------
function Import-StructureConfig {
    <#
        Read the JSON model and refuse to hand back anything that would only fail
        halfway through a provisioning run: unreplaced placeholders, a content type
        pointing at a column that is not defined, a container pointing at a group
        that does not exist in the model.
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [switch] $AllowPlaceholders
    )

    if (-not (Test-Path $Path)) { throw "Configuration file not found: $Path" }

    try {
        $config = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "Configuration file is not valid JSON ($Path): $($_.Exception.Message)"
    }

    foreach ($required in @('tenant', 'sites', 'columns', 'contentTypes', 'containers')) {
        if (-not (Get-ConfigValue $config $required)) {
            throw "Configuration is missing the '$required' section ($Path)."
        }
    }

    # CHANGEME is what ships in the repo. Catching it here beats a sign-in that
    # fails with a tenant-not-found five minutes into the run.
    if (-not $AllowPlaceholders) {
        $placeholders = @()
        if ($config.tenant -match 'CHANGEME') { $placeholders += 'tenant' }
        foreach ($property in $config.sites.PSObject.Properties) {
            if ($property.Value -match 'CHANGEME') { $placeholders += "sites.$($property.Name)" }
        }
        if ($placeholders.Count -gt 0) {
            throw ("Fill in the tenant and site URLs first - still on CHANGEME: {0} ({1})" -f ($placeholders -join ', '), $Path)
        }
    }

    $columnNames = @($config.columns | ForEach-Object { $_.internalName })
    foreach ($column in $config.columns) {
        if ($column.type -eq 'Taxonomy' -and -not (Get-ConfigValue $config 'termStore')) {
            throw "Column '$($column.internalName)' is a Taxonomy column but the configuration has no termStore section."
        }
        if ($column.type -in @('Choice', 'MultiChoice') -and -not (Get-ConfigValue $column 'choices')) {
            throw "Column '$($column.internalName)' is a $($column.type) column without any choices."
        }
    }

    $contentTypeNames = @($config.contentTypes | ForEach-Object { $_.name })
    foreach ($contentType in $config.contentTypes) {
        foreach ($field in $contentType.fields) {
            if ($field.internalName -notin $columnNames) {
                throw "Content type '$($contentType.name)' refers to column '$($field.internalName)', which is not defined in the columns section."
            }
        }
    }

    $groupNames = @(Get-ConfigValue $config 'groups' @() | ForEach-Object { $_.displayName })
    foreach ($container in $config.containers) {
        if ((Get-ConfigValue $container 'site') -notin $config.sites.PSObject.Properties.Name) {
            throw "Container '$($container.key)' points at site '$(Get-ConfigValue $container 'site')', which is not defined in the sites section."
        }
        foreach ($name in (Get-ConfigValue $container 'contentTypes' @())) {
            if ($name -notin $contentTypeNames) {
                throw "Container '$($container.key)' binds content type '$name', which is not defined in the contentTypes section."
            }
        }
        foreach ($permission in (Get-ConfigValue $container 'permissions' @())) {
            if ($permission.group -notin $groupNames) {
                throw "Container '$($container.key)' grants '$($permission.group)', which is not defined in the groups section."
            }
        }
    }

    return $config
}

function Get-StructureSiteUrl {
    <# The site URL a container lives on, trailing slash removed. #>
    param([Parameter(Mandatory)] $Config, [Parameter(Mandatory)] [string] $SiteKey)

    $url = Get-ConfigValue $Config.sites $SiteKey
    if (-not $url) { throw "No site URL configured for '$SiteKey'." }
    return $url.TrimEnd('/')
}

# -- Connection ----------------------------------------------------------------
function Connect-Structure {
    <#
        One connection per site URL, cached for the run. Interactive by default;
        app-only when a certificate thumbprint or a client secret is supplied, which
        is what the scheduled share-status audit uses.

        Returns a PnP connection object - every call below passes it explicitly, so
        two site collections (team site and private-channel site) can be worked on in
        the same run without reconnecting all the time.
    #>
    param(
        [Parameter(Mandatory)] [string] $Url,
        [Parameter(Mandatory)] [string] $Tenant,
        [string] $ClientId,
        [string] $Thumbprint,
        [string] $CertificatePath,
        [securestring] $CertificatePassword,
        [switch] $Interactive
    )

    if (-not $script:StructureConnections) { $script:StructureConnections = @{} }
    $key = $Url.TrimEnd('/').ToLowerInvariant()
    if ($script:StructureConnections.ContainsKey($key)) { return $script:StructureConnections[$key] }

    $appOnly = [bool]($Thumbprint -or $CertificatePath)
    if ($appOnly -and -not $ClientId) {
        throw 'App-only sign-in needs -ClientId together with the certificate.'
    }
    if (-not $appOnly -and -not $Interactive -and -not $ClientId) {
        throw 'Pass -Interactive (with -ClientId of your PnP app) or a certificate for app-only sign-in.'
    }

    $splat = @{ Url = $Url; ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($appOnly) {
        $splat['ClientId'] = $ClientId
        $splat['Tenant']   = $Tenant
        if ($Thumbprint) {
            $splat['Thumbprint'] = $Thumbprint
        } else {
            $splat['CertificatePath'] = $CertificatePath
            if ($CertificatePassword) { $splat['CertificatePassword'] = $CertificatePassword }
        }
    } else {
        $splat['Interactive'] = $true
        if ($ClientId) { $splat['ClientId'] = $ClientId }
    }

    # A freshly created app registration is not replicated everywhere yet, so the
    # first sign-in against a new tenant can fail with "application not found".
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $connection = Connect-PnPOnline @splat
            $script:StructureConnections[$key] = $connection
            Write-Ok ("Connected to $Url ({0})" -f $(if ($appOnly) { 'app-only' } else { 'interactive' }))
            return $connection
        } catch {
            if ($attempt -eq 4) { throw "Could not connect to ${Url}: $($_.Exception.Message)" }
            Write-Warn "Sign-in attempt $attempt failed ($($_.Exception.Message.Trim())) - retrying in 10s..."
            Start-Sleep -Seconds 10
        }
    }
}

function Disconnect-Structure {
    <# Close every cached connection - only worth calling at the end of a run. #>
    if (-not $script:StructureConnections) { return }
    foreach ($connection in $script:StructureConnections.Values) {
        try { Disconnect-PnPOnline -Connection $connection -ErrorAction SilentlyContinue } catch { }
    }
    $script:StructureConnections = @{}
}

function Assert-PnPModule {
    if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
        throw "Module 'PnP.PowerShell' is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
}

# -- CSOM primitives -----------------------------------------------------------
# PnP covers most of the model, but not these three. Doing them through CSOM keeps
# the scripts idempotent instead of "remove and re-add and hope".

function Set-FieldLinkRequired {
    <#
        Flip the Required flag on a field link that is already part of a content
        type. Add-PnPFieldToContentType -Required only applies when the link is
        created, so without this a second run can never repair a field that was made
        optional by hand.

        Returns $true when something was actually changed.
    #>
    param(
        [Parameter(Mandatory)] $ContentType,
        [Parameter(Mandatory)] [string] $InternalName,
        [Parameter(Mandatory)] [bool] $Required,
        [Parameter(Mandatory)] $Connection,
        [switch] $WhatIfMode
    )

    $context = Get-PnPContext -Connection $Connection
    $context.Load($ContentType.FieldLinks)
    Invoke-PnPQuery -Connection $Connection

    $link = $ContentType.FieldLinks | Where-Object { $_.Name -eq $InternalName }
    if (-not $link) { return $false }
    if ($link.Required -eq $Required) { return $false }
    if ($WhatIfMode) { return $true }

    $link.Required = $Required
    # $true: push the change down to content types that inherit from this one.
    $ContentType.Update($true)
    Invoke-PnPQuery -Connection $Connection
    return $true
}

function Get-FieldChoiceValue {
    <#
        The choices of a Choice/MultiChoice site column, read from the schema XML.
        The typed Choices property needs a cast to FieldChoice and an extra round
        trip; the schema is already on the object Get-PnPField handed back.
    #>
    param([Parameter(Mandatory)] $Field)

    try {
        $schema = [xml] $Field.SchemaXml
    } catch {
        return @()
    }
    $node = $schema.Field.SelectSingleNode('CHOICES')
    if (-not $node) { return @() }
    return @($node.ChildNodes | ForEach-Object { $_.InnerText })
}

function Get-SecurablePrincipal {
    <#
        Resolve an Entra ID security group to a SharePoint principal.

        SharePoint knows Entra groups by claim, not by name: c:0t.c|tenant|<objectid>
        for a security group. EnsureUser on that claim creates the site-level user
        entry when it does not exist yet, which is what makes a brand new group
        assignable without anyone opening the site first.
    #>
    param(
        [Parameter(Mandatory)] [string] $ObjectId,
        [Parameter(Mandatory)] $Connection
    )

    $claim = "c:0t.c|tenant|$ObjectId"
    $web   = Get-PnPWeb -Connection $Connection -Includes SiteUsers

    $context   = Get-PnPContext -Connection $Connection
    $principal = $web.EnsureUser($claim)
    $context.Load($principal)
    Invoke-PnPQuery -Connection $Connection
    return $principal
}

function Set-SecurableRole {
    <#
        Bring the role assignments on a list or a folder in line with what the config
        says, and report per change what happened.

        Breaks inheritance first when needed. -KeepExisting copies the inherited
        assignments across (the default, and the safe choice inside a Teams library);
        without it the securable starts clean and only the configured groups plus the
        site's own owners survive.

        Returns a list of change descriptions - empty means it was already right.
    #>
    param(
        [Parameter(Mandatory)] $Securable,
        [Parameter(Mandatory)] [array] $Desired,
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] [hashtable] $PrincipalMap,
        [switch] $KeepExisting,
        [switch] $RemoveOther,
        [switch] $WhatIfMode
    )

    $changes  = @()
    $context  = Get-PnPContext -Connection $Connection
    $web      = Get-PnPWeb -Connection $Connection

    $context.Load($Securable)
    $context.Load($Securable.RoleAssignments)
    try {
        $context.ExecuteQuery()
    } catch {
        throw "Could not read the current permissions: $($_.Exception.Message)"
    }

    if (-not $Securable.HasUniqueRoleAssignments) {
        $changes += "break inheritance (keep existing: $([bool]$KeepExisting))"
        if (-not $WhatIfMode) {
            $Securable.BreakRoleInheritance([bool]$KeepExisting, $false)
            $context.ExecuteQuery()
            $context.Load($Securable.RoleAssignments)
            $context.ExecuteQuery()
        }
    }

    # What is on there now, per login name, so a second run is a no-op.
    $current = @{}
    if (-not ($WhatIfMode -and $changes.Count -gt 0)) {
        foreach ($assignment in $Securable.RoleAssignments) {
            $context.Load($assignment.Member)
            $context.Load($assignment.RoleDefinitionBindings)
        }
        $context.ExecuteQuery()
        foreach ($assignment in $Securable.RoleAssignments) {
            $current[$assignment.Member.LoginName] = @{
                Roles  = @($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name })
                Member = $assignment.Member
            }
        }
    }

    $desiredLogins = @()
    foreach ($entry in $Desired) {
        $principal = $PrincipalMap[$entry.group]
        if (-not $principal) {
            $changes += "SKIPPED $($entry.group) - group not resolved"
            continue
        }
        $desiredLogins += $principal.LoginName

        $existing = $current[$principal.LoginName]
        if ($existing -and $entry.role -in $existing.Roles) { continue }

        $changes += "grant $($entry.role) to $($entry.group)"
        if ($WhatIfMode) { continue }

        $role     = $web.RoleDefinitions.GetByName($entry.role)
        $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
        $bindings.Add($role)
        $Securable.RoleAssignments.Add($principal, $bindings) | Out-Null
        $context.ExecuteQuery()
    }

    if ($RemoveOther) {
        foreach ($login in $current.Keys) {
            if ($login -in $desiredLogins) { continue }
            # Site owners/admins keep their access - removing those locks the client
            # out of their own library.
            if ($login -match 'Owners|Eigenaren|SharingLinks|spo-grid-all-users|_spOwner') { continue }

            $changes += "remove $login"
            if ($WhatIfMode) { continue }
            $assignment = $Securable.RoleAssignments | Where-Object { $_.Member.LoginName -eq $login } | Select-Object -First 1
            if ($assignment) {
                $assignment.DeleteObject()
                $context.ExecuteQuery()
            }
        }
    }

    return $changes
}

function Get-StructureRoleReport {
    <#
        The role assignments on a securable, flattened to one row per principal.
        Used by the drift check and by the -Report output of the permission script.
    #>
    param(
        [Parameter(Mandatory)] $Securable,
        [Parameter(Mandatory)] $Connection
    )

    $context = Get-PnPContext -Connection $Connection
    $context.Load($Securable.RoleAssignments)
    $context.ExecuteQuery()

    foreach ($assignment in $Securable.RoleAssignments) {
        $context.Load($assignment.Member)
        $context.Load($assignment.RoleDefinitionBindings)
    }
    $context.ExecuteQuery()

    foreach ($assignment in $Securable.RoleAssignments) {
        $roles = @($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name } | Where-Object { $_ -ne 'Limited Access' })
        if ($roles.Count -eq 0) { continue }
        [PSCustomObject]@{
            Principal = $assignment.Member.Title
            LoginName = $assignment.Member.LoginName
            Roles     = ($roles -join ', ')
        }
    }
}

function Get-ChannelFolderItem {
    <#
        The list item behind a channel folder, which is the securable object you need
        to set folder permissions. Returns $null when the folder is not there.
    #>
    param(
        [Parameter(Mandatory)] [string] $ListTitle,
        [Parameter(Mandatory)] [string] $FolderName,
        [Parameter(Mandatory)] $Connection
    )

    $list = Get-PnPList -Identity $ListTitle -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $list) { return $null }

    $context = Get-PnPContext -Connection $Connection
    $context.Load($list.RootFolder)
    $context.ExecuteQuery()

    $serverRelativeUrl = "$($list.RootFolder.ServerRelativeUrl.TrimEnd('/'))/$FolderName"
    try {
        $folder = Get-PnPFolder -Url $serverRelativeUrl -Connection $Connection -ErrorAction Stop
    } catch {
        return $null
    }

    $context.Load($folder.ListItemAllFields)
    $context.ExecuteQuery()
    return $folder.ListItemAllFields
}

# -- Entra ID groups -----------------------------------------------------------
function Connect-StructureGraph {
    <#
        Graph sign-in for the group work. Uses the same certificate as PnP when one
        is given, so an unattended run needs one app registration, not two.
    #>
    param(
        [Parameter(Mandatory)] [string] $Tenant,
        [string] $ClientId,
        [string] $Thumbprint,
        [string[]] $Scopes = @('Group.ReadWrite.All', 'Directory.Read.All')
    )

    foreach ($module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Module '$module' is required for the Entra ID groups. Run: Install-Module $module -Scope CurrentUser"
        }
    }
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop

    if ($Thumbprint -and $ClientId) {
        Connect-MgGraph -TenantId $Tenant -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome | Out-Null
    } else {
        Connect-MgGraph -TenantId $Tenant -Scopes $Scopes -NoWelcome -ContextScope Process | Out-Null
    }
    Write-Ok "Graph connected as $((Get-MgContext).Account ?? 'app-only')"
}

function Resolve-StructureGroup {
    <#
        Look up an Entra ID security group by display name, optionally creating it.
        Returns the group object, or $null when it does not exist and -Create was not
        given.
    #>
    param(
        [Parameter(Mandatory)] $Definition,
        [switch] $Create,
        [switch] $WhatIfMode
    )

    $escaped = $Definition.displayName -replace "'", "''"
    $group   = @(Get-MgGroup -Filter "displayName eq '$escaped'" -ConsistencyLevel eventual -CountVariable c -All) |
               Select-Object -First 1
    if ($group) { return $group }
    if (-not $Create) { return $null }

    if ($WhatIfMode) {
        Write-Change "would create security group $($Definition.displayName)"
        return $null
    }

    $group = New-MgGroup -DisplayName $Definition.displayName `
        -MailNickname (Get-ConfigValue $Definition 'mailNickname' ($Definition.displayName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()) `
        -Description  (Get-ConfigValue $Definition 'description' '') `
        -SecurityEnabled -MailEnabled:$false
    Write-Change "created security group $($Definition.displayName) ($($group.Id))"

    # A brand new group is not immediately resolvable in SharePoint.
    Start-Sleep -Seconds 5
    return $group
}
