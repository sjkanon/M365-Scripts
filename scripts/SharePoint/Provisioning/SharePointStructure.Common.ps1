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

# Declared here, not on first use. Every script in this folder runs under StrictMode,
# where reading a variable that was never assigned is an error - so the usual
# "if (-not $script:X) { $script:X = @{} }" throws instead of initialising it.
$script:StructureConnections = @{}

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
        Property access that returns a default instead of tripping Set-StrictMode when
        the field is simply not there. Used on configuration objects, on Graph
        responses and on whatever shape a PnP cmdlet decides to hand back.

        Hashtables are handled separately on purpose: PSObject.Properties on a
        Hashtable lists the *Hashtable's own* members - Count, Keys, Values - and not
        its keys, so without this branch every lookup against one would quietly answer
        "not there". Quietly wrong is worse than an error, and that is exactly the
        shape Invoke-MgGraphRequest returns unless it is asked for PSObject.
    #>
    param($Object, [string] $Name, $Default = $null)

    if ($null -eq $Object) { return $Default }

    if ($Object -is [System.Collections.IDictionary]) {
        if (-not $Object.Contains($Name)) { return $Default }
        $value = $Object[$Name]
        if ($null -eq $value) { return $Default }
        return $value
    }

    # No [psobject] test guards what follows, on purpose. -is [psobject] is False for
    # an ErrorRecord, for an Exception and for a plain array, so a guard here made this
    # helper answer "not there" for every field of exactly the objects the error paths
    # inspect - which is how a failed Graph call came back as "no detail returned".
    # PSObject.Properties is available on any object at all, so probing it directly is
    # both safer and wider than asking what the object is first.

    # Enumerated one by one rather than as .Properties.Name: member enumeration over
    # an empty property collection is itself an error under StrictMode, so an object
    # with no properties at all would break the very helper meant to survive that.
    $names = @($Object.PSObject.Properties | ForEach-Object { $_.Name })
    if ($names -notcontains $Name) { return $Default }

    $value = $Object.$Name
    if ($null -eq $value) { return $Default }
    return $value
}

function Get-ValueShape {
    <#
        A short, printable description of what a value *is* - its type and the names
        of its fields, never its contents. For the moments where the useful question
        is not what an API said but what shape it said it in, which is exactly when
        something turns out not to have the property everything downstream assumed.

        Contents are left out on purpose: this goes in warnings, and a group's fields
        are safe to print where its data is not.
    #>
    param($Value, [int] $Depth = 0)

    if ($null -eq $Value)       { return '<null>' }
    $type = $Value.GetType().Name
    if ($Value -is [string])    { return "$type '$Value'" }
    if ($Value -is [ValueType]) { return "$type '$Value'" }

    if ($Value -is [System.Collections.IDictionary]) {
        return "$type{$(@($Value.Keys | ForEach-Object { [string] $_ }) -join ', ')}"
    }

    # Depth-limited rather than fully recursive: one level in is enough to tell a
    # collection of groups from a collection of something else, and a cyclic object
    # graph should not be able to hang a warning.
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @($Value)
        if ($items.Count -eq 0) { return "$type(empty)" }
        $inner = if ($Depth -lt 2) { Get-ValueShape $items[0] ($Depth + 1) } else { '...' }
        return "$type($($items.Count)) of $inner"
    }

    $names = @($Value.PSObject.Properties | ForEach-Object { $_.Name })
    if ($names.Count -eq 0) { return $type }
    return "$type{$($names -join ', ')}"
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

    # Every return is comma-wrapped: PowerShell unrolls an array on the way out, so a
    # column with one choice would come back as a bare string and an empty one as
    # $null - and under StrictMode the caller then dies on .Count.
    try {
        $schema = [xml] $Field.SchemaXml
    } catch {
        return ,@()
    }
    $node = $schema.Field.SelectSingleNode('CHOICES')
    if (-not $node) { return ,@() }
    return ,@($node.ChildNodes | ForEach-Object { $_.InnerText })
}

function Test-FieldHiddenInForms {
    <#
        Whether a column is already kept out of the new and edit forms. Read from the
        schema XML, which is on the object Get-PnPField handed back - the typed
        properties would each cost their own round trip.
    #>
    param([Parameter(Mandatory)] $Field)

    $schema = [string] $Field.SchemaXml
    return ($schema -match 'ShowInNewForm="FALSE"') -and ($schema -match 'ShowInEditForm="FALSE"')
}

function Set-FieldHiddenInForms {
    <#
        Take a column out of the new and edit forms while leaving it in the views and
        in the details pane. That is what a script-maintained column needs: Deelstatus
        says what IS, so asking a user to fill it in is asking for a wrong answer.

        Not the same as hiding the field outright - it stays readable, filterable and
        groupable, it just stops being a question.
    #>
    param(
        [Parameter(Mandatory)] $Field,
        [Parameter(Mandatory)] $Connection
    )

    $Field.SetShowInNewForm($false)
    $Field.SetShowInEditForm($false)
    # $true: reach the lists that already use the column, not just new ones.
    $Field.UpdateAndPushChanges($true)
    Invoke-PnPQuery -Connection $Connection
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

    # Comma-wrapped: an empty list of changes is the normal case - everything already
    # matched - and returning it bare hands the caller $null instead.
    return ,$changes
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

function Add-StructureView {
    <#
        Create one view from a configuration entry, if it is not already there.

        Handles the three things the config can ask for beyond a column list:
          groupBy    a GroupBy clause - never on a multi-value column, SharePoint
                     refuses to group on those
          where      raw CAML, so a brand view can say "Butterstone OR Beide" without
                     this script growing a query language of its own
          recursive  Scope = RecursiveAll: show every file in the library regardless
                     of which pillar folder it sits in. This is what makes "everything
                     of one brand" a flat list instead of a folder hunt.

        Returns $true when it created the view.
    #>
    param(
        [Parameter(Mandatory)] [string] $ListTitle,
        [Parameter(Mandatory)] [string] $Title,
        [Parameter(Mandatory)] $Definition,
        [Parameter(Mandatory)] $Connection,
        [switch] $WhatIfMode
    )

    if (Get-PnPView -List $ListTitle -Identity $Title -Connection $Connection -ErrorAction SilentlyContinue) {
        return $false
    }
    if ($WhatIfMode) { return $true }

    $query = ''
    $where = Get-ConfigValue $Definition 'where'
    if ($where) { $query += "<Where>$where</Where>" }

    $groupBy = Get-ConfigValue $Definition 'groupBy'
    if ($groupBy) { $query += "<GroupBy Collapse=`"TRUE`" GroupLimit=`"100`"><FieldRef Name=`"$groupBy`" /></GroupBy>" }

    $splat = @{
        List       = $ListTitle
        Title      = $Title
        Fields     = @(Get-ConfigValue $Definition 'fields' @())
        Connection = $Connection
    }
    if ($query) { $splat['Query'] = $query }
    # Never -SetAsDefault: the default view of a Teams library is what every member of
    # the channel sees the second they open Files.
    $view = Add-PnPView @splat

    if (Get-ConfigValue $Definition 'recursive' $false) {
        # Add-PnPView cannot express the scope, so it is set afterwards.
        $view.Scope = [Microsoft.SharePoint.Client.ViewScope]::RecursiveAll
        $view.Update()
        Invoke-PnPQuery -Connection $Connection
    }
    return $true
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
        [string[]] $Scopes = @('Group.ReadWrite.All', 'Directory.Read.All'),
        [switch] $AuthenticationOnly
    )

    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Module 'Microsoft.Graph.Authentication' is required. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Only the group work needs the Groups cmdlets; the team step talks raw Graph and
    # should not be blocked on a module it never calls.
    if (-not $AuthenticationOnly) {
        if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Groups')) {
            throw "Module 'Microsoft.Graph.Groups' is required for the Entra ID groups. Run: Install-Module Microsoft.Graph.Groups -Scope CurrentUser"
        }
        Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    }

    # Announced before it happens, not after: a browser sign-in can take a minute, a
    # cached one takes none and prints nothing, and without this line a failure
    # anywhere around here is impossible to place in the sequence.
    Write-Step "Signing in to Graph on $Tenant..."

    if ($Thumbprint -and $ClientId) {
        Connect-MgGraph -TenantId $Tenant -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome | Out-Null
    } elseif ($ClientId -and -not $PSBoundParameters.ContainsKey('Scopes')) {
        # An app this set provisioned already carries the Graph scopes, so signing in
        # with it keeps the whole run on one app registration. An app that came from
        # somewhere else may not, hence the fallback to the Graph SDK's own app.
        try {
            Connect-MgGraph -TenantId $Tenant -ClientId $ClientId -NoWelcome -ContextScope Process -ErrorAction Stop | Out-Null
        } catch {
            Write-Warn "App $ClientId could not be used for Graph ($($_.Exception.Message.Trim())) - falling back to the Microsoft Graph PowerShell app."
            Connect-MgGraph -TenantId $Tenant -Scopes $Scopes -NoWelcome -ContextScope Process | Out-Null
        }
    } else {
        Connect-MgGraph -TenantId $Tenant -Scopes $Scopes -NoWelcome -ContextScope Process | Out-Null
    }
    # Get-MgContext comes back $null when the sign-in silently did not take, and
    # reading .Account off that is a property error blamed on the wrong line.
    $context = Get-MgContext
    if (-not $context) { throw "Graph sign-in did not take on $Tenant - no context afterwards." }
    Write-Ok "Graph connected as $((Get-ConfigValue $context 'Account') ?? 'app-only')"
}

# -- Graph ---------------------------------------------------------------------
function Invoke-StructureGraph {
    <#
        One Graph call, shaped so the rest of this set can trust what comes back.

        Two things Invoke-MgGraphRequest gets wrong for us by default. It hands back a
        Hashtable, and under StrictMode - which every script here runs under - asking
        it for a key that is not there is fatal rather than empty; -OutputType PSObject
        plus Get-ConfigValue makes a missing field read as $null instead. And it
        surfaces only the HTTP status, while Graph puts the reason in the body, so
        "BadRequest" is all you get unless the body is unpacked.
    #>
    param(
        [Parameter(Mandatory)] [string] $Url,
        [string] $Method = 'GET',
        $Body
    )

    $uri   = if ($Url -match '^https://') { $Url } else { "https://graph.microsoft.com/$Url" }
    $splat = @{ Uri = $uri; Method = $Method; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($Body) {
        $splat['Body']        = ($Body | ConvertTo-Json -Depth 8)
        $splat['ContentType'] = 'application/json'
    }

    try {
        return Invoke-MgGraphRequest @splat
    } catch {
        # $_ is captured first and never read again. A nested try inside a catch
        # rebinds $_ to its own error, so the handler that was meant to explain the
        # failure reports itself instead - and the real one is never seen. That is
        # worth more care than the happy path: this is the code that runs precisely
        # when you have no other information.
        $err    = $_
        $detail = ''

        $raw = Get-ConfigValue (Get-ConfigValue $err 'ErrorDetails') 'Message'
        if ($raw) {
            $parsed = $null
            try { $parsed = $raw | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }

            $inner   = Get-ConfigValue $parsed 'error'
            $code    = Get-ConfigValue $inner 'code'
            $message = Get-ConfigValue $inner 'message'
            # Not JSON, or not shaped like a Graph error? Then the body itself is the
            # most informative thing available.
            $detail  = if ($message) { "$code`: $message" } else { [string] $raw }
        }

        if (-not $detail) { $detail = [string] (Get-ConfigValue (Get-ConfigValue $err 'Exception') 'Message') }
        if (-not $detail) { $detail = 'no detail returned' }

        throw "$Method $Url -> $detail"
    }
}

function Get-StructureGraphCollection {
    <#
        Every item of a Graph collection, following @odata.nextLink. Always an array,
        even for none or one - the comma keeps PowerShell from unrolling it on the way
        out.
    #>
    param([Parameter(Mandatory)] [string] $Url)

    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Url
    while ($next) {
        $page = Invoke-StructureGraph -Url $next
        foreach ($item in @(Get-ConfigValue $page 'value' @())) { $items.Add($item) }
        $next = Get-ConfigValue $page '@odata.nextLink'
    }
    return ,$items.ToArray()
}

# -- App registration ----------------------------------------------------------
# PnP.PowerShell no longer ships a shared multi-tenant app, so every tenant needs one
# of its own. The client ID is cached in pnp.appid.json at the repo root - the same
# store Find-SiteContent.ps1 and Restore-RecycleBinItems.ps1 use, so an app created
# by one of them is reused here and the other way round.

function Get-StructureAppStorePath {
    $repoRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
    return (Join-Path $repoRoot 'pnp.appid.json')
}

function Get-CachedStructureClientId {
    param([Parameter(Mandatory)] [string] $Tenant)

    $store = Get-StructureAppStorePath
    if (-not (Test-Path $store)) { return $null }
    try {
        return (Get-Content $store -Raw | ConvertFrom-Json).$Tenant
    } catch {
        Write-Warn "Could not read ${store}: $($_.Exception.Message)"
        return $null
    }
}

function Set-CachedStructureClientId {
    param([Parameter(Mandatory)] [string] $Tenant, [Parameter(Mandatory)] [string] $Id)

    $path  = Get-StructureAppStorePath
    $store = @{}
    if (Test-Path $path) {
        try {
            (Get-Content $path -Raw | ConvertFrom-Json).PSObject.Properties |
                ForEach-Object { $store[$_.Name] = $_.Value }
        } catch { }
    }
    $store[$Tenant] = $Id
    $store | ConvertTo-Json | Set-Content -Path $path -Encoding UTF8
    Write-Ok "Client ID cached in $path"
}

function New-StructureApp {
    <#
        Create (or reuse) the public-client app registration this set signs in with,
        and admin-consent the delegated scopes it needs. Returns an object with the
        app id, the directory object id and whether it was created just now - the
        last one is what lets the caller clean up a temporary app afterwards.

        Needs a Graph sign-in as Global or Application Administrator, once.
    #>
    param(
        [Parameter(Mandatory)] [string] $Tenant,
        [Parameter(Mandatory)] [string] $DisplayName
    )

    foreach ($module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Module '$module' is required to create the app registration. Run: Install-Module $module -Scope CurrentUser"
        }
    }
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Write-Step "Signing in as a Global Administrator of $Tenant to register the app..."
    Connect-MgGraph -TenantId $Tenant -NoWelcome -ContextScope Process -Scopes @(
        'Application.ReadWrite.All'
        'DelegatedPermissionGrant.ReadWrite.All'
        'Directory.Read.All'
    ) | Out-Null
    Write-Ok "Signed in as $((Get-MgContext).Account)"

    $escaped = $DisplayName -replace "'", "''"
    $app     = @(Get-MgApplication -Filter "displayName eq '$escaped'" -All) | Select-Object -First 1
    $created = $false

    if ($app) {
        Write-Ok "Reusing app '$DisplayName' ($($app.AppId))"
    } else {
        $app = New-MgApplication -DisplayName $DisplayName `
            -SignInAudience 'AzureADMyOrg' `
            -IsFallbackPublicClient `
            -PublicClient @{ RedirectUris = @('http://localhost') }
        $created = $true
        Write-Change "App registration created: $($app.AppId)"
    }

    # Windows signs in through the Web Account Manager broker by default, and the
    # broker uses a redirect URI built from the app's own id - which cannot be
    # registered before the app exists, so it is patched in here. Without it the
    # sign-in dies on AADTS50011 (redirect URI mismatch) and the app looks broken.
    $wantedUris = @(
        'http://localhost'
        "ms-appx-web://Microsoft.AAD.BrokerPlugin/$($app.AppId)"
        'https://login.microsoftonline.com/common/oauth2/nativeclient'
    )
    $currentUris = @()
    if ($app.PublicClient -and $app.PublicClient.RedirectUris) { $currentUris = @($app.PublicClient.RedirectUris) }
    $missingUris = @($wantedUris | Where-Object { $_ -notin $currentUris })

    if ($missingUris.Count -gt 0) {
        $merged = @(($currentUris + $wantedUris) | Select-Object -Unique)
        Update-MgApplication -ApplicationId $app.Id -PublicClient @{ RedirectUris = $merged } | Out-Null
        Write-Change "Redirect URIs added: $($missingUris.Count) (broker sign-in on Windows)"
        # A redirect URI is checked at sign-in against a replicated copy, so the first
        # attempt right after this still fails if we do not give it a moment.
        Start-Sleep -Seconds 10
    } else {
        Write-Ok 'Redirect URIs already in place.'
    }

    $sp = @(Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -All) | Select-Object -First 1
    if (-not $sp) {
        $sp = New-MgServicePrincipal -AppId $app.AppId
        Write-Change 'Service principal created.'
    }

    # Everything this set does, in delegated form:
    #   AllSites.FullControl   content types, list permissions, breaking inheritance
    #   TermStore.ReadWrite    the Leverancier term set
    #   User.Read.All          resolving people behind sharing links
    #   Group.ReadWrite.All    the pillar security groups, and creating the team
    #   Channel.Create         the channels, the private MGMT one included
    $resources = @(
        @{ AppId = '00000003-0000-0ff1-ce00-000000000000'; Name = 'SharePoint'
           Scopes = @('AllSites.FullControl', 'TermStore.ReadWrite.All', 'User.Read.All') }
        @{ AppId = '00000003-0000-0000-c000-000000000000'; Name = 'Graph'
           Scopes = @('User.Read', 'Group.ReadWrite.All', 'Directory.Read.All',
                      'Channel.Create', 'ChannelSettings.ReadWrite.All', 'Team.Create') }
    )

    foreach ($resource in $resources) {
        $resourceSp = @(Get-MgServicePrincipal -Filter "appId eq '$($resource.AppId)'" -All) | Select-Object -First 1
        if (-not $resourceSp) {
            Write-Warn "Service principal for $($resource.Name) not found - skipped."
            continue
        }

        $valid = @($resource.Scopes | Where-Object { $resourceSp.Oauth2PermissionScopes.Value -contains $_ })
        if ($valid.Count -eq 0) { continue }

        $grant = @(Get-MgOauth2PermissionGrant -All -Filter "clientId eq '$($sp.Id)' and consentType eq 'AllPrincipals'") |
                 Where-Object { $_.ResourceId -eq $resourceSp.Id } | Select-Object -First 1

        $existing = if ($grant -and $grant.Scope) { $grant.Scope -split ' ' } else { @() }
        $missing  = @($valid | Where-Object { $_ -notin $existing })
        if ($missing.Count -eq 0) {
            Write-Ok "$($resource.Name): scopes already consented."
            continue
        }

        $merged = ((@($existing) + $valid) | Where-Object { $_ } | Select-Object -Unique) -join ' '
        if ($grant) {
            Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -Scope $merged | Out-Null
        } else {
            New-MgOauth2PermissionGrant -ClientId $sp.Id -ResourceId $resourceSp.Id `
                -ConsentType 'AllPrincipals' -Scope $merged | Out-Null
        }
        Write-Change "$($resource.Name): consented $($missing -join ', ')"
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    return [PSCustomObject]@{
        AppId    = $app.AppId
        ObjectId = $app.Id
        Created  = $created
        Name     = $DisplayName
    }
}

function Remove-StructureApp {
    <#
        Delete an app registration this run created. Only ever called for an app that
        New-StructureApp reported as Created - an app that was already there predates
        this run and is somebody else's to remove.
    #>
    param(
        [Parameter(Mandatory)] [string] $Tenant,
        [Parameter(Mandatory)] [string] $ObjectId,
        [Parameter(Mandatory)] [string] $DisplayName
    )

    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Connect-MgGraph -TenantId $Tenant -NoWelcome -ContextScope Process -Scopes @('Application.ReadWrite.All') | Out-Null
    Remove-MgApplication -ApplicationId $ObjectId -ErrorAction Stop
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Change "Temporary app registration '$DisplayName' removed."

    # The cache must not keep pointing at an app that no longer exists.
    $path = Get-StructureAppStorePath
    if (Test-Path $path) {
        try {
            $store = @{}
            (Get-Content $path -Raw | ConvertFrom-Json).PSObject.Properties |
                ForEach-Object { $store[$_.Name] = $_.Value }
            if ($store.ContainsKey($Tenant)) {
                $store.Remove($Tenant)
                $store | ConvertTo-Json | Set-Content -Path $path -Encoding UTF8
            }
        } catch {
            Write-Warn "Could not clean the client ID cache: $($_.Exception.Message)"
        }
    }
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
