#Requires -Version 7.0
<#
.SYNOPSIS
    Build the whole SharePoint structure in one run: app registration, metadata model,
    libraries, permissions, views and a verification pass. Supports -WhatIf.

.DESCRIPTION
    The one-command version of this folder: from an empty tenant to a working
    structure, asking what everything should be called on the way.

        Ask       with no configuration for this tenant, New-StructureConfig.ps1 runs
                  first and asks the questions - client, tenant, team, brands,
                  pillars, groups, labels. Enter takes the suggestion. -AllNames asks
                  for the derived names too. Nobody has to open a JSON file.
        0. App    create or reuse the Entra app registration and admin-consent the
                  delegated scopes. Cached in pnp.appid.json, shared with the other
                  PnP scripts in this repo.
        1. Team   New-SharePointTeam.ps1 - the Microsoft 365 team, its channels
                  including the private ones, and the site URLs read back from Graph
                  and written into the configuration. -SkipTeam when it exists.
        2. Metadata  New-SharePointMetadata.ps1 - term set, site columns, content
                  types, on every site in the configuration.
        3. Libraries  Set-SharePointLibraries.ps1 -EnsureGroups - Entra ID security
                  groups, libraries, channel folders, content type binding, default
                  metadata, the pillar views, the cross-cutting brand views and the
                  permissions.
        4. Verify Test-SharePointStructure.ps1 - read-only, reports anything that did
                  not land.
        5. Audit  optional (-RunAudit): the first deelstatus pass, so the column has a
                  value on every existing document from day one.

    Each step is a separate script, so anything that goes wrong can be rerun on its
    own without starting over. This one just runs them in order and stops at the first
    failure rather than building on top of a broken step.

    Which configuration it uses
    ---------------------------
    Without -ConfigPath it looks beside itself for a *.config.json that no longer
    contains the CHANGEME placeholders. Exactly one: that one. More than one: it says
    so and asks for -ConfigPath. None: it runs the wizard and uses what that writes.

    Dry run first
    -------------
    -WhatIf walks the whole thing and reports what each step would do, without an app
    registration being created and without touching the tenant. That is the run to do
    first, and to show the customer.

    The temporary app registration
    ------------------------------
    PnP.PowerShell no longer ships a shared multi-tenant app, so the tenant needs one.
    Three ways to handle that:

      (default)      create it, cache the client ID in pnp.appid.json, leave it in
                     place - the scheduled audit and later drift checks reuse it
      -TemporaryApp  create it, use it, delete it again at the end. Nothing is left
                     behind in the customer's tenant, but every later run has to
                     register one again. Pick this for a one-off build on a tenant you
                     do not manage day to day
      -ClientId      use an app you already have; nothing is created or deleted

    -TemporaryApp only ever deletes an app this run created. An app that was already
    there predates the run and is somebody else's to remove.

    A word on what this changes
    ---------------------------
    Step 2 changes permissions on a live team: pillars get their own security groups
    and the folders stop inheriting from the team. Members who are not in the right
    group lose access to that channel's files. Run it with -WhatIf first, and have the
    group membership ready before you run it for real.

    Exit codes
    ----------
        0  built and verified
        1  a step failed
        2  built, but the verification found differences

.PARAMETER ConfigPath
    Path to the structure configuration JSON.
    Default: petsolutions.config.json next to this script.

.PARAMETER AllNames
    Passed to the wizard when it runs: ask for every name, including the channel,
    folder, content type, group and view names that are otherwise derived.

.PARAMETER TemporaryApp
    Delete the app registration again when the run finishes, if this run created it.

.PARAMETER AppName
    Display name of the app registration to create or reuse.
    Default: "M365-Scripts SharePoint Structure".

.PARAMETER ClientId
    Use an existing app registration instead of creating one. Skips step 0 entirely.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.PARAMETER RunAudit
    Also run the first deelstatus audit at the end, so every existing document gets a
    value in the Deelstatus column immediately.

.PARAMETER SkipChannelFolderPermissions
    Passed to step 2: leave Teams channel folders inheriting instead of giving them
    unique permissions. See the note on standard channels in the readme.

.PARAMETER RemoveStockContentType
    Passed to step 2: remove the built-in Document content type from the libraries.

.PARAMETER SkipVerify
    Do not run the verification pass.

.PARAMETER Disconnect
    Sign out of PnP and Graph when finished.

.EXAMPLE
    # From an empty tenant: it asks what everything should be called, then builds
    .\Install-SharePointStructure.ps1

.EXAMPLE
    # Same, but decide every single name yourself
    .\Install-SharePointStructure.ps1 -AllNames

.EXAMPLE
    # Dry run against a configuration that already exists
    .\Install-SharePointStructure.ps1 -WhatIf

.EXAMPLE
    # One-off build on a tenant you do not manage: leave nothing behind
    .\Install-SharePointStructure.ps1 -TemporaryApp -RunAudit

.EXAMPLE
    # Conservative: everything except the unsupported channel-folder permissions
    .\Install-SharePointStructure.ps1 -SkipChannelFolderPermissions

.EXAMPLE
    # Use an app registration you already have
    .\Install-SharePointStructure.ps1 -ClientId 11111111-2222-3333-4444-555555555555

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x, Microsoft.Graph.Applications, Microsoft.Graph.Groups
    Rights  : Global or Application Administrator for step 0; site owner or SharePoint
              administrator for the rest; term store administrator for the term set.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,

    [switch] $TemporaryApp,
    [string] $AppName = 'M365-Scripts SharePoint Structure',
    [string] $ClientId,
    [string] $Tenant,

    [switch] $AllNames,
    [switch] $SkipTeam,
    [string] $Owner,
    [switch] $RunAudit,
    [switch] $SkipChannelFolderPermissions,
    [switch] $RemoveStockContentType,
    [switch] $SkipVerify,
    [switch] $Disconnect
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

$simulate = [bool] $WhatIfPreference

# No configuration named, or the one lying here is still the shipped example? Then ask
# the questions instead of telling the operator to go and edit JSON.
if (-not $ConfigPath) {
    $candidates = @(Get-ChildItem -Path $PSScriptRoot -Filter '*.config.json' -ErrorAction SilentlyContinue |
                    Where-Object { (Get-Content $_.FullName -Raw) -notmatch 'CHANGEME' })

    if ($candidates.Count -eq 1) {
        $ConfigPath = $candidates[0].FullName
        Write-Host "  Configuratie: $($candidates[0].Name)" -ForegroundColor DarkGray
    } elseif ($candidates.Count -gt 1) {
        throw ("More than one configuration here - pass -ConfigPath. Found: {0}" -f (($candidates | ForEach-Object { $_.Name }) -join ', '))
    } else {
        Write-Host ''
        Write-Host '  Nog geen configuratie voor deze klant. Ik stel eerst een paar vragen.' -ForegroundColor Cyan

        # The wizard puts the path it wrote on the pipeline and nothing else, so take
        # it from there. Its exit code is no use: a script that simply ends sets none,
        # and the stale value from whatever ran before would be read instead.
        $wizardArgs = @{}
        if ($AllNames) { $wizardArgs['All'] = $true }
        $ConfigPath = & (Join-Path $PSScriptRoot 'New-StructureConfig.ps1') @wizardArgs |
                      Select-Object -Last 1

        if (-not $ConfigPath -or -not (Test-Path $ConfigPath)) {
            throw 'No configuration was created - nothing to build.'
        }
    }
}

# Unless the team already exists, the site URLs are what step 1 discovers and writes
# back - so placeholders there are expected rather than an error.
$config = Import-StructureConfig -Path $ConfigPath -AllowPlaceholders:(-not $SkipTeam)
if (-not $Tenant) { $Tenant = $config.tenant }
if ($Tenant -match 'CHANGEME') { throw "Fill in the tenant in $ConfigPath first, or delete that file and rerun to be asked instead." }

Assert-PnPModule

$app        = $null
$exitCode   = 0
$stepResult = [System.Collections.Generic.List[object]]::new()
$timer      = [Diagnostics.Stopwatch]::StartNew()

function Invoke-Step {
    <#
        Run one of the scripts in this folder and record how it went. A step that
        throws or comes back non-zero stops the run: the next step would be building
        on something that is not there.
    #>
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $Script,
        [Parameter(Mandatory)] [hashtable] $Arguments,
        [int[]] $AcceptExitCode = @(0)
    )

    Write-Head "$Name"
    $path = Join-Path $PSScriptRoot $Script
    if (-not (Test-Path $path)) { throw "Step script not found: $path" }

    $global:LASTEXITCODE = 0
    try {
        & $path @Arguments
    } catch {
        $stepResult.Add([PSCustomObject]@{ Step = $Name; Result = 'FAILED'; Detail = $_.Exception.Message })
        throw "$Name failed: $($_.Exception.Message)"
    }

    $code = if ($null -eq $LASTEXITCODE) { 0 } else { $LASTEXITCODE }
    if ($code -notin $AcceptExitCode) {
        $stepResult.Add([PSCustomObject]@{ Step = $Name; Result = 'FAILED'; Detail = "exit code $code" })
        throw "$Name failed with exit code $code."
    }

    $stepResult.Add([PSCustomObject]@{ Step = $Name; Result = 'ok'; Detail = "exit code $code" })
    return $code
}

try {
    Write-Host ''
    Write-Host '  ┌──────────────────────────────────────────────────────────────┐' -ForegroundColor Cyan
    Write-Host '  │  SharePoint structure - build everything in one run           │' -ForegroundColor Cyan
    Write-Host '  └──────────────────────────────────────────────────────────────┘' -ForegroundColor Cyan
    Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
    Write-Host "  Tenant : $Tenant" -ForegroundColor Cyan
    foreach ($property in $config.sites.PSObject.Properties) {
        Write-Host "  Site   : $($property.Name) -> $($property.Value)" -ForegroundColor Cyan
    }
    Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be changed anywhere' } else { 'APPLY' })" `
        -ForegroundColor $(if ($simulate) { 'Yellow' } else { 'Cyan' })

    if (-not $simulate) {
        Write-Host ''
        Write-Warn 'Step 2 changes permissions on a live team: each pillar folder stops inheriting'
        Write-Warn 'and only its own security group keeps access. Make sure the groups have members.'
    }

    # -- 0. App registration ---------------------------------------------------
    Write-Head '0. App registration'
    if ($ClientId) {
        Write-Skip "Using the app registration you passed in ($ClientId)"
    } elseif ($simulate) {
        $cached = Get-CachedStructureClientId -Tenant $Tenant
        if ($cached) {
            $ClientId = $cached
            Write-Ok "Would reuse the cached app registration $cached"
        } else {
            Write-Skip "Would create the app registration '$AppName' and admin-consent its delegated scopes"
            Write-Skip 'The steps below cannot connect without one, so they are reported from the configuration only'
        }
    } else {
        $ClientId = Get-CachedStructureClientId -Tenant $Tenant
        if ($ClientId) {
            Write-Ok "Reusing the cached app registration $ClientId"
            if ($TemporaryApp) {
                Write-Warn 'This app was already cached, so -TemporaryApp will not remove it - it is not ours to delete.'
            }
        } else {
            $app = New-StructureApp -Tenant $Tenant -DisplayName $AppName
            $ClientId = $app.AppId
            if (-not $TemporaryApp) { Set-CachedStructureClientId -Tenant $Tenant -Id $ClientId }
            else { Write-Skip 'Not cached - -TemporaryApp removes this registration at the end of the run' }

            # A brand new registration is not replicated to every region yet, and the
            # first sign-in against it otherwise fails with "application not found".
            Write-Step 'Waiting 20s for the app registration to propagate...'
            Start-Sleep -Seconds 20
        }
    }

    $common = @{ Interactive = $true; ConfigPath = $ConfigPath; Tenant = $Tenant }
    if ($ClientId) { $common['ClientId'] = $ClientId }
    if ($simulate) { $common['WhatIf'] = $true }

    if ($simulate -and -not $ClientId) {
        Write-Head 'Summary'
        Write-Host '    No app registration exists for this tenant yet, so a -WhatIf run cannot sign in' -ForegroundColor Yellow
        Write-Host '    and report per step. The configuration itself validated cleanly.' -ForegroundColor Yellow
        Write-Host '    Run once without -WhatIf, or pass -ClientId of an existing app to dry-run properly.' -ForegroundColor Yellow
        Write-Host ''
        exit 0
    }

    # -- 1. Team and channels --------------------------------------------------
    # First, because everything below needs a site to connect to - and the private
    # channel's site URL does not exist until the channel does.
    if ($SkipTeam) {
        Write-Head '1. Team and channels'
        Write-Skip 'Skipped (-SkipTeam) - the team and its channels are assumed to exist'
    } else {
        $teamArgs = @{ ConfigPath = $ConfigPath; Tenant = $Tenant; Interactive = $true }
        if ($ClientId) { $teamArgs['ClientId'] = $ClientId }
        if ($Owner)    { $teamArgs['Owner'] = $Owner }
        if ($simulate) { $teamArgs['WhatIf'] = $true }

        Invoke-Step -Name '1. Team and channels (incl. the private MGMT channel)' `
            -Script 'New-SharePointTeam.ps1' -Arguments $teamArgs | Out-Null

        # That step writes the discovered site URLs back into the file, so everything
        # below has to read the config again rather than the copy loaded at startup.
        if (-not $simulate) {
            $config = Import-StructureConfig -Path $ConfigPath
            foreach ($property in $config.sites.PSObject.Properties) {
                Write-Ok "site '$($property.Name)' -> $($property.Value)"
            }
        }
    }

    # -- 2. Metadata -----------------------------------------------------------
    Invoke-Step -Name '2. Metadata model (term set, columns, content types)' `
        -Script 'New-SharePointMetadata.ps1' -Arguments $common | Out-Null

    # -- 2. Libraries, groups, views and permissions ----------------------------
    $librariesArgs = $common.Clone()
    $librariesArgs['EnsureGroups'] = $true
    if ($SkipChannelFolderPermissions) { $librariesArgs['SkipChannelFolderPermissions'] = $true }
    if ($RemoveStockContentType)       { $librariesArgs['RemoveStockContentType'] = $true }

    Invoke-Step -Name '3. Libraries, groups, views and permissions' `
        -Script 'Set-SharePointLibraries.ps1' -Arguments $librariesArgs | Out-Null

    # -- 3. Verify -------------------------------------------------------------
    # Exit code 2 is drift, not a failure: on a -WhatIf run everything is "missing"
    # because nothing was built, and that is the expected answer.
    if ($SkipVerify) {
        Write-Head '4. Verification'
        Write-Skip 'Skipped (-SkipVerify)'
    } else {
        $verifyArgs = @{ ConfigPath = $ConfigPath; Tenant = $Tenant; Interactive = $true; IncludeGroups = $true }
        if ($ClientId) { $verifyArgs['ClientId'] = $ClientId }
        $verify = Invoke-Step -Name '4. Verification (read only)' `
            -Script 'Test-SharePointStructure.ps1' -Arguments $verifyArgs -AcceptExitCode @(0, 2)
        if ($verify -eq 2 -and -not $simulate) { $exitCode = 2 }
    }

    # -- 4. First audit --------------------------------------------------------
    if ($RunAudit) {
        $auditArgs = $common.Clone()
        # Exit 2 means "something is shared wider than its tag allows" - worth knowing
        # on day one, but not a reason to call the build failed.
        Invoke-Step -Name '5. First deelstatus audit' `
            -Script 'Update-SharePointShareStatus.ps1' -Arguments $auditArgs -AcceptExitCode @(0, 2) | Out-Null
    }

    # -- Summary ---------------------------------------------------------------
    Write-Head 'Build summary'
    $stepResult | Format-Table -AutoSize | Out-Host
    Write-Host "    Duration: $([int]$timer.Elapsed.TotalMinutes)m $($timer.Elapsed.Seconds)s" -ForegroundColor Cyan

    if ($simulate) {
        Write-Host '    Dry run only - rerun without -WhatIf to build it for real.' -ForegroundColor Yellow
    } elseif ($exitCode -eq 2) {
        Write-Warn 'Built, but the verification reported differences - see the drift report above.'
    } else {
        Write-Ok 'Structure built and verified.'
        Write-Host ''
        Write-Host '    Next:' -ForegroundColor Cyan
        Write-Host '      1. Put people in the SG-PETSOL-* security groups (Entra ID portal)' -ForegroundColor Gray
        Write-Host '      2. Hand the users Petsolutions-SharePoint-Handleiding.md' -ForegroundColor Gray
        Write-Host '      3. Schedule Update-SharePointShareStatus.ps1 nightly and' -ForegroundColor Gray
        Write-Host '         Test-SharePointStructure.ps1 weekly' -ForegroundColor Gray
    }
    Write-Host ''
} catch {
    Write-Host ''
    Write-Bad "Build stopped: $($_.Exception.Message)"
    if ($stepResult.Count -gt 0) {
        Write-Host ''
        $stepResult | Format-Table -AutoSize | Out-Host
    }
    Write-Host '    Each step is its own script - fix the cause and rerun just that one.' -ForegroundColor DarkGray
    Write-Host ''
    $exitCode = 1
} finally {
    if ($app -and $app.Created -and $TemporaryApp -and -not $simulate) {
        Write-Head 'Cleanup'
        try {
            Remove-StructureApp -Tenant $Tenant -ObjectId $app.ObjectId -DisplayName $app.Name
        } catch {
            Write-Warn "Could not remove the temporary app registration '$($app.Name)': $($_.Exception.Message)"
            Write-Warn "Remove it by hand in Entra ID > App registrations (app id $($app.AppId))."
        }
    }
    if ($Disconnect) {
        Disconnect-Structure
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}

exit $exitCode
