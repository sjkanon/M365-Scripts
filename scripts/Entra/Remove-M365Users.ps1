#Requires -Version 7.0
<#
.SYNOPSIS
    Bulk delete M365 user accounts from a tenant.

.DESCRIPTION
    Removes specified user accounts from Entra ID / Microsoft 365.
    Accepts users via -UserList parameter, a CSV/TXT file, or interactive pipeline.
    Revokes sessions and removes licenses before deletion.
    Defaults to dry-run mode — pass -Apply to perform actual deletions.

    Signs in to Microsoft Graph through scripts\Startup\Connect-M365.ps1: delegated as
    the admin by default (device code / GDAP customer per load.config.ps1), app-only with
    -ClientId and -CertificateThumbprint or -AppOnly. An existing Graph session that fits
    is reused and left connected; only a session this script opened is disconnected.
    Delegated scope: User.ReadWrite.All.

.PARAMETER UserList
    Array of UPNs to delete. e.g. -UserList "user1@domain.com","user2@domain.com"

.PARAMETER CsvPath
    Path to a CSV file with a 'UserPrincipalName' or 'UPN' column, or a plain TXT
    file with one UPN per line.

.PARAMETER Apply
    Actually perform the deletions. Without this switch, the script runs in dry-run mode.

.PARAMETER SkipLicenseRemoval
    Skip removing licenses before deletion.

.PARAMETER SkipSessionRevoke
    Skip revoking active sessions before deletion.

.PARAMETER OutputPath
    Path for the CSV results report. Default: .\DeletedAccounts_<timestamp>.csv

.PARAMETER TenantId
    Optional: Entra ID tenant ID or domain to connect to. Defaults to the GDAP customer
    when authMode is GDAP, else the tenant you sign in to.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Dry run with inline list (default — no changes made)
    .\Remove-M365Users.ps1 -UserList "user1@domain.com","user2@domain.com"

.EXAMPLE
    # Actual deletion from a CSV file
    .\Remove-M365Users.ps1 -CsvPath .\users.csv -Apply

.EXAMPLE
    # Pipeline input
    "user1@domain.com","user2@domain.com" | .\Remove-M365Users.ps1 -Apply

.EXAMPLE
    # Specific tenant, skip license removal, custom output path
    .\Remove-M365Users.ps1 -CsvPath .\users.txt -TenantId "contoso.com" -SkipLicenseRemoval -Apply -OutputPath "C:\Reports\deleted.csv"
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'List')]
param (
    [Parameter(ParameterSetName = 'List', ValueFromPipeline, Position = 0)]
    [string[]] $UserList,

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [switch] $Apply,

    [switch] $SkipLicenseRemoval,

    [switch] $SkipSessionRevoke,

    [string] $OutputPath,

    [string] $TenantId,

    [string] $ClientId,

    [string] $CertificateThumbprint,

    [switch] $AppOnly
)

begin {
    . (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

    # ── Output folder ─────────────────────────────────────────────────────────
    $outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
    if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "DeletedAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

    # ── Module check ──────────────────────────────────────────────────────────
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Error "Microsoft.Graph.Users module not found. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
        exit 1
    }

    # ── Connect (reuses a fitting session; Get-MgContext never throws, so the old
    #    try/catch around it never connected when there was no session) ───────────
    $script:Graph = Connect-M365Graph -Scopes 'User.ReadWrite.All' -TenantId $TenantId `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

    # ── Load users from file ──────────────────────────────────────────────────
    $allUsers = [System.Collections.Generic.List[string]]::new()

    if ($PSCmdlet.ParameterSetName -eq 'File') {
        $ext = [System.IO.Path]::GetExtension($CsvPath).ToLower()
        if ($ext -eq '.csv') {
            $raw = Import-Csv -Path $CsvPath
            $col = $raw[0].PSObject.Properties.Name |
                   Where-Object { $_ -match 'userprincipalname|^upn$' } |
                   Select-Object -First 1
            if (-not $col) {
                Write-Error "CSV must have a 'UserPrincipalName' or 'UPN' column. Found: $($raw[0].PSObject.Properties.Name -join ', ')"
                exit 1
            }
            $raw.$col | Where-Object { $_ -match '@' } | ForEach-Object { $allUsers.Add($_) }
        } else {
            Get-Content -Path $CsvPath |
                Where-Object { $_ -notmatch '^\s*#' -and $_ -match '@' } |
                ForEach-Object { $allUsers.Add($_.Trim()) }
        }
    }

    # ── Init results ──────────────────────────────────────────────────────────
    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $notFound = [System.Collections.Generic.List[string]]::new()

    if (-not $Apply) {
        Write-Host ""
        Write-Host "  ================================================" -ForegroundColor Yellow
        Write-Host "   DRY RUN MODE — no changes will be made" -ForegroundColor Yellow
        Write-Host "   Add -Apply to perform actual deletions." -ForegroundColor Yellow
        Write-Host "  ================================================" -ForegroundColor Yellow
    }
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'List' -and $UserList) {
        $UserList | Where-Object { $_ -match '@' } | ForEach-Object { $allUsers.Add($_.Trim()) }
    }
}

end {
    # ── Deduplicate ───────────────────────────────────────────────────────────
    $uniqueUpns = $allUsers | Sort-Object -Unique

    if ($uniqueUpns.Count -eq 0) {
        Write-Warning "No valid UPNs found. Use -UserList or -CsvPath."
        Disconnect-M365Graph $script:Graph
        return
    }

    Write-Host ""
    Write-Host "  Processing $($uniqueUpns.Count) account(s)..." -ForegroundColor Cyan
    Write-Host ""

    # ── Process each user ─────────────────────────────────────────────────────
    foreach ($userUpn in $uniqueUpns) {

        Write-Host "  -> $userUpn" -NoNewline

        try {
            $user = Get-MgUser -UserId $userUpn `
                -Property 'Id,DisplayName,UserPrincipalName,AssignedLicenses,AccountEnabled' `
                -ErrorAction Stop

            if (-not $Apply) {
                Write-Host "  [DRY RUN — $($user.DisplayName)]" -ForegroundColor DarkYellow
                $status = 'DryRun'
            } else {
                if (-not $SkipSessionRevoke) {
                    Revoke-MgUserSignInSession -UserId $user.Id | Out-Null
                }

                if (-not $SkipLicenseRemoval -and $user.AssignedLicenses.Count -gt 0) {
                    Set-MgUserLicense -UserId $user.Id -BodyParameter @{
                        AddLicenses    = @()
                        RemoveLicenses = $user.AssignedLicenses.SkuId
                    } | Out-Null
                }

                # Soft-delete — account lands in Deleted Users, recoverable for 30 days
                Remove-MgUser -UserId $user.Id -Confirm:$false
                Write-Host "  Deleted ($($user.DisplayName))" -ForegroundColor Green
                $status = 'Deleted'
            }

            $results.Add([PSCustomObject]@{
                UPN         = $userUpn
                DisplayName = $user.DisplayName
                Licenses    = ($user.AssignedLicenses.SkuId -join '; ')
                Status      = $status
                Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            })
        } catch {
            if ($_.Exception.Message -match 'Request_ResourceNotFound|does not exist') {
                Write-Host "  Not found" -ForegroundColor DarkYellow
                $notFound.Add($userUpn)
            } else {
                Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                $results.Add([PSCustomObject]@{
                    UPN         = $userUpn
                    DisplayName = 'N/A'
                    Licenses    = ''
                    Status      = "Error: $($_.Exception.Message)"
                    Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                })
            }
        }
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    $deleted = ($results | Where-Object Status -eq 'Deleted').Count
    $errors  = ($results | Where-Object { $_.Status -like 'Error*' }).Count

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Summary" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "  Total     : $($uniqueUpns.Count)"
    Write-Host "  Deleted   : $deleted"              -ForegroundColor Green
    Write-Host "  Not found : $($notFound.Count)"    -ForegroundColor DarkYellow
    Write-Host "  Errors    : $errors"               -ForegroundColor Red

    if ($notFound.Count -gt 0) {
        Write-Host ""
        Write-Host "  Not found:" -ForegroundColor DarkYellow
        $notFound | ForEach-Object { Write-Host "    - $_" }
    }

    # ── Export ────────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Report saved to: $OutputPath" -ForegroundColor Cyan
    Write-Host ""

    Disconnect-M365Graph $script:Graph
}
