<#
.SYNOPSIS
    Bulk delete M365 user accounts from a tenant.

.DESCRIPTION
    Removes specified user accounts from Entra ID / Microsoft 365.
    Accepts users via -UserList parameter, a CSV/TXT file, or interactive pipeline.
    Revokes sessions and removes licenses before deletion.

.PARAMETER UserList
    Array of UPNs to delete. e.g. -UserList "user1@domain.com","user2@domain.com"

.PARAMETER CsvPath
    Path to a CSV file with a 'UserPrincipalName' column, or a plain TXT file (one UPN per line).

.PARAMETER DryRun
    Simulate the deletion without making any changes. Default: $true

.PARAMETER SkipLicenseRemoval
    Skip removing licenses before deletion.

.PARAMETER SkipSessionRevoke
    Skip revoking active sessions before deletion.

.PARAMETER OutputPath
    Path for the CSV results report. Default: .\DeletedAccounts_<timestamp>.csv

.PARAMETER TenantId
    Optional: specify the Entra ID tenant ID or domain to connect to.

.EXAMPLE
    # Dry run with inline list
    .\Remove-M365Users.ps1 -UserList "user1@domain.com","user2@domain.com"

.EXAMPLE
    # Actual deletion from a CSV file
    .\Remove-M365Users.ps1 -CsvPath .\users.csv -DryRun:$false

.EXAMPLE
    # Pipeline input
    "user1@domain.com","user2@domain.com" | .\Remove-M365Users.ps1 -DryRun:$false

.EXAMPLE
    # Specific tenant, skip license removal, custom output path
    .\Remove-M365Users.ps1 -CsvPath .\users.txt -TenantId "bravehub.io" -SkipLicenseRemoval -DryRun:$false -OutputPath "C:\Reports\deleted.csv"
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'List')]
param (
    [Parameter(ParameterSetName = 'List', ValueFromPipeline, Position = 0)]
    [string[]]$UserList,

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [switch]$DryRun = $true,

    [switch]$SkipLicenseRemoval,

    [switch]$SkipSessionRevoke,

    [string]$OutputPath = ".\DeletedAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [string]$TenantId
)

begin {
    # ─── MODULE CHECK ─────────────────────────────────────────────────────────
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Error "Microsoft.Graph.Users module not found. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
        exit 1
    }

    # ─── CONNECT ──────────────────────────────────────────────────────────────
    Write-Host "`n🔌 Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $connectParams = @{ Scopes = "User.ReadWrite.All"; NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams

    # ─── LOAD USERS FROM FILE ─────────────────────────────────────────────────
    $allUsers = [System.Collections.Generic.List[string]]::new()

    if ($PSCmdlet.ParameterSetName -eq 'File') {
        $ext = [System.IO.Path]::GetExtension($CsvPath).ToLower()
        if ($ext -eq '.csv') {
            $raw = Import-Csv -Path $CsvPath
            # Accept 'UserPrincipalName' or 'UPN' column headers
            $col = ($raw[0].PSObject.Properties.Name | Where-Object { $_ -match 'userprincipalname|^upn$' } | Select-Object -First 1)
            if (-not $col) {
                Write-Error "CSV must have a 'UserPrincipalName' or 'UPN' column. Found: $($raw[0].PSObject.Properties.Name -join ', ')"
                exit 1
            }
            $raw.$col | Where-Object { $_ -match '@' } | ForEach-Object { $allUsers.Add($_) }
        }
        else {
            # Plain TXT: one UPN per line, skip comments and blanks
            Get-Content -Path $CsvPath |
                Where-Object { $_ -notmatch '^\s*#' -and $_ -match '@' } |
                ForEach-Object { $allUsers.Add($_.Trim()) }
        }
    }

    # ─── INIT RESULTS ─────────────────────────────────────────────────────────
    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $notFound = [System.Collections.Generic.List[string]]::new()

    if ($DryRun) {
        Write-Host "⚠️  DRY RUN MODE — no changes will be made`n" -ForegroundColor Yellow
    }
}

process {
    # Collect pipeline / -UserList input
    if ($PSCmdlet.ParameterSetName -eq 'List' -and $UserList) {
        $UserList | Where-Object { $_ -match '@' } | ForEach-Object { $allUsers.Add($_.Trim()) }
    }
}

end {
    # ─── DEDUPLICATE ──────────────────────────────────────────────────────────
    $upns = $allUsers | Sort-Object -Unique

    if ($upns.Count -eq 0) {
        Write-Warning "No valid UPNs found. Use -UserList or -CsvPath."
        Disconnect-MgGraph | Out-Null
        return
    }

    Write-Host "Processing $($upns.Count) unique account(s)...`n" -ForegroundColor Cyan

    # ─── PROCESS EACH USER ────────────────────────────────────────────────────
    foreach ($upn in $upns) {

        Write-Host "  → $upn" -NoNewline

        try {
            $user = Get-MgUser -UserId $upn `
                -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses,AccountEnabled" `
                -ErrorAction Stop

            if ($DryRun) {
                Write-Host "  [DRY RUN — $($user.DisplayName)]" -ForegroundColor DarkYellow
                $status = "DryRun"
            }
            else {
                # 1. Revoke sessions
                if (-not $SkipSessionRevoke) {
                    Revoke-MgUserSignInSession -UserId $user.Id | Out-Null
                }

                # 2. Remove licenses
                if (-not $SkipLicenseRemoval -and $user.AssignedLicenses.Count -gt 0) {
                    Set-MgUserLicense -UserId $user.Id -BodyParameter @{
                        AddLicenses    = @()
                        RemoveLicenses = $user.AssignedLicenses.SkuId
                    } | Out-Null
                }

                # 3. Delete (soft-delete → Deleted Users, 30d recoverable)
                Remove-MgUser -UserId $user.Id -Confirm:$false
                Write-Host "  ✅ Deleted ($($user.DisplayName))" -ForegroundColor Green
                $status = "Deleted"
            }

            $results.Add([PSCustomObject]@{
                UPN         = $upn
                DisplayName = $user.DisplayName
                Licenses    = ($user.AssignedLicenses.SkuId -join '; ')
                Status      = $status
                Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            })
        }
        catch {
            if ($_.Exception.Message -match "Request_ResourceNotFound|does not exist") {
                Write-Host "  ⚠️  Not found" -ForegroundColor DarkYellow
                $notFound.Add($upn)
            }
            else {
                Write-Host "  ❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red
                $results.Add([PSCustomObject]@{
                    UPN         = $upn
                    DisplayName = "N/A"
                    Licenses    = ""
                    Status      = "Error: $($_.Exception.Message)"
                    Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                })
            }
        }
    }

    # ─── SUMMARY ──────────────────────────────────────────────────────────────
    $deleted = ($results | Where-Object Status -eq 'Deleted').Count
    $errors  = ($results | Where-Object { $_.Status -like 'Error*' }).Count

    Write-Host "`n─────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "─────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "  Total     : $($upns.Count)"
    Write-Host "  Deleted   : $deleted"   -ForegroundColor Green
    Write-Host "  Not found : $($notFound.Count)" -ForegroundColor DarkYellow
    Write-Host "  Errors    : $errors"    -ForegroundColor Red

    if ($notFound.Count -gt 0) {
        Write-Host "`nNot found:" -ForegroundColor DarkYellow
        $notFound | ForEach-Object { Write-Host "  - $_" }
    }

    # ─── EXPORT ───────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n📄 Report saved to: $OutputPath" -ForegroundColor Cyan

    Disconnect-MgGraph | Out-Null
    Write-Host "✔️  Done.`n" -ForegroundColor Green
}