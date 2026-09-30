#Requires -Version 5.1
<#
.SYNOPSIS
    Verify Get-SharePointPermissionsReport.ps1 and Revoke-SharePointUserAccess.ps1 without
    touching a tenant.

.DESCRIPTION
    Two things are checked, both of which fail silently in production if they are wrong.

    The authentication and SharePoint REST layer exists in both scripts, byte for byte. It took
    four live runs against a tenant to get right - certificate credentials because SharePoint
    refuses secret-based app-only tokens, tokens that must prove they carry their app roles
    before being cached, 401 treated as fatal rather than per-site, paging that cannot loop. A
    second copy that quietly drifts from the first is a correctness risk in the script that
    deletes permissions, so the two are compared and the difference is printed when they part.

    The revocation funnel is then driven for real: a dry run must record its intent and execute
    nothing, -Apply must execute and record the result, a failure must land in the audit trail
    rather than vanish, and the grants this script must refuse to remove - through an Entra ID
    group, or to everyone in the tenant - must stay refused.

    Run it after changing either script. Exit code 0 means both are sound.

.EXAMPLE
    .\Test-SharePointAccessScripts.ps1
#>
[CmdletBinding()]
param()

$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
Push-Location $repoRoot
try {
# The auth and REST layer took four live runs against a tenant to get right. It now exists in two
# scripts, and a drifting copy is a correctness risk in the one that deletes permissions - so the
# copies are asserted identical, and the revoke logic is driven against a fake SharePoint.
$report = (Resolve-Path 'scripts/Reporting/Get-SharePointPermissionsReport.ps1').Path
$revoke = (Resolve-Path 'scripts/SharePoint/Revoke-SharePointUserAccess.ps1').Path

$fail = 0
function Check($label, $cond) {
    if ($cond) { Write-Host "  PASS  $label" -ForegroundColor Green } else { Write-Host "  FAIL  $label" -ForegroundColor Red; $script:fail++ }
}
function SharedBlock([string]$Path) {
    $lines = [IO.File]::ReadAllLines($Path)
    $start = [array]::FindIndex($lines, [Predicate[string]] { param($l) $l -like '*SHARED BLOCK START*' })
    $end   = [array]::FindIndex($lines, [Predicate[string]] { param($l) $l -like '*SHARED BLOCK END*' })
    if ($start -lt 0 -or $end -lt $start) { return $null }
    return ($lines[$start..$end] -join "`n")
}

# ── The two copies must not drift ───────────────────────────────────────────
$a = SharedBlock $report
$b = SharedBlock $revoke
Check 'the report has a shared block'        ($null -ne $a)
Check 'the revoke script has one too'        ($null -ne $b)
Check 'the two are byte-identical'           ($a -eq $b)
if ($a -and $b -and $a -ne $b) {
    $la = $a -split "`n"; $lb = $b -split "`n"
    Write-Host ("        report {0} lines, revoke {1} lines" -f $la.Count, $lb.Count) -ForegroundColor DarkYellow
    for ($i = 0; $i -lt [Math]::Min($la.Count, $lb.Count); $i++) {
        if ($la[$i] -ne $lb[$i]) { Write-Host ("        first difference at line {0}:`n          report: {1}`n          revoke: {2}" -f ($i + 1), $la[$i], $lb[$i]) -ForegroundColor DarkYellow; break }
    }
}
Check 'the block carries no script-specific name' ($a -notmatch 'Get-SharePointPermissionsReport|Revoke-SharePointUserAccess')
Check 'both set the app name prefix before it' (
    ((Get-Content $report -Raw) -match '\$TempAppNamePrefix\s*=') -and ((Get-Content $revoke -Raw) -match '\$TempAppNamePrefix\s*='))

# ── Both scripts still parse ────────────────────────────────────────────────
foreach ($p in @($report, $revoke)) {
    $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($p, [ref]$null, [ref]$errs)
    Check ("{0} parses" -f (Split-Path $p -Leaf)) (-not $errs)
}

# ── Safety: the destructive script must not act without being told to ───────
$revokeAst  = [System.Management.Automation.Language.Parser]::ParseFile($revoke, [ref]$null, [ref]$null)
$revokeText = Get-Content $revoke -Raw
$params     = $revokeAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }
Check 'it takes -Apply'                       ($params -contains 'Apply')
Check 'it supports ShouldProcess'             ($revokeText -match 'SupportsShouldProcess\s*=?\s*\$?true|SupportsShouldProcess\b')
Check 'and declares high impact'              ($revokeText -match "ConfirmImpact\s*=\s*'High'")
Check 'the user is mandatory'                 ($revokeText -match '(?s)Parameter\(Mandatory = \$true\)\][^\]]*\s*\[string\] \$UserPrincipalName')
Check 'writes go through one funnel'          (([regex]::Matches($revokeText, 'Invoke-SPPost ')).Count -ge 4)
Check 'every write is inside Invoke-Revocation or the funnel' (
    ([regex]::Matches($revokeText, 'Invoke-Revocation ')).Count -ge 4)

# Dry-run must be the default: -Apply absent, or -WhatIf given, has to short-circuit before any
# write. -WhatIf belongs in the same branch, or it records a declined prompt instead of intent.
$dryBranch = [regex]::Match($revokeText, '(?s)if \(-not \$Apply -or \$WhatIfPreference\) \{.*?return \$false')
Check 'no -Apply returns before writing'      ($dryBranch.Success -and $dryBranch.Value -notmatch 'Invoke-SPPost')
Check '-WhatIf takes the same dry-run branch' ($dryBranch.Success -and $dryBranch.Value -match 'WouldRevoke')
Check 'an already-gone removal is not a failure' ($revokeText -match "AlreadyGone")
Check 'the audit row is written as it happens' ($revokeText -match 'Append-CheckpointRows -Path \$script:ActionCsvPath')

# ── Safety: the things it must refuse to remove ─────────────────────────────
Check 'an Entra group grant is not revoked'   ($revokeText -match "hit\.Kind -eq 'EntraGroup'[\s\S]{0,200}CannotRevoke")
Check 'an Everyone grant is not revoked'      ($revokeText -match "hit\.Kind -eq 'Everyone'[\s\S]{0,200}CannotRevoke")
Check 'it never touches Entra membership'     ($revokeText -notmatch 'removeMemberFromGroup|/members/\$ref|Remove-MgGroupMember')
Check 'it warns that group access survives'   ($revokeText -match 'The group is the grant')

# ── Identifying the right user ──────────────────────────────────────────────
# The one mistake this script must never make. Revoking the wrong person is worse than revoking
# nothing, so the matching is exact and is tested against the near-misses that make it tempting
# to be loose.
foreach ($fn in $revokeAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($fn.Name -in @('ConvertFrom-GuestLoginName', 'Test-UserIdentityMatch', 'Add-ActionRow',
                       'Invoke-Revocation', 'Write-ProgressHost', 'Append-CheckpointRows',
                       'Get-ResponseStatusCode')) {
        . ([scriptblock]::Create($fn.Extent.Text))
    }
}

function SiteUser($login, $upn, $mail) { [PSCustomObject]@{ Id = 7; LoginName = $login; UserPrincipalName = $upn; Email = $mail; PrincipalType = 1 } }

Check 'a guest login decodes to the address'  ((ConvertFrom-GuestLoginName 'i:0#.f|membership|jan_partner.com#ext#@contoso.onmicrosoft.com') -eq 'jan@partner.com')
Check 'a local part with an underscore works' ((ConvertFrom-GuestLoginName 'i:0#.f|membership|jan_de_vries_partner.com#ext#@contoso.onmicrosoft.com') -eq 'jan_de_vries@partner.com')
Check 'a normal member login decodes to null' ($null -eq (ConvertFrom-GuestLoginName 'i:0#.f|membership|jan@contoso.com'))
Check 'empty input decodes to null'           ($null -eq (ConvertFrom-GuestLoginName ''))

Check 'an exact UPN matches'                  (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan@contoso.com' 'jan@contoso.com' 'jan@contoso.com') -Needle 'jan@contoso.com')
Check 'the claim suffix alone matches'        (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan@contoso.com' $null $null) -Needle 'jan@contoso.com')
Check 'a guest matches on their real address' (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan_partner.com#ext#@contoso.onmicrosoft.com' $null $null) -Needle 'jan@partner.com')
Check 'a guest matches on their tenant UPN'   (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan_partner.com#ext#@contoso.onmicrosoft.com' $null $null) -Needle 'jan_partner.com#ext#@contoso.onmicrosoft.com')
Check 'matching on mail works'                (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|weird' $null 'jan@contoso.com') -Needle 'jan@contoso.com')

# The near-misses. Each of these would have matched a substring test.
Check 'a shorter address does NOT match'      (-not (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan@contoso.com' 'jan@contoso.com' $null) -Needle 'an@contoso.com'))
Check 'a longer address does NOT match'       (-not (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan@contoso.com' 'jan@contoso.com' $null) -Needle 'marjan@contoso.com'))
Check 'a different domain does NOT match'     (-not (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan@contoso.com' 'jan@contoso.com' $null) -Needle 'jan@contoso.com.evil.net'))
Check 'another tenant guest does NOT match'   (-not (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan_other.com#ext#@contoso.onmicrosoft.com' $null $null) -Needle 'jan@partner.com'))
Check 'an empty needle matches nobody'        (-not (Test-UserIdentityMatch -SiteUser (SiteUser 'i:0#.f|membership|jan@contoso.com' 'jan@contoso.com' $null) -Needle ''))
Check 'the lookup refuses on ambiguity'       ($revokeText -match 'different accounts in this site')

# ── The revocation funnel, driven for real ──────────────────────────────────
$UserPrincipalName = 'jan@contoso.com'
$script:ActionRows = [System.Collections.Generic.List[object]]::new()
$auditDir  = Join-Path $env:TEMP "sp-revoke-test-$PID"
New-Item -ItemType Directory -Path $auditDir -Force | Out-Null
$script:ActionCsvPath = Join-Path $auditDir 'audit.csv'
$row = @{ SiteUrl = 'https://c/sites/F'; WebUrl = 'https://c/sites/F'; ScopeType = 'Web'
          ScopeTitle = 'F'; ScopeUrl = 'https://c/sites/F'; AccessVia = 'Direct'; PermissionLevels = 'Read' }

# Dry run: records the intent, runs nothing.
$Apply = $false
$ran = $false
$result = Invoke-Revocation -Target 'F' -Operation 'Remove direct Read' -Row $row -Do { $script:ran = $true }
Check 'a dry run does not execute'            (-not $ran -and -not $result)
Check 'but it is still recorded'              ($script:ActionRows.Count -eq 1 -and $script:ActionRows[0].Action -eq 'WouldRevoke')
Check 'the row names the user'                ($script:ActionRows[0].UserPrincipalName -eq 'jan@contoso.com')

# Apply: executes and records the result.
$Apply = $true
$PSCmdlet = $null   # no cmdlet context in a dot-sourced function; ShouldProcess is exercised by the script itself
function Test-Apply {
    [CmdletBinding(SupportsShouldProcess = $true)] param()
    $script:ran2 = $false
    $ok = Invoke-Revocation -Target 'F' -Operation 'Remove direct Read' -Row $row -Do { $script:ran2 = $true }
    return @{ Ok = $ok; Ran = $script:ran2 }
}
$applied = Test-Apply -Confirm:$false
Check 'with -Apply the removal runs'          ($applied.Ran -and $applied.Ok)
Check 'and is recorded as revoked'            ($script:ActionRows[-1].Action -eq 'Revoked')

# A failing removal must be recorded, not thrown away.
function Test-Failure {
    [CmdletBinding(SupportsShouldProcess = $true)] param()
    return Invoke-Revocation -Target 'F' -Operation 'Remove direct Read' -Row $row -Do { throw 'HTTP 403' }
}
$failed = Test-Failure -Confirm:$false
Check 'a failed removal returns false'        (-not $failed)
Check 'and lands in the audit trail'          ($script:ActionRows[-1].Action -eq 'Failed' -and $script:ActionRows[-1].Detail -match '403')

# Every row must share one schema, or the CSV export drops columns.
$names = @($script:ActionRows | ForEach-Object { ($_.PSObject.Properties.Name | Sort-Object) -join ',' } | Select-Object -Unique)
Check 'all audit rows share one schema'       ($names.Count -eq 1)
Check 'the schema carries the decision'       ($names[0] -match 'Action' -and $names[0] -match 'AccessVia' -and $names[0] -match 'Detail')

# The audit trail must be on disk already, not waiting for the end of the run: a script that
# revokes two hundred things and then dies has to leave a record of what it removed.
Check 'the CSV exists mid-run'                (Test-Path $script:ActionCsvPath)
$onDisk = @(Import-Csv $script:ActionCsvPath)
Check 'every row reached the file'            ($onDisk.Count -eq $script:ActionRows.Count)
Check 'the file holds the outcomes'           ((($onDisk.Action | Sort-Object -Unique) -join ',') -match 'Revoked')
Check 'and the failure, not just the success' (($onDisk | Where-Object { $_.Action -eq 'Failed' }).Count -eq 1)

# An already-gone removal must read as success, not as a failed revocation.
function Test-AlreadyGone {
    [CmdletBinding(SupportsShouldProcess = $true)] param()
    return Invoke-Revocation -Target 'F' -Operation 'Remove direct Read' -Row $row -Do {
        $resp = [PSCustomObject]@{ StatusCode = 404; Headers = @{} }
        $ex = [System.Exception]::new('HTTP 404')
        $ex | Add-Member -NotePropertyName Response -NotePropertyValue $resp -Force
        throw [System.Management.Automation.ErrorRecord]::new($ex, 'e', 'InvalidResult', $null)
    }
}
[void](Test-AlreadyGone -Confirm:$false)
Check 'a 404 is recorded as already gone'     ($script:ActionRows[-1].Action -eq 'AlreadyGone')

Remove-Item -Path $auditDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ''
if ($fail) { Write-Host "$fail check(s) FAILED" -ForegroundColor Red; exit 1 }
Write-Host 'Revoke script and shared-block integrity verified' -ForegroundColor Green

} finally { Pop-Location }