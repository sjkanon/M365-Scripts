#Requires -Version 7
<#
.SYNOPSIS
    Keep the generated documentation in step with every change: the script index,
    the readme headers, and the link check - and flag readmes whose translations
    fell behind.

.DESCRIPTION
    One script, three callers:

      -Mode PostEdit   Claude Code PostToolUse hook (Edit/Write), run in the background
                       (asyncRewake). When the touched file is a .ps1, .psd1 or .md in this
                       repo it regenerates INDEX.md and the readme headers, runs the link
                       check, and checks that every module a script loads is listed in
                       scripts/Startup/RequiredModules.psd1.
                       Silent when all is well; on a problem it writes it to stderr and
                       exits 2, which wakes Claude to fix it.
      -Mode Stop       Claude Code Stop hook. Blocks the stop once when a readme.md was
                       changed but its readme.nl.md / readme.fr.md were not, or a script
                       changed while no readme in its folder did.
      -Mode PreCommit  git pre-commit hook for changes made by hand. Regenerates, re-stages
                       what it regenerated, fails on a broken link, and warns about readmes
                       whose translations were not updated (a hook cannot translate).
#>
[CmdletBinding()]
param(
    [ValidateSet('PostEdit', 'Stop', 'PreCommit')] [string] $Mode = 'PostEdit'
)

$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$startup = Join-Path $root 'scripts/Startup'

function Invoke-Tool([string] $Name, [string[]] $Arguments = @()) {
    $output = & pwsh -NoProfile -File (Join-Path $startup $Name) @Arguments 2>&1 | Out-String
    [pscustomobject]@{ Ok = ($LASTEXITCODE -eq 0); Output = $output.Trim() }
}

function Get-UntranslatedReadme {
    # English readmes changed (staged or not) while a sibling translation was not.
    param([string[]] $Changed)
    $set = [System.Collections.Generic.HashSet[string]]::new([string[]]$Changed, [StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $Changed) {
        if ($file -notmatch '(^|/)readme\.md$') { continue }
        $base = $file -replace 'readme\.md$', 'readme'
        $missing = @('.nl.md', '.fr.md') | Where-Object { -not $set.Contains("$base$_") }
        if ($missing) { $file }
    }
}

function Get-ChangedFile([switch] $Staged) {
    $gitArgs = if ($Staged) { @('diff', '--cached', '--name-only', '--diff-filter=ACMR') }
               else { @('status', '--porcelain', '--untracked-files=all') }
    $lines = @(git -C $root -c core.quotepath=false @gitArgs)
    if (-not $Staged) { $lines = $lines | ForEach-Object { $_.Substring(3).Trim('"') } }
    $lines | Where-Object { $_ }
}

function Sync-Generated {
    $problems = @()
    $index = Invoke-Tool 'Update-ScriptIndex.ps1'
    if (-not $index.Ok) { $problems += "Update-ScriptIndex.ps1 failed:`n$($index.Output)" }
    $header = Invoke-Tool 'Update-ReadmeHeader.ps1'
    if (-not $header.Ok) { $problems += "Update-ReadmeHeader.ps1 reports a problem (missing language version?):`n$($header.Output)" }
    $links = Invoke-Tool 'Test-MarkdownLinks.ps1'
    if (-not $links.Ok) { $problems += "Broken markdown links:`n$($links.Output)" }
    $modules = Invoke-Tool 'Test-RequiredModules.ps1'
    if (-not $modules.Ok) { $problems += "Modules used by a script but missing from scripts/Startup/RequiredModules.psd1 (load.ps1 would never install them):`n$($modules.Output)" }
    , $problems
}

switch ($Mode) {
    'PostEdit' {
        $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
        $path = $payload.tool_input.file_path
        if (-not $path) { exit 0 }
        $full = [System.IO.Path]::GetFullPath($path)
        if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { exit 0 }
        if ($full -notmatch '\.(ps1|psd1|md)$' -or $full -match '[\\/]\.claude[\\/]') { exit 0 }

        # Edits arrive in bursts and each run rewrites INDEX.md: one run at a time.
        $mutex = [System.Threading.Mutex]::new($false, 'Global\M365-Scripts-sync-docs')
        [void] $mutex.WaitOne()
        try { $problems = Sync-Generated } finally { $mutex.ReleaseMutex() }
        if ($problems.Count) {
            [Console]::Error.WriteLine("Documentation check after editing $([System.IO.Path]::GetFileName($full)):`n`n" + ($problems -join "`n`n"))
            exit 2
        }
        exit 0
    }

    'Stop' {
        $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
        if ($payload.stop_hook_active) { exit 0 }   # already reminded once this turn
        $changed  = @(Get-ChangedFile)
        $reasons  = @()
        $behind   = @(Get-UntranslatedReadme -Changed $changed)
        if ($behind.Count) {
            $reasons += "These English readmes changed but their Dutch/French versions did not: $($behind -join ', '). Apply the same change to readme.nl.md and readme.fr.md."
        }
        # A changed script whose folder readme was not touched at all.
        $undocumented = @($changed | Where-Object { $_ -match '^scripts/.+\.ps1$' } | Where-Object {
            $folder = ($_ -replace '/[^/]+$', '')
            -not ($changed | Where-Object { $_ -like "$folder/readme*.md" })
        })
        if ($undocumented.Count) {
            $reasons += "These scripts changed while no readme in their folder did: $($undocumented -join ', '). If the change is functional, update the folder readme and the root Version History in all three languages (and menu.ps1 when a script was added, renamed or removed)."
        }
        if ($reasons.Count) {
            @{ decision = 'block'
               reason   = (($reasons -join "`n") + "`nFollow .claude/CLAUDE.md, run Update-ReadmeHeader.ps1 and Test-MarkdownLinks.ps1, then commit and push. If nothing needs to change, say so and stop.") } |
                ConvertTo-Json -Compress
        }
        exit 0
    }

    'PreCommit' {
        $staged = @(Get-ChangedFile -Staged)
        $problems = Sync-Generated
        # Re-stage what the generators rewrote, never someone's unrelated work: a readme
        # only when it is already part of this commit, and the index only when every
        # script it lists is committed or in this commit. INDEX.md is built from the files
        # on disk, so a script still being written would otherwise land in the index as
        # a link to a file GitHub does not have.
        $restage = @($staged | Where-Object { $_ -match 'readme(\.nl|\.fr)?\.md$' })
        $wip = @(git -C $root -c core.quotepath=false ls-files --others --exclude-standard -- 'scripts/*.ps1') +
               @(git -C $root -c core.quotepath=false diff --name-only --diff-filter=A -- 'scripts/*.ps1') |
               Where-Object { $_ -and $_ -notin $staged }
        if ($wip.Count) {
            Write-Host "NOTE: scripts/INDEX.md not added to this commit - it would list scripts that are not committed: $($wip -join ', ')" -ForegroundColor Yellow
        } else {
            $restage += 'scripts/INDEX.md'
        }
        if ($restage.Count) { git -C $root add -- @restage }

        $behind = @(Get-UntranslatedReadme -Changed @(Get-ChangedFile -Staged))
        if ($behind.Count) {
            Write-Host "WARNING: English readme changed without its Dutch/French version: $($behind -join ', ')" -ForegroundColor Yellow
            Write-Host "         Ask Claude to update the translations." -ForegroundColor Yellow
        }
        $unlisted = @($problems | Where-Object { $_ -like 'Modules used by a script*' })
        if ($unlisted.Count) {
            Write-Host "WARNING: $($unlisted -join "`n")" -ForegroundColor Yellow
        }
        $blocking = @($problems | Where-Object { $_ -like 'Broken markdown links*' })
        if ($blocking.Count) {
            Write-Host ($blocking -join "`n") -ForegroundColor Red
            Write-Host 'Commit stopped: fix the links above (or commit with --no-verify if you must).' -ForegroundColor Red
            exit 1
        }
        exit 0
    }
}
