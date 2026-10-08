# The PowerShell modules this repository depends on — the one list that
# Install-Modules.ps1, Update-Modules.ps1 and load.ps1 all read.
#
# Add a module here and the next start of load.ps1 offers to install it on every
# machine; raise a MinimumVersion and it offers the update.
#
#   Name              Module name on the PowerShell Gallery
#   MinimumVersion    Older than this counts as outdated, not just "update available"
#   WindowsOnly       Skipped on macOS and Linux
#   MinimumPSVersion  Skipped on an older PowerShell (PnP.PowerShell 3 needs 7.4)
#   ImportAtStartup   Imported by load.ps1 before the menu opens
@{
    Modules = @(
        @{ Name = 'ExchangeOnlineManagement';                     MinimumVersion = '3.0.0'; ImportAtStartup = $true }
        @{ Name = 'Microsoft.Graph.Authentication';               MinimumVersion = '2.0.0'; ImportAtStartup = $true }
        @{ Name = 'Microsoft.Graph.Sites';                        MinimumVersion = '2.0.0'; ImportAtStartup = $true }
        @{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinimumVersion = '2.0.0'; ImportAtStartup = $true }
        @{ Name = 'Microsoft.Graph.Identity.SignIns';             MinimumVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Identity.Governance';          MinimumVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Applications';                 MinimumVersion = '2.0.0'; ImportAtStartup = $true }
        @{ Name = 'Microsoft.Graph.Calendar';                     MinimumVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Groups';                       MinimumVersion = '2.0.0'; ImportAtStartup = $true }
        @{ Name = 'Microsoft.Graph.Users';                        MinimumVersion = '2.0.0'; ImportAtStartup = $true }
        @{ Name = 'PnP.PowerShell';                               MinimumVersion = '2.0.0'; MinimumPSVersion = '7.4' }
        @{ Name = 'MicrosoftTeams';                               MinimumVersion = '5.0.0' }
        @{ Name = 'ImportExcel';                                  MinimumVersion = '7.0.0' }
        @{ Name = 'WindowsAutopilotIntune';                       WindowsOnly = $true }
        @{ Name = 'IntuneWin32App';                               WindowsOnly = $true }
    )
}
