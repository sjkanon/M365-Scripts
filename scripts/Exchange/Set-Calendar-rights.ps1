#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Geeft een gebruiker toegangsrechten op de agenda van een andere gebruiker.

.DESCRIPTION
    Voegt een MailboxFolderPermission toe op de agendasmap van de opgegeven mailbox.
    Werkt voor Nederlandstalige (\Agenda), Franstalige (\Calendrier) en Engelstalige (\Calendar) mailboxen.
    Geschikt voor Belgische omgevingen met gemengde NL/FR taalinstellingen.
    De standaarddomeinnaam wordt automatisch opgehaald via Exchange Online.

.PARAMETER User
    De gebruikersnaam (zonder domein) die de rechten ontvangt.
    Voorbeeld: Sjoerd.Kanon

.PARAMETER TargetMailbox
    De gebruikersnaam (zonder domein) van de mailbox waarop rechten worden gezet.
    Voorbeeld: Jan.Jansen

.PARAMETER AccessRights
    Het toegangsniveau dat wordt toegekend. Geldige waarden:
    Owner, PublishingEditor, Editor, PublishingAuthor, Author,
    NonEditingAuthor, Reviewer, Contributor, AvailabilityOnly, LimitedDetails

.EXAMPLE
    .\Set-CalendarRights.ps1 -User Sjoerd.Kanon -TargetMailbox Jan.Jansen -AccessRights Reviewer

.EXAMPLE
    .\Set-CalendarRights.ps1 -User Sjoerd.Kanon -TargetMailbox Jan.Jansen -AccessRights Editor -WhatIf

.NOTES
    Vereist een actieve Exchange Online verbinding (Connect-ExchangeOnline).
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [string]$User,

    [Parameter(Mandatory)]
    [string]$TargetMailbox,

    [Parameter(Mandatory)]
    [ValidateSet(
        'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author',
        'NonEditingAuthor', 'Reviewer', 'Contributor', 'AvailabilityOnly', 'LimitedDetails'
    )]
    [string]$AccessRights
)

# Haal het standaarddomein op via Exchange Online (vervangt Get-MsolDomain)
$defaultDomain = (Get-AcceptedDomain | Where-Object { $_.Default }).DomainName
if (-not $defaultDomain) {
    throw 'Kan het standaarddomein niet ophalen. Controleer of je verbonden bent via Connect-ExchangeOnline.'
}

$userUPN    = "$User@$defaultDomain"
$targetUPN  = "$TargetMailbox@$defaultDomain"

# Alle paden om locale-varianten (NL/FR/EN) te dekken
$calendarPaths = @(
    "$targetUPN`:\Agenda",       # Nederlands
    "$targetUPN`:\Calendrier",   # Frans (België)
    "$targetUPN`:\Calendar"      # Engels
)

Write-Verbose "Gebruiker  : $userUPN"
Write-Verbose "Doelmap    : $targetUPN"
Write-Verbose "Rechten    : $AccessRights"

foreach ($path in $calendarPaths) {
    try {
        if ($PSCmdlet.ShouldProcess($path, "Add-MailboxFolderPermission ($AccessRights) voor $userUPN")) {
            Add-MailboxFolderPermission -Identity $path -User $userUPN -AccessRights $AccessRights -ErrorAction Stop
            Write-Host "Rechten toegevoegd op: $path" -ForegroundColor Green
        }
    }
    catch [System.Exception] {
        # Map bestaat niet in deze locale of recht bestaat al — overslaan
        Write-Verbose "Overgeslagen ($path): $($_.Exception.Message)"
    }
}

# Toon huidige rechten als verificatie
Write-Host "`nHuidige agendarechten:" -ForegroundColor Cyan
foreach ($path in $calendarPaths) {
    $perms = Get-MailboxFolderPermission -Identity $path -ErrorAction SilentlyContinue
    if ($perms) {
        $perms | Where-Object { $_.User.DisplayName -notin @('Default', 'Anonymous') } |
            Select-Object Identity, User, AccessRights |
            Format-Table -AutoSize
    }
}
