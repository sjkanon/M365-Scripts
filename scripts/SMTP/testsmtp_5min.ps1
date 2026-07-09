#Requires -Version 5.1
<#
.SYNOPSIS
    Recurring SMTP test — sends a test email every N minutes until stopped.

.DESCRIPTION
    Sends a test email on a fixed interval. Useful for sustained relay testing
    or reproducing intermittent failures. Stop with Ctrl+C.

    Windows: uses a saved encrypted password file (DPAPI). Save it once:
        Read-Host -AsSecureString "Enter SMTP password" | ConvertFrom-SecureString | Set-Content "$env:USERPROFILE\smtp_test_password.txt"

    macOS / Linux: DPAPI is not available — prompts for password once at startup
    and keeps it in memory for the session.

.PARAMETER SmtpServer
    SMTP server hostname. Default: smtp.office365.com

.PARAMETER Port
    SMTP port. Default: 587

.PARAMETER From
    Sender address (displayed in the From header).

.PARAMETER AuthAs
    SMTP auth username. Defaults to $From. Set this when sending from a shared
    mailbox: $From = shared mailbox, $AuthAs = account with Send As permission.

.PARAMETER To
    Recipient address for the test email.

.PARAMETER Subject
    Email subject. Default: "SMTP Recurring Test — PowerShell"

.PARAMETER Body
    Email body. Default: generic test message.

.PARAMETER IntervalSeconds
    Seconds between sends. Default: 300 (5 minutes).

.PARAMETER SavedKeyPath
    Path to the saved encrypted password file.
    Default: $env:USERPROFILE\smtp_test_password.txt

.EXAMPLE
    .\testsmtp_5min.ps1 -From "sender@domain.com" -To "recipient@domain.com"

.EXAMPLE
    .\testsmtp_5min.ps1 -From "shared@domain.com" -AuthAs "user@domain.com" -To "recipient@domain.com" -IntervalSeconds 60

.NOTES
    Author  : Sjoerd Kanon
    Version : 2.1
    Platform: Windows (saved file), macOS/Linux (interactive prompt once)
#>

[CmdletBinding()]
param (
    [string] $SmtpServer      = "smtp.office365.com",
    [int]    $Port            = 587,
    [string] $From            = "sender@domain.com",
    [string] $AuthAs          = "",
    [string] $To              = "recipient@domain.com",
    [string] $Subject         = "SMTP Recurring Test — PowerShell",
    [string] $Body            = "This is a recurring test email sent from PowerShell via SMTP with STARTTLS.",
    [int]    $IntervalSeconds = 300,
    [string] $SavedKeyPath    = "$env:USERPROFILE\smtp_test_password.txt"
)

if (-not $AuthAs) { $AuthAs = $From }

# Platform detection — $IsWindows is undefined in PS5.1 (Windows only), $true in PS6+ on Windows
$onWindows = ($IsWindows -eq $true) -or ($null -eq (Get-Variable IsWindows -ErrorAction SilentlyContinue).Value)

if ($onWindows) {
    if (-not (Test-Path $SavedKeyPath)) {
        Write-Host "No saved password found at: $SavedKeyPath" -ForegroundColor Yellow
        Write-Host "Run this once to save it:" -ForegroundColor Yellow
        Write-Host "  Read-Host -AsSecureString 'Enter SMTP password' | ConvertFrom-SecureString | Set-Content `"$SavedKeyPath`"" -ForegroundColor Cyan
        exit 1
    }
    $SecurePassword = Get-Content $SavedKeyPath | ConvertTo-SecureString
} else {
    # macOS / Linux — DPAPI not available, prompt once and keep in memory
    Write-Host "macOS/Linux detected — password will be prompted once and kept in memory." -ForegroundColor Cyan
    $SecurePassword = Read-Host -AsSecureString "Enter SMTP password for $AuthAs"
}

$Credential = New-Object System.Management.Automation.PSCredential($AuthAs, $SecurePassword)

function Send-TestEmail {
    try {
        $smtp                = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
        $smtp.EnableSsl      = $true
        $smtp.Credentials    = $Credential.GetNetworkCredential()
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network

        $mail         = New-Object System.Net.Mail.MailMessage
        $mail.From    = $From
        $mail.To.Add($To)
        $mail.Subject = $Subject
        $mail.Body    = $Body

        $smtp.Send($mail)
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sent successfully to $To" -ForegroundColor Green
    } catch {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Failed: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        if ($mail) { $mail.Dispose() }
        if ($smtp) { $smtp.Dispose() }
    }
}

Write-Host "Starting recurring SMTP test — every $($IntervalSeconds / 60) min. Ctrl+C to stop." -ForegroundColor Cyan
Write-Host "Server : $SmtpServer`:$Port  |  From: $From  →  To: $To"
Write-Host ""

while ($true) {
    Send-TestEmail
    Start-Sleep -Seconds $IntervalSeconds
}
