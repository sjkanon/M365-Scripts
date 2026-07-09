#Requires -Version 5.1
<#
.SYNOPSIS
    One-time SMTP connectivity test.

.DESCRIPTION
    Sends a single test email via SMTP with STARTTLS. Prompts for credentials
    interactively — nothing is stored on disk.

.PARAMETER SmtpServer
    SMTP server hostname. Default: smtp.office365.com

.PARAMETER Port
    SMTP port. Default: 587

.PARAMETER From
    Sender address (also used as SMTP auth username).

.PARAMETER To
    Recipient address for the test email.

.PARAMETER Subject
    Email subject. Default: "SMTP Test — PowerShell"

.PARAMETER Body
    Email body. Default: generic test message.

.EXAMPLE
    .\testsmtp.ps1 -From "sender@domain.com" -To "recipient@domain.com"

.EXAMPLE
    .\testsmtp.ps1 -From "sender@domain.com" -To "recipient@domain.com" -SmtpServer "mail.domain.com" -Port 25

.NOTES
    Author  : Sjoerd Kanon
    Version : 2.0
#>

[CmdletBinding()]
param (
    [string] $SmtpServer = "smtp.office365.com",
    [int]    $Port       = 587,
    [string] $From       = "sender@domain.com",
    [string] $To         = "recipient@domain.com",
    [string] $Subject    = "SMTP Test — PowerShell",
    [string] $Body       = "This is a test email sent from PowerShell via SMTP with STARTTLS."
)

$Password   = Read-Host -AsSecureString "Enter SMTP password for $From"
$Credential = New-Object System.Management.Automation.PSCredential($From, $Password)

Write-Host "Connecting to $SmtpServer`:$Port ..." -ForegroundColor Cyan

try {
    $smtp                  = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
    $smtp.EnableSsl        = $true
    $smtp.Credentials      = $Credential.GetNetworkCredential()
    $smtp.DeliveryMethod   = [System.Net.Mail.SmtpDeliveryMethod]::Network

    $mail         = New-Object System.Net.Mail.MailMessage
    $mail.From    = $From
    $mail.To.Add($To)
    $mail.Subject = $Subject
    $mail.Body    = $Body

    $smtp.Send($mail)
    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Email sent successfully to $To" -ForegroundColor Green
} catch {
    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Failed to send email:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
} finally {
    if ($mail) { $mail.Dispose() }
    if ($smtp) { $smtp.Dispose() }
}
