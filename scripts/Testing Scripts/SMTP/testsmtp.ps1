#Requires -Version 5.1
# ==============================================================================
# testsmtp.ps1
# One-time SMTP connectivity test — prompts for credentials interactively.
#
# Usage:
#   .\testsmtp.ps1
#   Prompts for password at runtime — no credentials stored on disk.
# ==============================================================================

# ==============================================================================
# CONFIGURATION — change these per test
# ==============================================================================

$SMTPServer = "smtp.office365.com"
$SMTPPort   = 587

$From    = "sender@domain.com"
$To      = "recipient@domain.com"
$Subject = "SMTP Test — PowerShell"
$Body    = "This is a test email sent from PowerShell via SMTP with STARTTLS."

# ==============================================================================
# SCRIPT
# ==============================================================================

$Password   = Read-Host -AsSecureString "Enter SMTP password for $From"
$Credential = New-Object System.Management.Automation.PSCredential($From, $Password)

Write-Host "Connecting to $SMTPServer`:$SMTPPort ..." -ForegroundColor Cyan

try {
    $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
    $smtp.EnableSsl             = $true
    $smtp.Credentials          = $Credential.GetNetworkCredential()
    $smtp.DeliveryMethod       = [System.Net.Mail.SmtpDeliveryMethod]::Network

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
