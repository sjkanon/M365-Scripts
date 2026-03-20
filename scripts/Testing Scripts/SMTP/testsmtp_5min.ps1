#Requires -Version 5.1
# ==============================================================================
# testsmtp_5min.ps1
# Recurring SMTP test — sends a test email every 5 minutes until stopped.
# Uses a saved encrypted password so it can run unattended.
#
# First run — save password to disk (run once, then remove the line):
#   Read-Host -AsSecureString "Enter SMTP password" | ConvertFrom-SecureString | Set-Content "$env:USERPROFILE\smtp_test_password.txt"
#
# Then run the script normally:
#   .\testsmtp_5min.ps1
#
# Stop with Ctrl+C.
#
# Note: ConvertFrom-SecureString uses Windows DPAPI — the saved file can only
#       be decrypted by the same user on the same machine.
# ==============================================================================

# ==============================================================================
# CONFIGURATION — change these per test
# ==============================================================================

$SMTPServer = "smtp.office365.com"
$SMTPPort   = 587

$From    = "sender@domain.com"       # SMTP auth username
$AuthAs  = "sender@domain.com"       # If different from $From (e.g. shared mailbox), set here
$To      = "recipient@domain.com"
$Subject = "SMTP Recurring Test — PowerShell"
$Body    = "This is a recurring test email sent from PowerShell via SMTP with STARTTLS."

$IntervalSeconds = 300                # 5 minutes

# ==============================================================================
# CREDENTIALS
# ==============================================================================

$SecurePasswordPath = "$env:USERPROFILE\smtp_test_password.txt"

if (-not (Test-Path $SecurePasswordPath)) {
    Write-Host "No saved password found at $SecurePasswordPath" -ForegroundColor Yellow
    Write-Host "Run this once to save it:" -ForegroundColor Yellow
    Write-Host "  Read-Host -AsSecureString 'Enter SMTP password' | ConvertFrom-SecureString | Set-Content `"$SecurePasswordPath`"" -ForegroundColor Cyan
    exit 1
}

$SecurePassword = Get-Content $SecurePasswordPath | ConvertTo-SecureString
$Credential     = New-Object System.Management.Automation.PSCredential($AuthAs, $SecurePassword)

# ==============================================================================
# FUNCTION
# ==============================================================================

function Send-TestEmail {
    try {
        $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
        $smtp.EnableSsl       = $true
        $smtp.Credentials     = $Credential.GetNetworkCredential()
        $smtp.DeliveryMethod  = [System.Net.Mail.SmtpDeliveryMethod]::Network

        $mail         = New-Object System.Net.Mail.MailMessage
        $mail.From    = $From
        $mail.To.Add($To)
        $mail.Subject = $Subject
        $mail.Body    = $Body

        $smtp.Send($mail)
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Email sent successfully to $To" -ForegroundColor Green
    } catch {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Failed to send email:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    } finally {
        if ($mail) { $mail.Dispose() }
        if ($smtp) { $smtp.Dispose() }
    }
}

# ==============================================================================
# LOOP
# ==============================================================================

Write-Host "Starting recurring SMTP test — sending every $($IntervalSeconds / 60) minutes. Press Ctrl+C to stop." -ForegroundColor Cyan
Write-Host "Server : $SMTPServer`:$SMTPPort"
Write-Host "From   : $From  →  To: $To"
Write-Host ""

while ($true) {
    Send-TestEmail
    Start-Sleep -Seconds $IntervalSeconds
}
