# SMTP Server Configuration
$SMTPServer = "smtp.office365.com"  
$SMTPPort = 587

# Email Configuration
$From = "info@mobisafetyscan.be"
$User = "outsystems@vias.be"
$To = "sjoerd.kanon@first.eu"
$Subject = "Test Email from PowerShell"
$Body = "This is a test email sent from PowerShell using SMTP with STARTTLS"

# Beveiligd wachtwoord opslaan (Eenmalig instellen en ophalen bij elk scriptuitvoering)
$SecurePasswordPath = "$env:USERPROFILE\secure_smtp_password.txt"

# Opslaan van wachtwoord in versleutelde vorm (Doe dit één keer handmatig en verwijder de commentaarregel)
# Read-Host -AsSecureString "Enter email password" | ConvertFrom-SecureString | Set-Content $SecurePasswordPath

# Ophalen en decoderen van wachtwoord
$SecurePassword = Get-Content $SecurePasswordPath | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $SecurePassword

# Functie om e-mail te verzenden
Function Send-TestEmail {
    Try {
        $SMTPMessage = @{
            From = $From
            To = $To
            Subject = $Subject
            Body = $Body
            SmtpServer = $SMTPServer
            Port = $SMTPPort
            Credential = $Credential
            UseSSL = $true
            ErrorAction = 'Stop'
        }

        Send-MailMessage @SMTPMessage
        Write-Host "$(Get-Date) - Email sent successfully!" -ForegroundColor Green
    }
    Catch {
        Write-Host "$(Get-Date) - An error occurred while sending the email:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

# Herhaal elke 5 minuten
While ($true) {
    Send-TestEmail
    Start-Sleep -Seconds 300  # 5 minuten wachten
}
