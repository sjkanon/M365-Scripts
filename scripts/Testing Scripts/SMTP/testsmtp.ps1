# SMTP Server Configuration
$SMTPServer = "smtp.office365.com"  # Adjust if using a different server
$SMTPPort = 587

# Email Configuration
$From = "smtp-auth0@vias.be"
$To = "sjoerd.kanon@first.eu"
$Subject = "Test Email from PowerShell"
$Body = "This is a test email sent from PowerShell using SMTP with STARTTLS"

# Create secure credential prompt
$Password = Read-Host -AsSecureString "Enter email password"
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, $Password

# Configure and send the email
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
    Write-Host "Email sent successfully!" -ForegroundColor Green
}
Catch {
    Write-Host "An error occurred while sending the email:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}