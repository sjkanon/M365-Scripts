$users = (Get-Recipient | Where-Object 'RecipientTypeDetails' -eq 'UserMailbox').PrimarySmtpAddress
foreach ($user in $users) {
$Delegateds = get-mailbox -ResultSize Unlimited | Get-MailboxPermission -User $user

foreach ($Delegated in $Delegateds) {
$smb = (get-mailbox -Identity $Delegated.Identity).primarysmtpaddress

        $Mailboxpermissies = [pscustomobject][ordered]@{
            Gebruiker      = $user
            Postvak        = $smb
        }

        $Mailboxpermissies | Export-CSV -Path ./lelymachtigingen.csv -Append -NoTypeInformation

}}