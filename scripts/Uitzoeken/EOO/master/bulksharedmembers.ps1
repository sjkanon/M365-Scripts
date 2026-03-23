$users = (Import-Csv .\boetechshared.csv)
ForEach ($user in $users) {
$identity = $user.mailbox
$lid =  $user.gebruiker

add-MailboxPermission -Identity $identity -User $lid -AccessRights FullAccess -InheritanceType All -AutoMapping $false
Remove-RecipientPermission $identity -AccessRights SendAs -Trustee $lid -Confirm:$false
Add-RecipientPermission $identity -AccessRights SendAs -Trustee $lid -Confirm:$false
}