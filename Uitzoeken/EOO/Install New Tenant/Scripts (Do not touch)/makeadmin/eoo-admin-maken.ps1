function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
 
$password = Get-RandomCharacters -length 13 -characters 'abcdefghiklmnoprstuvwxyz'
$password += Get-RandomCharacters -length 2 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password += Get-RandomCharacters -length 1 -characters '1234567890'
$password += Get-RandomCharacters -length 2 -characters '!"$%&/()=?}][{@#*+'
$password = Scramble-String $password

$MsolDomains = Get-MsolDomain -TenantId $cid

$regex = '^[^.]*\.onmicrosoft\.com$'
        $Domain = $MsolDomains |
            Where-Object Name -Match $regex |
            Select-Object -ExpandProperty Name |
            Select-Object -First 1


new-mailbox -shared -Name "EOO Beheeraccount" -alias eooadmin -PrimarySmtpAddress "eooadmin@$domain"
sleep 20
$defaultdomain = (Get-MsolDomain -TenantId $cid | Where-Object isdefault -eq 'true').name
Set-MsolUserPrincipalName -TenantId $cid -UserPrincipalName "eooadmin@$defaultdomain" -NewUserPrincipalName "eooadmin@$domain"
set-MsolUser -TenantId $cid -UserPrincipalName eooadmin@$domain -FirstName EOO -LastName Admin
Set-MsolUserPassword -ForceChangePassword $false -TenantId $cid -UserPrincipalName "eooadmin@$domain" -NewPassword $password
Add-MsolRoleMember -RoleMemberEmailAddress "eooadmin@$domain" -RoleName "company administrator" -TenantId $cid