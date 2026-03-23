$users = Import-excel -Path '.\Planning, gebruikers, mappen, ed..xlsx' -WorksheetName gebruikers

foreach ($user in $users) {
 $firstname = $user.Voornaam
 $lastname = $user.Achternaam
 $functie = $user.Functie
$BIG = $user.'BIG-nummer'
$workhours = $user.'Aanwezig op:'


$Body = @"
<p>Met vriendelijke groet,</p>
<p>$firstname $lastname</span><br>$functie</span><br>$BIG<br>$werkdagen&nbsp;</span></p>
<img src="http://handtekening.eoo.support/FysioNU/signature.png" alt="FysioNU">
"@

$Body
}

