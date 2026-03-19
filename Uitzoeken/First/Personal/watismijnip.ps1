$jsonip = Invoke-WebRequest https://ipinfo.io/json
$ipinfo = Convertfrom-JSON $jsonip

$ipinfo.ip | Set-Clipboard

