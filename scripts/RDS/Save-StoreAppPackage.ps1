#Requires -Version 5.1
<#
.SYNOPSIS
    Download Photos and Snipping Tool as complete, Microsoft-signed package files - the
    bundle and the frameworks it needs - for provisioning on session hosts, without the
    Store app, winget or a sign-in.

.DESCRIPTION
    Snipping Tool and Photos have no installer. The Store updated them per user on the
    session hosts and kept only the main package, not the bundle: every registration
    of such a build fails with 0x80070490, and Windows refuses to provision an app
    from its installed folder (Add-AppxProvisionedPackage -FolderPath: "No applicable
    main package was found"). Provisioning needs the package files.

    This script asks Microsoft for them the way the Store does, anonymously: the
    product's Windows Update category from the display catalog
    (displaycatalog.mp.microsoft.com), the packages in it from the delivery service
    (fe3.delivery.mp.microsoft.com, SyncUpdates), and a download link per file
    (GetExtendedUpdateInfo2). Per app:

      1. Bundle      The newest bundle (or -Version), downloaded to
                     <OutputFolder>\<package name>.
      2. Check       Size and SHA-256 against what the delivery service published, and
                     a valid Authenticode signature by Microsoft Corporation. Anything
                     that does not match is deleted.
      3. Frameworks  The bundle's <Architecture> package is read for its dependencies;
                     for each, the newest framework of that name in the same category
                     at or above the minimum version is downloaded and checked the same
                     way.
      4. Swap        The folder then holds exactly that bundle and those frameworks -
                     the layout Repair-AppxPackageStore.ps1 -StoreFolder,
                     Repair-M365AppsHost.ps1 and the watchdog provision from. No license
                     file: the download is anonymous, so provisioning runs with
                     -SkipLicense, as for an image.

    -CopyTo copies the folder to other hosts' C:\IT\AppWatchdog\Packages over their
    admin share (no WinRM). Nothing is installed or provisioned here.

.PARAMETER App
    Photos, SnippingTool or both (default both).

.PARAMETER OutputFolder
    Where the package folders go. Default C:\IT\AppWatchdog\Packages.

.PARAMETER Architecture
    The hosts' architecture, for the frameworks. Default x64.

.PARAMETER Version
    A bundle version to take instead of the newest, e.g. 2022.2608.41.0. Only with a
    single -App.

.PARAMETER CopyTo
    Hosts to copy the result to, e.g. lem-avd-4,lem-avd-5.

.EXAMPLE
    # On a session host: newest Photos and Snipping Tool, then to the other hosts
    .\Save-StoreAppPackage.ps1 -CopyTo lem-avd-4,lem-avd-5

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows PowerShell 5.1 or PowerShell 7, elevated for the default folder;
              needs fe3.delivery.mp.microsoft.com, displaycatalog.mp.microsoft.com and
              *.dl.delivery.mp.microsoft.com (the files come over http, each checked by
              SHA-256 and signature)
    Exit    : 0 every app downloaded and checked, 1 otherwise
#>
[CmdletBinding()]
param (
    [ValidateSet('Photos', 'SnippingTool')]
    [string[]] $App = @('Photos', 'SnippingTool'),
    [string]   $OutputFolder = 'C:\IT\AppWatchdog\Packages',
    [ValidateSet('x64', 'arm64', 'x86')]
    [string]   $Architecture = 'x64',
    [string]   $Version,
    [string[]] $CopyTo
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# -- Constants -------------------------------------------------------------------
$Products = @{
    Photos       = '9WZDNCRFJBH4'
    SnippingTool = '9MZ95KL8MR0L'
}
$CatalogUrl = 'https://displaycatalog.mp.microsoft.com/v7.0/products/{0}?market=US&languages=en-US&fieldsTemplate=Details'
$Fe3Url     = 'https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx'
# What a client says it already has: without these lists the service answers nothing.
# The values the Store client sends, as StoreLib (github.com/StoreDev/StoreLib) records them.
$InstalledNonLeaf = (
    '1,2,3,11,19,544,549,2359974,2359977,5169044,8788830,23110993,23110994,54341900,54343656,59830006' +
    ',59830007,59830008,60484010,62450018,62450019,62450020,66027979,66053150,97657898,98822896,98959' +
    '022,98959023,98959024,98959025,98959026,104433538,104900364,105489019,117765322,129905029,130040' +
    '031,132387090,132393049,133399034,138537048,140377312,143747671,158941041,158941042,158941043,15' +
    '8941044,159123858,159130928,164836897,164847386,164848327,164852241,164852246,164852252,16485225' +
    '3'
) -split ','
$OtherCached = (
    '10,17,2359977,5143990,5169043,5169047,8806526,9125350,9154769,10809856,23110995,23110996,2311099' +
    '9,23111000,23111001,23111002,23111003,23111004,24513870,28880263,30077688,30486944,30526991,3052' +
    '8442,30530496,30530501,30530504,30530962,30535326,30536242,30539913,30545142,30545145,30545488,3' +
    '0546212,30547779,30548797,30548860,30549262,30551160,30551161,30551164,30553016,30553744,3055401' +
    '4,30559008,30559011,30560006,30560011,30561006,30563261,30565215,30578059,30664998,30677904,3068' +
    '1618,30682195,30685055,30702579,30708772,30709591,30711304,30715418,30720106,30720273,30732075,3' +
    '0866952,30866964,30870749,30877852,30878437,30890151,30892149,30990917,31049444,31190936,3119696' +
    '1,31197811,31198836,31202713,31203522,31205442,31205557,31207585,31208440,31208451,31209591,3121' +
    '0536,31211625,31212713,31213588,31218518,31219420,31220279,31220302,31222086,31227080,31229030,3' +
    '1238236,31254198,31258008,36436779,36437850,36464012,41916569,47249982,47283134,58577027,5857804' +
    '0,58578041,58628920,59107045,59125697,59142249,60466586,60478936,66450441,66467021,66479051,7520' +
    '2978,77436021,77449129,85159569,90199702,90212090,96911147,97110308,98528428,98665206,98837995,9' +
    '8842922,98842977,98846632,98866485,98874250,98879075,98904649,98918872,98945691,98959458,9898470' +
    '7,100220125,100238731,100662329,100795834,100862457,103124811,103348671,104369981,104372472,1043' +
    '85324,104465831,104465834,104467697,104473368,104482267,104505005,104523840,104550085,104558084,' +
    '104659441,104659675,104664678,104668274,104671092,104673242,104674239,104679268,104686047,104698' +
    '649,104751469,104752478,104755145,104761158,104762266,104786484,104853747,104873258,104983051,10' +
    '5063056,105116588,105178523,105318602,105362613,105364552,105368563,105369591,105370746,10537350' +
    '3,105373615,105376634,105377546,105378752,105379574,105381626,105382587,105425313,105495146,1058' +
    '62607,105939029,105995585,106017178,106129726,106768485,107825194,111906429,115121473,115578654,' +
    '116630363,117835105,117850671,118638500,118662027,118872681,118873829,118879289,118889092,119501' +
    '720,119551648,119569538,119640702,119667998,119674103,119697201,119706266,119744627,119773746,12' +
    '0072697,120144309,120214154,120357027,120392612,120399120,120553945,120783545,120797092,12088167' +
    '6,120889689,120999554,121168608,121268830,121341838,121729951,121803677,122165810,125408034,1272' +
    '93130,127566683,127762067,127861893,128571722,128647535,128698922,128701748,128771507,129037212,' +
    '129079800,129175415,129317272,129319665,129365668,129378095,129424803,129590730,129603714,129625' +
    '954,129692391,129714980,129721097,129886397,129968371,129972243,130009862,130033651,130040030,13' +
    '0040032,130040033,130091954,130100640,130131267,130131921,130144837,130171030,130172071,13019721' +
    '8,130212435,130291076,130402427,130405166,130676169,130698471,130713390,130785217,131396908,1314' +
    '55115,131682095,131689473,131701956,132142800,132525441,132765492,132801275,133399034,134522926,' +
    '134524022,134528994,134532942,134536993,134538001,134547533,134549216,134549317,134550159,134550' +
    '214,134550232,134551154,134551207,134551390,134553171,134553237,134554199,134554227,134555229,13' +
    '4555240,134556118,134557078,134560099,134560287,134562084,134562180,134563287,134565083,13456613' +
    '0,134568111,134624737,134666461,134672998,134684008,134916523,135100527,135219410,135222083,1353' +
    '06997,135463054,135779456,135812968,136097030,136131333,136146907,136157556,136320962,136450641,' +
    '136466000,136745792,136761546,136840245,138160034,138181244,138210071,138210107,138232200,138237' +
    '088,138277547,138287133,138306991,138324625,138341916,138372035,138372036,138375118,138378071,13' +
    '8380128,138380194,138534411,138618294,138931764,139536037,139536038,139536039,139536040,14036783' +
    '2,140406050,140421668,140422973,140423713,140436348,140483470,140615715,140802803,140896470,1411' +
    '89437,141192744,141382548,141461680,141624996,141627135,141659139,141872038,141993721,142006413,' +
    '142045136,142095667,142227273,142250480,142518788,142544931,142546314,142555433,142653044,143191' +
    '852,143258496,143299722,143331253,143432462,143632431,143695326,144219522,144590916,145410436,14' +
    '6720405,150810438,151258773,151315554,151400090,151429441,151439617,151453617,151466296,15151113' +
    '2,151636561,151823192,151827116,151850642,152016572,153111675,153114652,153123147,153267108,1533' +
    '89799,153395366,153718608,154171028,154315227,154559688,154978771,154979742,154985773,154989370,' +
    '155044852,155065458,155578573,156403304,159085959,159776047,159816630,160733048,160733049,160733' +
    '050,160733051,160733056,164824922,164824924,164824926,164824930,164831646,164831647,164831648,16' +
    '4831650,164835050,164835051,164835052,164835056,164835057,164835059,164836898,164836899,16483690' +
    '0,164845333,164845334,164845336,164845337,164845341,164845342,164845345,164845346,164845349,1648' +
    '45350,164845353,164845355,164845358,164845361,164845364,164847387,164847388,164847389,164847390,' +
    '164848328,164848329,164848330,164849448,164849449,164849451,164849452,164849454,164849455,164849' +
    '457,164849461,164850219,164850220,164850222,164850223,164850224,164850226,164850227,164850228,16' +
    '4850229,164850231,164850236,164850237,164850240,164850242,164850243,164852242,164852243,16485224' +
    '4,164852247,164852248,164852249,164852250,164852251,164852254,164852256,164852257,164852258,1648' +
    '52259,164852260,164852261,164852262,164853061,164853063,164853071,164853072,164853075,168118980,' +
    '168118981,168118983,168118984,168180375,168180376,168180378,168180379,168270830,168270831,168270' +
    '833,168270834,168270835'
) -split ','
$DeviceAttributes = 'BranchReadinessLevel=CB;CurrentBranch=rs_prerelease;OEMModel=Virtual Machine;FlightRing=WIS;AttrDataVer=21;SystemManufacturer=Microsoft Corporation;InstallLanguage=en-US;OSUILocale=en-US;InstallationType=Client;FlightingBranchName=external;FirmwareVersion=Hyper-V UEFI Release v2.5;SystemProductName=Virtual Machine;OSSkuId=48;FlightContent=Branch;App=WU;OEMName_Uncleaned=Microsoft Corporation;AppVer=10.0.16184.1001;OSArchitecture=AMD64;SystemSKU=None;UpdateManagementGroup=2;IsFlightingEnabled=1;IsDeviceRetailDemo=0;TelemetryLevel=3;OSVersion=10.0.16184.1001;DeviceFamily=Windows.Desktop;'

# -- Output ----------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ''; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

# -- Delivery service ------------------------------------------------------------
function Get-SoapHeader {
    param([string] $Action, [string] $To)
    $now = [DateTime]::UtcNow
    @"
<s:Header>
  <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/$Action</a:Action>
  <a:MessageID>urn:uuid:$([guid]::NewGuid())</a:MessageID>
  <a:To s:mustUnderstand="1">$To</a:To>
  <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
    <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
      <Created>$($now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'))</Created>
      <Expires>$($now.AddMinutes(5).ToString('yyyy-MM-ddTHH:mm:ss.fffZ'))</Expires>
    </Timestamp>
    <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
      <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL"><User /></TicketType>
    </wuws:WindowsUpdateTicketsToken>
  </o:Security>
</s:Header>
"@
}

function Invoke-Fe3 {
    param([string] $Action, [string] $Body, [switch] $Secured, [switch] $Decode)
    $to       = if ($Secured) { "$Fe3Url/secured" } else { $Fe3Url }
    $envelope = '<s:Envelope xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:s="http://www.w3.org/2003/05/soap-envelope">' +
                (Get-SoapHeader -Action $Action -To $to) + "<s:Body>$Body</s:Body></s:Envelope>"
    $response = Invoke-WebRequest -Uri $to -Method Post -ContentType 'application/soap+xml; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($envelope)) -UseBasicParsing
    # SyncUpdates sends the update XML HTML-encoded inside the SOAP answer; the other
    # answers are plain XML (a link's & stays &amp;).
    if ($Decode) { return [xml] [Net.WebUtility]::HtmlDecode($response.Content) }
    return [xml] $response.Content
}

function Get-Cookie {
    $body = '<GetCookie xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService"><oldCookie></oldCookie>' +
            '<lastChange>2015-10-21T17:01:07.1472913Z</lastChange><currentTime>' + [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ') +
            '</currentTime><protocolVersion>1.40</protocolVersion></GetCookie>'
    return (Invoke-Fe3 -Action 'GetCookie' -Body $body).GetElementsByTagName('EncryptedData')[0].InnerText
}

function Get-CategoryPackage {
    <#
        Every package file the delivery service lists for a Windows Update category:
        identity (name_version_arch__publisher), file name, size, SHA-256 and the update
        identity to ask the download link for.
    #>
    param([string] $Cookie, [string] $CategoryId)
    $ints = { param($list) -join @($list | ForEach-Object { "<int>$_</int>" }) }
    $body = '<SyncUpdates xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">' +
            "<cookie><Expiration>2045-03-11T02:02:48Z</Expiration><EncryptedData>$Cookie</EncryptedData></cookie><parameters>" +
            '<ExpressQuery>false</ExpressQuery>' +
            "<InstalledNonLeafUpdateIDs>$(& $ints $InstalledNonLeaf)</InstalledNonLeafUpdateIDs>" +
            "<OtherCachedUpdateIDs>$(& $ints $OtherCached)</OtherCachedUpdateIDs>" +
            '<SkipSoftwareSync>false</SkipSoftwareSync><NeedTwoGroupOutOfScopeUpdates>true</NeedTwoGroupOutOfScopeUpdates>' +
            "<FilterAppCategoryIds><CategoryIdentifier><Id>$CategoryId</Id></CategoryIdentifier></FilterAppCategoryIds>" +
            '<TreatAppCategoryIdsAsInstalled>true</TreatAppCategoryIdsAsInstalled><AlsoPerformRegularSync>false</AlsoPerformRegularSync><ComputerSpec/>' +
            '<ExtendedUpdateInfoParameters><XmlUpdateFragmentTypes><XmlUpdateFragmentType>Extended</XmlUpdateFragmentType></XmlUpdateFragmentTypes>' +
            '<Locales><string>en-US</string><string>en</string></Locales></ExtendedUpdateInfoParameters>' +
            '<ClientPreferredLanguages><string>en-US</string></ClientPreferredLanguages>' +
            "<ProductsParameters><SyncCurrentVersionOnly>false</SyncCurrentVersionOnly><DeviceAttributes>$DeviceAttributes</DeviceAttributes>" +
            '<CallerAttributes>Interactive=1;IsSeeker=0;</CallerAttributes><Products/></ProductsParameters></parameters></SyncUpdates>'
    $doc = Invoke-Fe3 -Action 'SyncUpdates' -Body $body -Decode
    # NewUpdates carries the update identity, ExtendedUpdateInfo the files; both by ID.
    $identity = @{}
    foreach ($info in $doc.GetElementsByTagName('UpdateInfo')) {
        $id = @($info.GetElementsByTagName('UpdateIdentity'))
        if ($id.Count -gt 0 -and @($info.GetElementsByTagName('SecuredFragment')).Count -gt 0) {
            $identity[[string] $info.ID] = [PSCustomObject]@{ UpdateId = $id[0].UpdateID; Revision = $id[0].RevisionNumber }
        }
    }
    $seen = @{}
    foreach ($file in $doc.GetElementsByTagName('File')) {
        $moniker = $file.GetAttribute('InstallerSpecificIdentifier')
        if (-not $moniker -or $seen.ContainsKey($moniker)) { continue }
        $update = $file.ParentNode.ParentNode.ParentNode
        $updateId = @($update.ChildNodes | Where-Object { $_.LocalName -eq 'ID' } | ForEach-Object { $_.InnerText }) | Select-Object -First 1
        if (-not $identity.ContainsKey([string] $updateId)) { continue }
        $sha256 = @($file.ChildNodes | Where-Object { $_.LocalName -eq 'AdditionalDigest' -and $_.GetAttribute('Algorithm') -eq 'SHA256' } | ForEach-Object { $_.InnerText }) | Select-Object -First 1
        $parts = $moniker -split '_'
        $seen[$moniker] = $true
        [PSCustomObject]@{
            Moniker  = $moniker
            Name     = $parts[0]
            Version  = [version] $parts[1]
            Arch     = $parts[2]
            Ext      = [IO.Path]::GetExtension($file.GetAttribute('FileName'))
            Size     = [long] $file.GetAttribute('Size')
            Sha1     = $file.GetAttribute('Digest')
            Sha256   = $sha256
            UpdateId = $identity[[string] $updateId].UpdateId
            Revision = $identity[[string] $updateId].Revision
        }
    }
}

function Get-FileUrl {
    <# The download link of one package file: the FileLocation whose digest is the file's. #>
    param($Package)
    $body = '<GetExtendedUpdateInfo2 xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService"><updateIDs><UpdateIdentity>' +
            "<UpdateID>$($Package.UpdateId)</UpdateID><RevisionNumber>$($Package.Revision)</RevisionNumber></UpdateIdentity></updateIDs>" +
            '<infoTypes><XmlUpdateFragmentType>FileUrl</XmlUpdateFragmentType><XmlUpdateFragmentType>FileDecryption</XmlUpdateFragmentType></infoTypes>' +
            "<deviceAttributes>$DeviceAttributes</deviceAttributes></GetExtendedUpdateInfo2>"
    $doc = Invoke-Fe3 -Action 'GetExtendedUpdateInfo2' -Body $body -Secured
    foreach ($location in $doc.GetElementsByTagName('FileLocation')) {
        $digest = @($location.ChildNodes | Where-Object { $_.LocalName -eq 'FileDigest' } | ForEach-Object { $_.InnerText }) | Select-Object -First 1
        if ($digest -eq $Package.Sha1) { return @($location.ChildNodes | Where-Object { $_.LocalName -eq 'Url' } | ForEach-Object { $_.InnerText }) | Select-Object -First 1 }
    }
    throw "no download link for $($Package.Moniker)"
}

# -- Files -----------------------------------------------------------------------
function Get-Sha256Base64 {
    param([string] $Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try { return [Convert]::ToBase64String($sha.ComputeHash($stream)) } finally { $stream.Dispose(); $sha.Dispose() }
}

function Save-Package {
    <# Download one package file into $Folder and check size, SHA-256 and Microsoft's signature. Throws when any does not match. #>
    param($Package, [string] $Folder)
    $path = Join-Path $Folder ($Package.Moniker + $Package.Ext)
    Write-Host ('         {0} ({1:N0} MB)' -f $Package.Moniker, ($Package.Size / 1MB))
    # Three tries, each with a fresh link: the links expire, and a large file can break off.
    for ($try = 1; ; $try++) {
        $client = New-Object Net.WebClient
        try { $client.DownloadFile((Get-FileUrl $Package), $path); break }
        catch {
            $why = $_.Exception; while ($why.InnerException) { $why = $why.InnerException }
            Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            if ($try -ge 3) { throw "$($Package.Moniker): download failed 3 times - $($why.Message)" }
            Write-Warn "download try $try failed ($($why.Message)) - again in 10 s"; Start-Sleep -Seconds 10
        } finally { $client.Dispose() }
    }
    $length = (Get-Item -LiteralPath $path).Length
    if ($length -ne $Package.Size) { Remove-Item -LiteralPath $path -Force; throw "$($Package.Moniker): $length bytes, expected $($Package.Size)" }
    if ($Package.Sha256 -and (Get-Sha256Base64 $path) -ne $Package.Sha256) { Remove-Item -LiteralPath $path -Force; throw "$($Package.Moniker): SHA-256 does not match what Microsoft published" }
    $signature = Get-AuthenticodeSignature -FilePath $path
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        Remove-Item -LiteralPath $path -Force; throw "$($Package.Moniker): not validly signed by Microsoft ($($signature.Status))"
    }
    return $path
}

function Get-BundleDependency {
    <# The package dependencies of the bundle's application package for $Architecture: name and minimum version. #>
    param([string] $BundlePath)
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $bundle = [IO.Compression.ZipFile]::OpenRead($BundlePath)
    $temp   = Join-Path ([IO.Path]::GetTempPath()) ('bundle_{0}.appx' -f [guid]::NewGuid().ToString('N'))
    try {
        $entry = $bundle.GetEntry('AppxMetadata/AppxBundleManifest.xml')
        $reader = New-Object IO.StreamReader($entry.Open())
        try { $manifest = [xml] $reader.ReadToEnd() } finally { $reader.Dispose() }
        $package = @($manifest.GetElementsByTagName('Package') | Where-Object { $_.GetAttribute('Type') -eq 'application' -and $_.GetAttribute('Architecture') -eq $Architecture }) | Select-Object -First 1
        if (-not $package) { throw "the bundle has no $Architecture application package" }
        $inner = $bundle.GetEntry($package.GetAttribute('FileName'))
        $in = $inner.Open(); $out = [IO.File]::Create($temp)
        try { $in.CopyTo($out) } finally { $out.Dispose(); $in.Dispose() }
        $appx = [IO.Compression.ZipFile]::OpenRead($temp)
        try {
            $reader = New-Object IO.StreamReader($appx.GetEntry('AppxManifest.xml').Open())
            try { $appManifest = [xml] $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $appx.Dispose() }
        foreach ($dep in $appManifest.GetElementsByTagName('PackageDependency')) {
            [PSCustomObject]@{ Name = $dep.GetAttribute('Name'); MinVersion = [version] $dep.GetAttribute('MinVersion') }
        }
    } finally {
        $bundle.Dispose()
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}

# -- One run -------------------------------------------------------------------------
if ($Version -and $App.Count -ne 1) { Write-Error '-Version needs exactly one -App'; exit 1 }
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

$exitCode = 0
$results  = [System.Collections.Generic.List[object]]::new()
$cookie   = Get-Cookie
foreach ($name in $App) {
    Write-Step "== $name"
    $row = [PSCustomObject]@{ App = $name; Bundle = ''; Frameworks = ''; Folder = ''; Result = 'failed' }
    $staging = $null
    try {
        $catalog = Invoke-RestMethod -Uri ($CatalogUrl -f $Products[$name]) -UseBasicParsing
        $fulfilment = $catalog.Product.DisplaySkuAvailabilities[0].Sku.Properties.FulfillmentData
        if ($fulfilment -is [string]) { $fulfilment = $fulfilment | ConvertFrom-Json }
        $packageName = $fulfilment.PackageFamilyName -replace '_[^_]+$'
        $all = @(Get-CategoryPackage -Cookie $cookie -CategoryId $fulfilment.WuCategoryId)
        $bundles = @($all | Where-Object { $_.Name -eq $packageName -and $_.Ext -match 'bundle$' -and $_.Ext -notmatch '^\.e' } | Sort-Object Version -Descending)
        if ($Version) { $bundles = @($bundles | Where-Object { $_.Version -eq [version] $Version }) }
        if ($bundles.Count -eq 0) { throw "no $packageName bundle$(if ($Version) { " $Version" }) at Microsoft" }
        $bundle = $bundles[0]
        Write-Ok "1. Bundle: $($bundle.Moniker)"

        # Into a new folder first: the old one stays until everything is checked.
        $staging = Join-Path $OutputFolder ('.new_{0}' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $staging -Force | Out-Null
        $bundlePath = Save-Package -Package $bundle -Folder $staging
        Write-Ok '2. Check: size, SHA-256 and Microsoft signature match'

        $frameworks = @()
        foreach ($dep in @(Get-BundleDependency -BundlePath $bundlePath)) {
            $match = @($all | Where-Object { $_.Name -eq $dep.Name -and $_.Arch -eq $Architecture -and $_.Version -ge $dep.MinVersion -and $_.Ext -notmatch '^\.e' } | Sort-Object Version -Descending) | Select-Object -First 1
            if (-not $match) { Write-Warn "$($dep.Name) >= $($dep.MinVersion) ($Architecture) not offered - the host must already have it"; continue }
            [void] (Save-Package -Package $match -Folder $staging)
            $frameworks += $match.Moniker
        }
        Write-Ok "3. Frameworks: $(if ($frameworks) { $frameworks -join ', ' } else { 'none needed' })"

        $final = Join-Path $OutputFolder $packageName
        if (Test-Path -LiteralPath $final) { Remove-Item -LiteralPath $final -Recurse -Force }
        Rename-Item -LiteralPath $staging -NewName $packageName
        $staging = $null
        Write-Ok "4. Ready: $final"
        $row.Bundle = [string] $bundle.Version; $row.Frameworks = $frameworks.Count; $row.Folder = $final; $row.Result = 'ok'
    } catch {
        Write-Bad "${name}: $($_.Exception.Message)"
        $exitCode = 1
    } finally {
        if ($staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $results.Add($row)
}

foreach ($target in @($CopyTo)) {
    if (-not $target) { continue }
    Write-Step "Copy to $target"
    $dest = "\\$target\c$\IT\AppWatchdog\Packages"
    foreach ($row in @($results | Where-Object { $_.Result -eq 'ok' })) {
        $leaf = Split-Path $row.Folder -Leaf
        & robocopy.exe $row.Folder (Join-Path $dest $leaf) /MIR /R:2 /W:5 /NFL /NDL /NJH /NJS /NP | Out-Null
        if ($LASTEXITCODE -ge 8) { Write-Bad "$target ${leaf}: copy failed (robocopy $LASTEXITCODE)"; $exitCode = 1 } else { Write-Ok "$target - $leaf" }
    }
}

Write-Host ''
$results | Format-Table App, Bundle, Frameworks, Result, Folder -AutoSize | Out-String -Width 200 | Write-Host
exit $exitCode
