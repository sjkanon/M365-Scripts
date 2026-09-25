<#
.SYNOPSIS
    Pin a file to the Windows taskbar from a script.

.DESCRIPTION
    Windows exposes no supported way to pin an item to the taskbar unattended. This
    borrows the hidden Windows.taskbarpin verb: it copies that verb's command handler
    into a temporary HKCU class, invokes it on the target through the Shell COM
    object, and removes the temporary key again. The companion to add-lock.ps1, which
    creates the shortcut this pins.

.PARAMETER Target
    Full path of the item to pin. The script warns and stops when it does not exist.
#>

param (
    [parameter(Mandatory=$True, HelpMessage="Target item to pin")]
    [ValidateNotNullOrEmpty()]
    [string] $Target
)
if (!(Test-Path $Target)) {
    Write-Warning "$Target does not exist"
    break
}

$KeyPath1  = "HKCU:\SOFTWARE\Classes"
$KeyPath2  = "*"
$KeyPath3  = "shell"
$KeyPath4  = "{:}"
$ValueName = "ExplorerCommandHandler"
$ValueData =
    (Get-ItemProperty `
        ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\" + `
            "CommandStore\shell\Windows.taskbarpin")
    ).ExplorerCommandHandler

$Key2 = (Get-Item $KeyPath1).OpenSubKey($KeyPath2, $true)
$Key3 = $Key2.CreateSubKey($KeyPath3, $true)
$Key4 = $Key3.CreateSubKey($KeyPath4, $true)
$Key4.SetValue($ValueName, $ValueData)

$Shell = New-Object -ComObject "Shell.Application"
$Folder = $Shell.Namespace((Get-Item $Target).DirectoryName)
$Item = $Folder.ParseName((Get-Item $Target).Name)
$Item.InvokeVerb("{:}")

$Key3.DeleteSubKey($KeyPath4)
if ($Key3.SubKeyCount -eq 0 -and $Key3.ValueCount -eq 0) {
    $Key2.DeleteSubKey($KeyPath3)
}
./add-shortcut-lock.ps1 