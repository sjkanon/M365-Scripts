<#PSScriptInfo

.VERSION 3.5.1

.GUID ebf446a3-3362-4774-83c0-b7299410b63f

.AUTHOR Michael Niehaus

.COMPANYNAME Microsoft

.COPYRIGHT

.TAGS Windows AutoPilot

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
Version 1.0:  Original published version.
Version 1.1:  Added -Append switch.
Version 1.2:  Added -Credential switch.
Version 1.3:  Added -Partner switch.
Version 1.4:  Switched from Get-WMIObject to Get-CimInstance.
Version 1.5:  Added -GroupTag parameter.
Version 1.6:  Bumped version number (no other change).
Version 2.0:  Added -Online parameter.
Version 2.1:  Bug fix.
Version 2.3:  Updated comments.
Version 2.4:  Updated "online" import logic to wait for the device to sync, added new parameter.
Version 2.5:  Added AssignedUser for Intune importing, and AssignedComputerName for online Intune importing.
Version 2.6:  Added support for app-based authentication via Connect-MSGraphApp.
Version 2.7:  Added new Reboot option for use with -Online -Assign.
Version 2.8:  Fixed up parameter sets.
Version 2.9:  Fixed typo installing AzureAD module.
Version 3.0:  Fixed typo for app-based auth, added logic to explicitly install NuGet (silently).
Version 3.2:  Fixed logic to explicitly install NuGet (silently).
Version 3.3:  Added more logging and error handling for group membership.
Version 3.4:  Added logic to verify that devices were added successfully.  Fixed a bug that could cause all Autopilot devices to be added to the specified AAD group.
Version 3.5:  Added logic to display the serial number of the gathered device.
Version 3.5.1 (M365-Scripts): The -Online part talks to Microsoft Graph directly (Microsoft.Graph.Authentication,
              Invoke-MgGraphRequest) instead of the retired AzureAD and Microsoft.Graph.Intune modules and
              WindowsAutopilotIntune. Delegated sign-in by default (-DeviceCode for OOBE), app-only with
              -AppId plus -AppSecret or -CertificateThumbprint. Fixed the import/sync loops reporting the last
              device for every device, the CSV header (duplicate "Hardware Hash" column), and a missing
              assignment status crashing -Assign.
#>

<#
.SYNOPSIS
Retrieves the Windows AutoPilot deployment details from one or more computers

MIT LICENSE

Copyright (c) 2020 Microsoft

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.DESCRIPTION
This script uses WMI to retrieve properties needed for a customer to register a device with Windows Autopilot.  Note that it is normal for the resulting CSV file to not collect a Windows Product ID (PKID) value since this is not required to register a device.  Only the serial number and hardware hash will be populated.

With -Online the devices are imported into Windows Autopilot through Microsoft Graph
(deviceManagement/importedWindowsAutopilotDeviceIdentities, beta - the same endpoints the
WindowsAutopilotIntune module used). Only Microsoft.Graph.Authentication is needed; it is
installed for the current user when missing.

This script stays standalone and runs on Windows PowerShell 5.1: it is copied to a USB stick
and started from OOBE (Shift+F10) or by scripts\Deployment\start.bat, so it does not load the
repository's scripts\Startup\Connect-M365.ps1. Sign-in follows the same rules:
  Delegated (default)  You sign in as an Intune admin. Use -DeviceCode when no browser can open
                       (OOBE). Scopes: DeviceManagementServiceConfig.ReadWrite.All, plus
                       GroupMember.ReadWrite.All and Device.Read.All with -AddToGroup.
  App-only (option)    -AppId with -AppSecret or -CertificateThumbprint, and -TenantId. The app
                       needs the same permissions as application permissions.
.PARAMETER Name
The names of the computers.  These can be provided via the pipeline (property name Name or one of the available aliases, DNSHostName, ComputerName, and Computer).
.PARAMETER OutputFile
The name of the CSV file to be created with the details for the computers.  If not specified, the details will be returned to the PowerShell
pipeline.
.PARAMETER Append
Switch to specify that new computer details should be appended to the specified output file, instead of overwriting the existing file.
.PARAMETER Credential
Credentials that should be used when connecting to a remote computer (not supported when gathering details from the local computer).
.PARAMETER Partner
Switch to specify that the created CSV file should use the schema for Partner Center (using serial number, make, and model).
.PARAMETER GroupTag
An optional tag value that should be included in a CSV file that is intended to be uploaded via Intune (not supported by Partner Center or Microsoft Store for Business).
.PARAMETER AssignedUser
An optional value specifying the UPN of the user to be assigned to the device.  This can only be specified for Intune (not supported by Partner Center or Microsoft Store for Business).
.PARAMETER Online
Add computers to Windows Autopilot via the Intune Graph API
.PARAMETER TenantId
Tenant ID or domain. Required for app-only; optional for delegated (defaults to the tenant of the account you sign in with).
.PARAMETER AppId
App registration (client ID) for app-only sign-in, with -AppSecret or -CertificateThumbprint.
.PARAMETER AppSecret
Client secret for -AppId.  Prefer -CertificateThumbprint.
.PARAMETER CertificateThumbprint
Certificate thumbprint for -AppId (certificate with private key in CurrentUser\My or LocalMachine\My).
.PARAMETER DeviceCode
Delegated sign-in with a device code instead of a browser window (useful in OOBE).
.PARAMETER AssignedComputerName
An optional value specifying the computer name to be assigned to the device.  This can only be specified with the -Online switch and only works with AAD join scenarios.
.PARAMETER AddToGroup
Specifies the name of the Azure AD group that the new device should be added to.
.PARAMETER Assign
Wait for the Autopilot profile assignment.  (This can take a while for dynamic groups.)
.PARAMETER Reboot
Reboot the device after the Autopilot profile has been assigned (necessary to download the profile and apply the computer name, if specified).
.EXAMPLE
.\Get-WindowsAutoPilotInfo.ps1 -ComputerName MYCOMPUTER -OutputFile .\MyComputer.csv
.EXAMPLE
.\Get-WindowsAutoPilotInfo.ps1 -ComputerName MYCOMPUTER -OutputFile .\MyComputer.csv -GroupTag Kiosk
.EXAMPLE
.\Get-WindowsAutoPilotInfo.ps1 -ComputerName MYCOMPUTER -OutputFile .\MyComputer.csv -GroupTag Kiosk -AssignedUser JohnDoe@contoso.com
.EXAMPLE
.\Get-WindowsAutoPilotInfo.ps1 -ComputerName MYCOMPUTER -OutputFile .\MyComputer.csv -Append
.EXAMPLE
.\Get-WindowsAutoPilotInfo.ps1 -ComputerName MYCOMPUTER1,MYCOMPUTER2 -OutputFile .\MyComputers.csv
.EXAMPLE
Get-ADComputer -Filter * | .\GetWindowsAutoPilotInfo.ps1 -OutputFile .\MyComputers.csv
.EXAMPLE
Get-CMCollectionMember -CollectionName "All Systems" | .\GetWindowsAutoPilotInfo.ps1 -OutputFile .\MyComputers.csv
.EXAMPLE
.\Get-WindowsAutoPilotInfo.ps1 -ComputerName MYCOMPUTER1,MYCOMPUTER2 -OutputFile .\MyComputers.csv -Partner
.EXAMPLE
.\GetWindowsAutoPilotInfo.ps1 -Online
.EXAMPLE
# From OOBE (Shift+F10): device code sign-in, add to a group, wait for the profile and reboot
.\Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode -GroupTag Corporate -AddToGroup "Autopilot Devices" -Assign -Reboot
.EXAMPLE
# App-only with a certificate
.\Get-WindowsAutoPilotInfo.ps1 -Online -TenantId contoso.onmicrosoft.com -AppId 00000000-0000-0000-0000-000000000000 -CertificateThumbprint AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA

#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
	[Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=0)][alias("DNSHostName","ComputerName","Computer")] [String[]] $Name = @("localhost"),
	[Parameter(Mandatory=$False)] [String] $OutputFile = "",
	[Parameter(Mandatory=$False)] [String] $GroupTag = "",
	[Parameter(Mandatory=$False)] [String] $AssignedUser = "",
	[Parameter(Mandatory=$False)] [Switch] $Append = $false,
	[Parameter(Mandatory=$False)] [System.Management.Automation.PSCredential] $Credential = $null,
	[Parameter(Mandatory=$False)] [Switch] $Partner = $false,
	[Parameter(Mandatory=$False)] [Switch] $Force = $false,
	[Parameter(Mandatory=$True,ParameterSetName = 'Online')] [Switch] $Online = $false,
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $TenantId = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AppId = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AppSecret = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $CertificateThumbprint = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [Switch] $DeviceCode = $false,
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AddToGroup = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AssignedComputerName = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [Switch] $Assign = $false,
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [Switch] $Reboot = $false
)

Begin
{
	# Initialize empty list
	$computers = @()
	$graphConnectedHere = $false

	# Autopilot import endpoints are used on beta, as the WindowsAutopilotIntune module did
	# (deploymentProfileAssignmentStatus is only exposed there).
	$apiBeta = "https://graph.microsoft.com/beta/deviceManagement"
	$apiV1 = "https://graph.microsoft.com/v1.0"

	function Get-AutopilotDeviceById {
		param([string] $Id)
		if (-not $Id) { return $null }
		try {
			return Invoke-MgGraphRequest -Method GET -Uri "$apiBeta/windowsAutopilotDeviceIdentities/$Id" -ErrorAction Stop
		} catch {
			# 404 until the imported device has synced into Autopilot
			return $null
		}
	}

	# If online, make sure we are able to authenticate
	if ($Online) {

		# Windows PowerShell 5.1 on a fresh device may still default to TLS 1.0/1.1
		[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

		# Get NuGet
		$provider = Get-PackageProvider NuGet -ListAvailable -ErrorAction Ignore
		if (-not $provider) {
			Write-Host "Installing provider NuGet"
			Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
		}

		# Microsoft.Graph.Authentication is the only module needed (Invoke-MgGraphRequest)
		$module = Import-Module Microsoft.Graph.Authentication -PassThru -ErrorAction Ignore
		if (-not $module) {
			Write-Host "Installing module Microsoft.Graph.Authentication"
			Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
			Import-Module Microsoft.Graph.Authentication -Scope Global
		}

		# Connect
		$connect = @{ NoWelcome = $true; ErrorAction = 'Stop' }
		if ($TenantId -ne "") { $connect['TenantId'] = $TenantId }
		if ($AppId -ne "")
		{
			if ($TenantId -eq "") { throw "App-only sign-in (-AppId) needs -TenantId." }
			if ($CertificateThumbprint -ne "") {
				$connect['ClientId'] = $AppId
				$connect['CertificateThumbprint'] = $CertificateThumbprint
			}
			elseif ($AppSecret -ne "") {
				$secure = ConvertTo-SecureString $AppSecret -AsPlainText -Force
				$connect['ClientSecretCredential'] = New-Object System.Management.Automation.PSCredential($AppId, $secure)
			}
			else { throw "-AppId needs -CertificateThumbprint or -AppSecret." }
			Connect-MgGraph @connect
			$graphConnectedHere = $true
			Write-Host "Connected to Intune tenant $TenantId using app-only authentication"
		}
		else {
			$scopes = @('DeviceManagementServiceConfig.ReadWrite.All')
			if ($AddToGroup) { $scopes += 'GroupMember.ReadWrite.All', 'Device.Read.All' }
			$ctx = Get-MgContext
			$missing = @($scopes | Where-Object { -not $ctx -or $ctx.AuthType -ne 'Delegated' -or $ctx.Scopes -notcontains $_ })
			$tenantOk = ($TenantId -eq "") -or ($ctx -and ($ctx.TenantId -eq $TenantId -or $ctx.Account -like "*@$TenantId"))
			if ($missing.Count -gt 0 -or -not $tenantOk) {
				$connect['Scopes'] = $scopes
				if ($DeviceCode) { $connect['UseDeviceCode'] = $true }
				Connect-MgGraph @connect
				$graphConnectedHere = $true
			}
			$ctx = Get-MgContext
			Write-Host "Connected to Intune tenant $($ctx.TenantId) as $($ctx.Account)"
		}

		# Force the output to a file
		if ($OutputFile -eq "")
		{
			$OutputFile = "$($env:TEMP)\autopilot.csv"
		}
	}
}

Process
{
	foreach ($comp in $Name)
	{
		$bad = $false

		# Get a CIM session
		if ($comp -eq "localhost") {
			$session = New-CimSession
		}
		else
		{
			$session = New-CimSession -ComputerName $comp -Credential $Credential
		}

		# Get the common properties.
		Write-Verbose "Checking $comp"
		$serial = (Get-CimInstance -CimSession $session -Class Win32_BIOS).SerialNumber

		# Get the hash (if available)
		$devDetail = (Get-CimInstance -CimSession $session -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'")
		if ($devDetail -and (-not $Force))
		{
			$hash = $devDetail.DeviceHardwareData
		}
		else
		{
			$bad = $true
			$hash = ""
		}

		# If the hash isn't available, get the make and model
		if ($bad -or $Force)
		{
			$cs = Get-CimInstance -CimSession $session -Class Win32_ComputerSystem
			$make = $cs.Manufacturer.Trim()
			$model = $cs.Model.Trim()
			if ($Partner)
			{
				$bad = $false
			}
		}
		else
		{
			$make = ""
			$model = ""
		}

		# Getting the PKID is generally problematic for anyone other than OEMs, so let's skip it here
		$product = ""

		# Depending on the format requested, create the necessary object
		if ($Partner)
		{
			# Create a pipeline object
			$c = New-Object psobject -Property @{
				"Device Serial Number" = $serial
				"Windows Product ID" = $product
				"Hardware Hash" = $hash
				"Manufacturer name" = $make
				"Device model" = $model
			}
			# From spec:
			#	"Manufacturer Name" = $make
			#	"Device Name" = $model

		}
		else
		{
			# Create a pipeline object
			$c = New-Object psobject -Property @{
				"Device Serial Number" = $serial
				"Windows Product ID" = $product
				"Hardware Hash" = $hash
				"Manufacturer name" = $make
				"Device model" = $model
			}

			if ($GroupTag -ne "")
			{
				Add-Member -InputObject $c -NotePropertyName "Group Tag" -NotePropertyValue $GroupTag
			}
			if ($AssignedUser -ne "")
			{
				Add-Member -InputObject $c -NotePropertyName "Assigned User" -NotePropertyValue $AssignedUser
			}
		}

		# Write the object to the pipeline or array
		if ($bad)
		{
			# Report an error when the hash isn't available
			Write-Error -Message "Unable to retrieve device hardware data (hash) from computer $comp" -Category DeviceError
		}
		elseif ($OutputFile -eq "")
		{
			$c
		}
		else
		{
			$computers += $c
			Write-Host "Gathered details for device with serial number: $serial"
		}

		Remove-CimSession $session
	}
}

End
{
	if ($OutputFile -ne "")
	{
		if ($Append)
		{
			if (Test-Path $OutputFile)
			{
				$computers += Import-CSV -Path $OutputFile
			}
		}
		# Intune's CSV import accepts only these columns (in this order); the earlier copy added
		# make/model and a second "Hardware Hash" column, which Select-Object rejects.
		if ($Partner)
		{
			$computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Manufacturer name", "Device model" | ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $OutputFile
		}
		elseif ($AssignedUser -ne "")
		{
			$computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag", "Assigned User" | ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $OutputFile
		}
		elseif ($GroupTag -ne "")
		{
			$computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" | ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $OutputFile
		}
		else
		{
			$computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash" | ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $OutputFile
		}
	}
	if ($Online)
	{
		# Add the devices
		$importStart = Get-Date
		$imported = @()
		foreach ($comp in $computers) {
			$body = @{
				'@odata.type'             = '#microsoft.graph.importedWindowsAutopilotDeviceIdentity'
				orderIdentifier           = [string]$comp.'Group Tag'
				groupTag                  = [string]$comp.'Group Tag'
				serialNumber              = [string]$comp.'Device Serial Number'
				productKey                = ''
				hardwareIdentifier        = [string]$comp.'Hardware Hash'
				assignedUserPrincipalName = [string]$comp.'Assigned User'
				state = @{
					'@odata.type'        = 'microsoft.graph.importedWindowsAutopilotDeviceIdentityState'
					deviceImportStatus   = 'pending'
					deviceRegistrationId = ''
					deviceErrorCode      = 0
					deviceErrorName      = ''
				}
			} | ConvertTo-Json -Depth 5
			$imported += Invoke-MgGraphRequest -Method POST -Uri "$apiBeta/importedWindowsAutopilotDeviceIdentities" -Body $body -ContentType 'application/json' -ErrorAction Stop
		}

		# Wait until the devices have been imported
		$processingCount = 1
		while ($processingCount -gt 0)
		{
			$current = @()
			$processingCount = 0
			foreach ($item in $imported) {
				$device = Invoke-MgGraphRequest -Method GET -Uri "$apiBeta/importedWindowsAutopilotDeviceIdentities/$($item.id)" -ErrorAction Stop
				if ($device.state.deviceImportStatus -in @('unknown', 'pending')) {
					$processingCount = $processingCount + 1
				}
				$current += $device
			}
			$deviceCount = $imported.Count
			Write-Host "Waiting for $processingCount of $deviceCount to be imported"
			if ($processingCount -gt 0){
				Start-Sleep 30
			}
		}
		$importDuration = (Get-Date) - $importStart
		$importSeconds = [Math]::Ceiling($importDuration.TotalSeconds)
		$successCount = 0
		foreach ($device in $current) {
			Write-Host "$($device.serialNumber): $($device.state.deviceImportStatus) $($device.state.deviceErrorCode) $($device.state.deviceErrorName)"
			if ($device.state.deviceImportStatus -eq "complete") {
				$successCount = $successCount + 1
			}
		}
		Write-Host "$successCount devices imported successfully.  Elapsed time to complete import: $importSeconds seconds"

		# Wait until the devices can be found in Intune (should sync automatically)
		$syncStart = Get-Date
		$completed = @($current | Where-Object { $_.state.deviceImportStatus -eq "complete" })
		$processingCount = 1
		while ($processingCount -gt 0)
		{
			$autopilotDevices = @()
			$processingCount = 0
			foreach ($item in $completed) {
				$device = Get-AutopilotDeviceById -Id $item.state.deviceRegistrationId
				if (-not $device) {
					$processingCount = $processingCount + 1
				}
				else {
					$autopilotDevices += $device
				}
			}
			$deviceCount = $completed.Count
			Write-Host "Waiting for $processingCount of $deviceCount to be synced"
			if ($processingCount -gt 0){
				Start-Sleep 30
			}
		}
		$syncDuration = (Get-Date) - $syncStart
		$syncSeconds = [Math]::Ceiling($syncDuration.TotalSeconds)
		Write-Host "All devices synced.  Elapsed time to complete sync: $syncSeconds seconds"

		# Add the device to the specified AAD group
		if ($AddToGroup)
		{
			$groupFilter = [uri]::EscapeDataString("displayName eq '$($AddToGroup -replace "'", "''")'")
			$aadGroup = @((Invoke-MgGraphRequest -Method GET -Uri "$apiV1/groups?`$filter=$groupFilter&`$select=id,displayName" -ErrorAction Stop).value) | Select-Object -First 1
			if ($aadGroup)
			{
				foreach ($apDevice in $autopilotDevices) {
					# The Entra device object can lag a little behind the Autopilot sync.
					$aadDevice = $null
					for ($i = 1; $i -le 10 -and -not $aadDevice; $i++) {
						$deviceFilter = [uri]::EscapeDataString("deviceId eq '$($apDevice.azureActiveDirectoryDeviceId)'")
						$aadDevice = @((Invoke-MgGraphRequest -Method GET -Uri "$apiV1/devices?`$filter=$deviceFilter&`$select=id,deviceId" -ErrorAction Stop).value) | Select-Object -First 1
						if (-not $aadDevice -and $i -lt 10) { Start-Sleep 30 }
					}
					if ($aadDevice) {
						Write-Host "Adding device $($apDevice.serialNumber) to group $AddToGroup"
						try {
							$ref = @{ '@odata.id' = "$apiV1/directoryObjects/$($aadDevice.id)" } | ConvertTo-Json
							Invoke-MgGraphRequest -Method POST -Uri "$apiV1/groups/$($aadGroup.id)/members/`$ref" -Body $ref -ContentType 'application/json' -ErrorAction Stop | Out-Null
							Write-Host "Added device $($apDevice.serialNumber) to group '$AddToGroup' ($($aadGroup.id))"
						} catch {
							Write-Error "Unable to add device $($apDevice.serialNumber) to group '$AddToGroup': $($_.Exception.Message)"
						}
					}
					else {
						Write-Error "Unable to find Azure AD device with ID $($apDevice.azureActiveDirectoryDeviceId)"
					}
				}
			}
			else {
				Write-Error "Unable to find group $AddToGroup"
			}
		}

		# Assign the computer name
		if ($AssignedComputerName -ne "")
		{
			foreach ($apDevice in $autopilotDevices) {
				$props = @{ displayName = $AssignedComputerName } | ConvertTo-Json
				Invoke-MgGraphRequest -Method POST -Uri "$apiBeta/windowsAutopilotDeviceIdentities/$($apDevice.id)/updateDeviceProperties" -Body $props -ContentType 'application/json' -ErrorAction Stop | Out-Null
			}
		}

		# Wait for assignment (if specified)
		if ($Assign)
		{
			$assignStart = Get-Date
			$processingCount = 1
			while ($processingCount -gt 0)
			{
				$processingCount = 0
				foreach ($apDevice in $autopilotDevices) {
					$device = Get-AutopilotDeviceById -Id $apDevice.id
					if (-not ([string]$device.deploymentProfileAssignmentStatus).StartsWith("assigned")) {
						$processingCount = $processingCount + 1
					}
				}
				$deviceCount = $autopilotDevices.Count
				Write-Host "Waiting for $processingCount of $deviceCount to be assigned"
				if ($processingCount -gt 0){
					Start-Sleep 30
				}
			}
			$assignDuration = (Get-Date) - $assignStart
			$assignSeconds = [Math]::Ceiling($assignDuration.TotalSeconds)
			Write-Host "Profiles assigned to all devices.  Elapsed time to complete assignment: $assignSeconds seconds"
			if ($Reboot)
			{
				if ($graphConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
				Restart-Computer -Force
			}
		}

		if ($graphConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
	}
}
