$drive = "$env:OneDriveCommercial\RDM\Back-up"
$archive = "$env:OneDriveCommercial\RDM\Back-up\Archief"
$datum = Get-Date -Format "d-M-y"
$time = Get-Date -Format "HH:mm:ss"


# Copy the file to the archive with the current date
Copy-Item -Path "$drive\RDM.rdm" -Destination "$archive\RDM-$datum.rdm" -Force