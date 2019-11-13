# to take command line paramaters (e.g. -sourceUser **** -sourceFolder ***** etc etc
#param([string]$sourceUser,[string]$sourceFolderID,[string]$destinationTeamDriveID)

$sourceUser = Read-Host -Prompt 'Enter the source user account'
$sourceFolderID = Read-Host -Prompt 'Enter the source folder ID'
$destinationTeamDriveID = Read-Host -Prompt 'Enter the destination Shared Drive ID'

#gam user $sourceUser move drivefile $sourceFolderID teamdriveparentid $destinationTeamDriveID
gam user $sourceUser move drivefile $sourceFolderID parentid $destinationTeamDriveID
