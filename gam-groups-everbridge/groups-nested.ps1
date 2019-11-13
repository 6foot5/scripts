# To run: powershell -noexit -file "[full-path-to-script]"

#-----------------------------------------------------------
# Begin variable definitions for file operations 
#-----------------------------------------------------------

$rootFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$archivePath = $rootFolder + "\archive"
$rightNow = get-date -format filedatetime

$csvHeader = "primaryEmail,name.givenName,name.familyName,suspended,suspensionReason,externalIds,externalIds.0.value,externalIds.0.type,Groups"

$tempFilePathGroups = $rootFolder + "\groupsync-" + $rightNow + ".csv"
$tempFilePathAllGroups = $rootFolder + "\all-groups-" + $rightNow + ".csv"
$tempFilePathAllUsers = $rootFolder + "\all-users-" + $rightNow + ".csv"

$permFileNameGroups = "groupsync.csv"

$permFilePathGroups = $rootFolder + "\" + $permFileNameGroups

#-----------------------------------------------------------
# End variable definitions for file operations
#-----------------------------------------------------------


#-----------------------------------------------------------
# If the archive folder does not yet exist, create it
#-----------------------------------------------------------

if (!(test-path $archivePath))
{
      New-Item -ItemType Directory -Force -Path $archivePath
}


#-----------------------------------------------------------
# Create a fresh group memebership file to be populated by this script
#-----------------------------------------------------------


add-content -path $tempFilePathGroups -value $csvHeader


#-----------------------------------------------------------
# GAM commands to dump user and group info into two CSV files
# ...then load those files into PS objects for processing
#-----------------------------------------------------------

gam print users firstname lastname suspended externalIds > $tempFilePathAllUsers

gam print group-members fields group,email recursive > $tempFilePathAllGroups

$AllUsers = Import-Csv -Path $tempFilePathAllUsers
$AllGroups = Import-Csv -Path $tempFilePathAllGroups


#-----------------------------------------------------------------
# For each user returned by GAM, store all properties in variables
# These variables will be used to create each line of the CSV
#-----------------------------------------------------------------

foreach ($user in $AllUsers) {
    
    $newLine = ""

    $userEmail = ""
    $userFirstName = ""
    $userLastName = ""
    $userSuspended = ""
    $userSuspendedReason = ""
    $userExternalIDs = ""
    $userExternalIDsValue = ""
    $userExternalIDsType = ""

    $properties = $user | Get-Member -MemberType Properties

    for($i=0; $i -lt $properties.Count; $i++)
    {
        $column = $properties[$i]
        $columnvalue = $user | Select -ExpandProperty $column.Name

        switch ($column.Name) {
           "primaryEmail" {
                $userEmail = $columnvalue; 
                break
           }
           "name.givenName" {
                $userFirstName = $columnvalue; 
                break
           }
           "name.familyName" {
                $userLastName = $columnvalue; 
                break
           }
           "suspended" {
                $userSuspended = $columnvalue; 
                break
           }
           "suspensionReason" {
                $userSuspendedReason = $columnvalue; 
                break
           }
           "externalIds" {
                $userExternalIDs = $columnvalue; 
                break
           }
           "externalIds.0.value" {
                $userExternalIDsValue = $columnvalue; 
                break
           }
           "externalIds.0.type" {
                $userExternalIDsType = $columnvalue; 
                break
           }
           default {
                break
           }
       }
    }

    Write-Output "----------------------------------"
    Write-Output "Processing User:"
    Write-Output $userEmail
    Write-Output "----------------------------------"


#-----------------------------------------------------------------
# For current user, find query for all group memberships
#-----------------------------------------------------------------

    $AllGroups | Where-Object -Property email -eq $userEmail -OutVariable thisUsersGroups


#-----------------------------------------------------------------
# For each group found, add it to the group list variable
#-----------------------------------------------------------------

    $groupList = ""

    foreach ($group in $thisUsersGroups) {
        $groupList = $groupList + " " + $group.group
    }

    $groupList = $groupList.TrimStart(" ")


#-----------------------------------------------------------------
# Concatenate all user properties, plus the group list
# Store in newLine variable to write to CSV (matching csvHeader)
#-----------------------------------------------------------------

    $newLine = $userEmail + "," + $userFirstName + "," + $userLastName + "," + $userSuspended + "," + $userSuspendedReason + "," +  $userExternalIDs + "," + $userExternalIDsValue + "," + $userExternalIDsType + "," + $groupList

    $newLine | add-content -path $tempFilePathGroups

} 

#-----------------------------------------------------------------
# End of loop; writing to group membership file is complete
#-----------------------------------------------------------------


#-----------------------------------------------------------------
# CLEANUP
# Remove prior "permanent" group membership file, if it still exists
# Move working / temp files to archive folder
# Rename the group membership file in place (with permanent name)
# ...This Permanent group membership file will be referenced by automated data sync routine
#-----------------------------------------------------------------

if (test-path $permFilePathGroups) {
    remove-item -path $permFilePathGroups
}

if (test-path $tempFilePathGroups) {
    copy-item $tempFilePathGroups -destination $archivePath
    rename-item -path $tempFilePathGroups -NewName $permFilePathGroups
}

if (test-path $tempFilePathAllGroups) {
    move-item $tempFilePathAllGroups -destination $archivePath
}

if (test-path $tempFilePathAllUsers) {
    move-item $tempFilePathAllUsers -destination $archivePath
}


