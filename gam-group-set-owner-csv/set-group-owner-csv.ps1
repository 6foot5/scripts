<#

This script will take as input a CSV file that lives next to this script in the same directory

- column header "groupid", all rows are group email addresses

set the newOwner variable below to match who the new manager of the group should be

set the groupListFile to match the name of the CSV you've put in this directory

#>

$newOwner = "jknighton@ashevillenc.gov"

$rootFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$groupListFile = $rootFolder + "\afd.csv"

$AllGroups = Import-Csv -Path $groupListFile

foreach ($group in $AllGroups) 
{
    $currentGroupID = $group.groupid
    gam update group $currentGroupID add manager $newOwner 
}

