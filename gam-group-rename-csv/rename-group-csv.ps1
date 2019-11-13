<#

This script will take as input a CSV file that lives next to this script in the same directory

- column headers "groupemail" and "groupnamenew", all rows are group email addresses and new group names

set the groupListFile to match the name of the CSV you've put in this directory

#>

$rootFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$groupListFile = $rootFolder + "\apd-groups.csv"

$AllGroups = Import-Csv -Path $groupListFile

foreach ($group in $AllGroups) {

    $currentGroupID = $group.groupemail
    $newGroupName = $group.groupnamenew

    gam update group $currentGroupID name "$newGroupName"
}

