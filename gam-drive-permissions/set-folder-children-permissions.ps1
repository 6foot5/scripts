$currentOwner = "[X]" # Email address of current folder owner
$newOwner = "[X]" # Email address of new owner
$parentFolder = "[X]" # Parent folder ID (all child objects will get new owner)

$rightNow = get-date -format filedatetime
$rootFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$tempFilePathFiles = $rootFolder + "\filelist-" + $rightNow + ".csv"

gam user $currentOwner show filelist select $parentFolder allfields > $tempFilePathFiles

$AllFiles = Import-Csv -Path $tempFilePathFiles

foreach ($file in $AllFiles) 
{
    $currentFileID = $file.id
    gam user $currentOwner add drivefileacl $currentFileID user $newOwner role owner
}

