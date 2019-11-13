
#-----------------------------------------------------------
# Begin variable definitions for file operations 
#-----------------------------------------------------------

$driveTargetUser = "" # Enter email address of target user account


$rootFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$scratchPath = $rootFolder + "\scratch"
$attachmentPath = $rootFolder + "\attachments"
$appLog = $scratchPath + "\log.txt"
$tmpSolution = $scratchPath + "\this-solution.html"
$scratchKBPath = $scratchPath + "\kb-with-path.csv"
$scratchSolutionsPath = $scratchPath + "\solutions-with-attachments.csv"
$driveURLPrefix = "https://drive.google.com/file/d/"

# Google sheets will automatically detect tab delimiter (commas not possible with data given)
$defaultDelimiter = "`t" 

#-----------------------------------------------------------
# If the scratch folder does not yet exist, create it
# Remove temp files from previous iterations
#-----------------------------------------------------------

if (!(test-path $scratchPath))
{
      New-Item -ItemType Directory -Force -Path $scratchPath
}

if (test-path $scratchKBPath) {
    remove-item -path $scratchKBPath
}

if (test-path $scratchSolutionsPath) {
    remove-item -path $scratchSolutionsPath
}

if (test-path $appLog) {
    remove-item -path $appLog
}


#-----------------------------------------------------------
# SQL database connection setup 
# (will authenticate with local Windows session credentials)
#-----------------------------------------------------------

$SQLServer = "" # Enter target database server name
$SQLDBName = "" # Enter target database name
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"


#-----------------------------------------------------------
# SQL query - fill KB Type dataset 
# (Solution categories)
#-----------------------------------------------------------

$SqlQueryKB = "SELECT * from KBTYPES order by KB_TYPEID;"
$SqlCmdKB = New-Object System.Data.SqlClient.SqlCommand
$SqlCmdKB.CommandText = $SqlQueryKB
$SqlCmdKB.Connection = $SqlConnection
$SqlAdapterKB = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapterKB.SelectCommand = $SqlCmdKB
$DataSetKB = New-Object System.Data.DataSet
$SqlAdapterKB.Fill($DataSetKB)

#-----------------------------------------------------------
# SQL query - fill Solutions dataset 
# (Solution data, without attachments)
#-----------------------------------------------------------

$SqlQueryS = "SELECT s.SOLUTIONID, s.SUMMARY, s.DETAIL, s.KB_TYPEID, s.LASTUPDATED
    FROM SOLUTION s  
    order by s.SOLUTIONID;"
$SqlCmdS = New-Object System.Data.SqlClient.SqlCommand
$SqlCmdS.CommandText = $SqlQueryS
$SqlCmdS.Connection = $SqlConnection
$SqlAdapterS = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapterS.SelectCommand = $SqlCmdS
$DataSetSolutions = New-Object System.Data.DataSet
$SqlAdapterS.Fill($DataSetSolutions)

#-----------------------------------------------------------
# SQL query - fill Attachments dataset, joined on solutions
# Gives all attachment info, joined with solution meta data
#-----------------------------------------------------------

$SqlQuerySA = "SELECT s.SOLUTIONID, s.SUMMARY, s.DETAIL, s.KB_TYPEID, s.LASTUPDATED, 
    a.AttachmentOwnerId, a.LogicalName, a.LastModifiedBy, a.LastModifiedDate 
    FROM SOLUTION s 
    LEFT JOIN Attachment a ON a.AttachmentOwnerId = s.SOLUTIONID 
    order by s.SOLUTIONID;"
$SqlCmdSA = New-Object System.Data.SqlClient.SqlCommand
$SqlCmdSA.CommandText = $SqlQuerySA
$SqlCmdSA.Connection = $SqlConnection
$SqlAdapterSA = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapterSA.SelectCommand = $SqlCmdSA
$DataSetAttachments = New-Object System.Data.DataSet
$SqlAdapterSA.Fill($DataSetAttachments)

#-----------------------------------------------------------
# Set query results to PS objects for handling in loops below
#-----------------------------------------------------------

$attachmentSet = $DataSetAttachments.Tables[0]
$solutionSet = $DataSetSolutions.Tables[0]
$kbSet = $DataSetKB.Tables[0]

#-----------------------------------------------------------
# Build scratch file to hold flattened directory path info
# for each KB Type ID. These will become Google Drive folders
# (may revisit and build dataset in PS object instead of CSV)
# This loop should yield one line for each solution category
#-----------------------------------------------------------

$fileHeaderKB = "KB_ID|KB_NAME|KB_PATH"
add-content -path $scratchKBPath -value $fileHeaderKB

foreach ($kb in $kbSet) 
{
    $thisKB = $kb.KB_TYPEID
    $flatPath = ""

    $nextParentID = $kb.PARENTID

    if ($nextParentID -eq -1)
    {
        $separator = ""
        $flatPath = $kb.TYPE
    }
    else 
    {
        $separator = "] - ["
        $flatPath = $separator + $kb.TYPE

        do
        {
            # look inside entire kbSet for info about next parent up the chain
            $kbSet | Where-Object -Property KB_TYPEID -eq $nextParentID -OutVariable thisParent
            
            $nextParentID = $thisParent.PARENTID

            if ($nextParentID -eq -1) { $separator = "" }

            $flatPath = $separator + $thisParent.TYPE + $flatPath

        } while ($nextParentID -ne -1)            
    }

    $flatPath = "[" + $flatPath + "]"

    $newLine = "" + $kb.KB_TYPEID + "|" + $kb.TYPE + "|" + $flatPath

    add-content -path $scratchKBPath -value $newLine
}

#-----------------------------------------------------------
# Read all flattened directory paths from temp file into PS object.
# For each directory path, Create it in Google Drive
#-----------------------------------------------------------

$allKBDirectories = Import-Csv -Path $scratchKBPath -Delimiter "|"

foreach ($kbDir in $allKBDirectories)
{
    $driveDir = $kbDir.KB_PATH
    gam user $driveTargetUser add drivefile drivefilename $kbDir.KB_PATH mimetype gfolder parentname '[solution-test]'
}


#-----------------------------------------------------------
# Set up spreadsheet with index of all solutions found
#-----------------------------------------------------------

$fileHeaderSolutions = "" +
    "S_NAME" + $defaultDelimiter +
    "ATT_NAMES" + $defaultDelimiter + 
    "ATT_OWNERS" + $defaultDelimiter +
    "S_UPDATED" + $defaultDelimiter +
    "S_PATH" + $defaultDelimiter + 
    "S_ID" + $defaultDelimiter + 
    "KB_ID" + $defaultDelimiter

add-content -path $scratchSolutionsPath -value $fileHeaderSolutions

$fileWriteCount = 0

#-----------------------------------------------------------
# For each solution, find any attachments, upload to drive
# Also create a drive file for track-it solution's free-form text field
#-----------------------------------------------------------

foreach ($solution in $solutionSet)
{
    # Read current directory for targeting correct parent in Drive
    $allKBDirectories | Where-Object -Property KB_ID -eq $solution.KB_TYPEID -OutVariable currentDir
    $currentDirPath = $currentDir.KB_PATH

    # Search attachment set for any related to current solution ID
    $attachmentSet | Where-Object -Property AttachmentOwnerId -eq $solution.SOLUTIONID -OutVariable currentAttachments
    $attNames = ""
    $attOwners = ""
    $attDrivePrefix = "" + $solution.SOLUTIONID
    
    $solutionTitle = $solution.SUMMARY   
    $pattern = '[^a-zA-Z0-9]'
    $solutionTitle = $solutionTitle -replace $pattern, '_' 
    $solutionDriveTitle = $solutionTitle + "-" + $attDrivePrefix

    foreach ($att in $currentAttachments)
    {
        $attPrefix = $attachmentPath + "\" + $att.AttachmentOwnerId + "\"
        $attOwners = $attOwners + $att.LastModifiedBy + " | "

        $driveAttLocal = $attPrefix + $att.LogicalName
        $driveAttName = $solutionDriveTitle + "-ATTACHMENT-(CONVERTED)"
        $driveAttNameOriginal = $solutionDriveTitle + "-ATTACHMENT-(ORIGINAL)"

        $driveNamePieces = $att.LogicalName.Split(".")
        $driveAttExt = $driveNamePieces[-1]

        # If local file exists (some solutions refer to files that have been deleted),
        # then upload to drive in the correct parent folder

        # TO ADD: check for PDF and force PDF mimetype (the OCR is terrible; better to keep PDFs readable as uploaded to track-it)

        if (test-path $driveAttLocal)
        {
            $fileWriteCount++
            if ( ($driveAttExt -eq "pdf") -or ($driveAttExt -eq "jpg") -or ($driveAttExt -eq "bmp") -or ($driveAttExt -eq "gif") -or ($driveAttExt -eq "png") ) 
            {
                $rightNow = get-date -format filedatetime
                $errMsg = $rightNow + " -- PDF or image found, don't convert: " + $driveAttLocal
                add-content -path $appLog -value $errMsg
                
                $gamConsoleLog = gam user $driveTargetUser add drivefile localfile $driveAttLocal drivefilename $driveAttNameOriginal parentname $currentDirPath            
                
                # Create direct link to drive file (original mime type)...

                $fileID = $gamConsoleLog.split("(")[-1].split(")")[0]
                $driveURL = $driveURLPrefix + $fileID
                $driveLink = "<a href=" + $driveURL + " target=_blank>View Att (" + $driveAttNameOriginal + ")</a>"
                $attNames = $attNames + $driveLink + "<br>"

                $gamConsoleLog = gam user $driveTargetUser add drivefile localfile $driveAttLocal convert drivefilename $driveAttName parentname $currentDirPath            

                # Create direct link to drive file (converted)...

                $fileID = $gamConsoleLog.split("(")[-1].split(")")[0]
                $driveURL = $driveURLPrefix + $fileID
                $driveLink = "<a href=" + $driveURL + " target=_blank>View Att (" + $driveAttName + ")</a>"
                $attNames = $attNames + $driveLink + "<br>"
            }
            else
            {
                $rightNow = get-date -format filedatetime
                $errMsg = $rightNow + " -- CONVERT this file: " + $driveAttLocal
                add-content -path $appLog -value $errMsg

                $gamConsoleLog = gam user $driveTargetUser add drivefile localfile $driveAttLocal convert drivefilename $driveAttName parentname $currentDirPath

                # Create direct link to drive file (converted)...

                $fileID = $gamConsoleLog.split("(")[-1].split(")")[0]
                $driveURL = $driveURLPrefix + $fileID
                $driveLink = "<a href=" + $driveURL + " target=_blank>View Att (" + $driveAttName + ")</a>"
                $attNames = $attNames + $driveLink + "<br>"
            }
        }
        else 
        {
            $rightNow = get-date -format filedatetime
            $errMsg = $rightNow + " -- not found: " + $driveAttLocal
            add-content -path $appLog -value $errMsg
        }
    }

    $errMsg = "Total files found: " + $fileWriteCount + " | At current solution ID: " + $solution.SOLUTIONID
    add-content -path $appLog -value $errMsg

    <#
        Use this if we don't need a local copy of all solution files
        i.e. each solution gets written to a temp file, uploaded, then deleted

    if (test-path $tmpSolution) {
        remove-item -path $tmpSolution
    }   
    #>

    $thisSolutionDetail = $scratchPath + "\" + $solution.SOLUTIONID + "-" + $solutionTitle + ".html"

    $thisSolutionHeader = "<hr><p>Last updated on: " + 
        $solution.LASTUPDATED + 
        " (according to Track-it database)</p><hr>&nbsp;<br>" +
        "Attached files (if applicable):<br>" +
        $attNames +
        "<hr>&nbsp;<br>"


    add-content -path $thisSolutionDetail -value $thisSolutionHeader
    add-content -path $thisSolutionDetail -value $solution.DETAIL

    $gamConsoleLog = gam user $driveTargetUser add drivefile localfile $thisSolutionDetail convert drivefilename $solutionDriveTitle parentname $currentDirPath

    # Create direct link to drive file (solution text from database)...

    $fileID = $gamConsoleLog.split("(")[-1].split(")")[0]
    $driveURL = $driveURLPrefix + $fileID
    $driveLink = "=hyperlink(`"" + $driveURL + "`", `"" + $solutionDriveTitle + "`")"
    $solutionTitle = $driveLink

    if ($attNames -ne "") 
    {
        $attNames = "Has Attachments!"
    }

    $newLine = "" + 
        $solutionTitle + $defaultDelimiter +
        $attNames + $defaultDelimiter +
        $attOwners + $defaultDelimiter +
        $solution.LASTUPDATED + $defaultDelimiter +
        $currentDirPath + $defaultDelimiter +
        $solution.SOLUTIONID + $defaultDelimiter + 
        $solution.KB_TYPEID + $defaultDelimiter

    add-content -path $scratchSolutionsPath -value $newLine

}

#-----------------------------------------------------------
# Upload spreadsheet index of solutions exported from Track-it
#-----------------------------------------------------------

gam user $driveTargetUser add drivefile localfile $scratchSolutionsPath convert drivefilename 'Index of Solutions copied from Track-it' parentname '[solution-test]'
