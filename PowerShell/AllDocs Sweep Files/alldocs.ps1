CLS

#Global Variables

#Path where you run script from
$path = (Split-Path $script:MyInvocation.MyCommand.Path)

#Path of our alldocs scanning folder
#$scans = "C:\temp\test\AllDocs"
$scans = "\\fileshare\scans\AllDocs"
$ScansFolders = (Get-ChildItem -Path $scans -Force -Attributes Directory)

#Scan each folder
foreach($scansfolder in $ScansFolders.fullname)
{
    #Check to see if any of the folders have a "ERROR_FILES" directory in there
    $ErrorFolders = (Get-ChildItem -Path $scansfolder -Force -Attributes Directory) | where {$_.name -eq "ERROR_FILES"}
    
    #For each root directory, 
    foreach($ErrorFolder in $ErrorFolders.fullname)
    {
        #This is the root of the Alldocs Folder
        $rootAlldocsFolder = $ErrorFolders.Parent.fullname

        #These are files that are in the "ERROR_FILES" Directory in the Alldocs folder
        $ScannedFiles = (Get-ChildItem -Path $ErrorFolder -Force)

        #Now we need to move the files
        foreach($ScannedFile in $ScannedFiles.fullname)
        {
            #Move the files to the root of the folder
            Move-Item -Path $ScannedFile -Destination $rootAlldocsFolder -force
        }

        #Once we are done moving the files to the root, delete the ErrorFolder
        Remove-Item –path $ErrorFolder
    }

    $GetThumbdbFile = (Get-ChildItem -Path $scansfolder -Force) | where {$_.name -eq "Thumbs.db"}

    foreach($ThumbdbFile in $GetThumbdbFile.fullName)
    {
        Remove-Item –path $ThumbdbFile -Force
    }
}
