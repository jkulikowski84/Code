CLS

#Path to where you want to save attachments from.
$FolderPath = "\\user@domain.com\Inbox\Folder1\SubFolder1"

#Check to make sure our Folder Path exists. If it doesn't, create it.
$Year = (Get-Date -Format 'yyyy')

if((Test-Path "\\File_Server\IT\Server Team\$Year") -eq $False)
{
    New-Item -Path "\\File_Server\IT\Server Team\$Year" -ItemType Directory
}

$MonthNum = (Get-Date -Format 'MM') + "_" + (Get-Culture).DateTimeFormat.GetMonthName((Get-Date -Format 'MM'))

if((Test-Path "\\File_Server\IT\Server Team\$Year\$MonthNum") -eq $False)
{
    New-Item -Path "\\File_Server\IT\Server Team\$Year\$MonthNum" -ItemType Directory
}

# use MAPI name space
$outlook = new-object -com outlook.application; 
$mapi = $outlook.GetNameSpace("MAPI");

# set the Inbox folder id
$olDefaultFolderInbox = 6
$inbox = $mapi.GetDefaultFolder($olDefaultFolderInbox)

# access the subfolder
$FirstlvlSubfolder = $inbox.Folders | Where-Object { $_.FolderPath -like "*Folder1*" }
$SecondlvlSubfolder = $FirstlvlSubfolder.Folders | Where-Object { $_.FolderPath -eq $FolderPath }

$emails = $SecondlvlSubfolder.Items

foreach ($email in $emails) 
{
    #Only save the emails with attachments
    if($NULL -ne $email.Attachments)
    {
        #Path to Save our Attachments
        $FolderYear = (Get-Date $($email.ReceivedTime) -Format 'yyyy')
        $FolderMonth = (Get-Date $($email.ReceivedTime) -Format 'MM') + "_" + (Get-Culture).DateTimeFormat.GetMonthName((Get-Date $($email.ReceivedTime) -Format 'MM'))
        $SaveFolderPath = "\\File_Server\IT\Server Team\" + $FolderYear + "\" + $FolderMonth

        $email.Attachments | foreach {
            $fileName = $_.FileName

            # save the attachment
            $_.saveasfile((Join-Path $SaveFolderPath $fileName))
        }
    }
} 