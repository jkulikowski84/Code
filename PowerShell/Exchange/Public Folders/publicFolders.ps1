CLS

#Get all out connected sessions
$SessionsRunning = get-pssession

#Check if we're already connected to Exchange
if($SessionsRunning.ComputerName -like "*ExchangeServer*")
{
    #If session is running we don't need to do anything
}
else
{
    $MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer.domain.com/PowerShell/ -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $MBXSession -AllowClobber -DisableNameChecking
}

#Grab all public Folders
$PublicFolders = Get-PublicFolder -Identity "\" -Recurse

#Grab all Public Folder Permissions
foreach ($PublicFolder in $PublicFolders) 
{
    if($NULL -ne $PublicFolder.ParentPath)
    {
        Get-PublicFolderClientPermission -Identity "$($PublicFolder.ParentPath)"
    }
}


