CLS

#Set-ExecutionPolicy RemoteSigned
$userUPN = "365admin@domain.onmicrosoft.com" 
$AESKeyFilePath = (Split-Path $script:MyInvocation.MyCommand.Path) + "\AES.key"
$SecurePwdFilePath =  (Split-Path $script:MyInvocation.MyCommand.Path) + "\AESpassword.txt"
$AESKey = Get-Content -Path $AESKeyFilePath -Force
$securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

#create a new psCredential object with required username and password
$UserCredential = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session -DisableNameChecking -AllowClobber

CLS

#Let's start with new rooms
$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited | Sort-Object WhenMailboxCreated -Descending

#Permissions and descriptions of the access rights are found here:
#https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-mailboxfolderpermission?view=exchange-ps

foreach ($room in $rooms) 
{ 
	$room.name
	write-output ""
	$ResourcePermission = $room.name + ":\Calendar";
	
	###Output the permissions. This is just for troubleshooting/debugging
	#$ResourcePermission
	
	###View the mailbox permissions for each resource room. This is also for troubleshooting
	#Get-MailboxFolderPermission $ResourcePermission

    ### Check User Groups ###
    $AccessRights = ((Get-MailboxFolderPermission $ResourcePermission).AccessRights)
    $users = ((Get-MailboxFolderPermission $ResourcePermission).user).displayname

    foreach ($user in $users)
    {
        if(($user -eq "Default") -and ($AccessRights -ne "LimitedDetails"))
        {
            Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights LimitedDetails
        }

        if(($user -eq "Anonymous") -and ($AccessRights -ne "None"))
        {
            Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights None
        }

        if(($user -eq "Room Schedulers") -and ($AccessRights -ne "Author"))
        {
            Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights Author
        }

        if(($user -eq "Room Schedulers Editor") -and ($AccessRights -ne "Editor"))
        {
            Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights Editor
        }
    }

    if($users -notcontains "Default")
    {
        Add-MailboxFolderPermission -Identity $ResourcePermission -User "Default" -AccessRights LimitedDetails
    }

    if($users -notcontains "Anonymous")
    {
        Add-MailboxFolderPermission -Identity $ResourcePermission -User "Anonymous" -AccessRights None
    }

    if($users -notcontains "Room Schedulers")
    {
        Add-MailboxFolderPermission -Identity $ResourcePermission -User "Room Schedulers" -AccessRights Author
    }

    if($users -notcontains "Room Schedulers Editor")
    {
        Add-MailboxFolderPermission -Identity $ResourcePermission -User "Room Schedulers Editor" -AccessRights Editor
    }

    write-output ""
}

Remove-PSSession $Session