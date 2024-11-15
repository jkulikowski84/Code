CLS

Clear-Variable DSQUsers, AddUsersArray, RemoveUsersArray -force -Confirm:$False -ErrorAction SilentlyContinue

$OUs = @(
    "OU=Employees,OU=domain,DC=domain,DC=com",
    "OU=Contractors,OU=domain,DC=domain,DC=com",
    "OU=Providers,OU=domain,DC=domain,DC=com",
    "OU=Sharepoint,OU=domain,DC=domain,DC=com",
)

$PropsDSQ = @(
    #Enabled Accounts ONLY
    "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"+
    #User Objects Only
    "(objectCategory=person)(objectClass=user)"+
    #Password is not expired
    "(!(userAccountControl:1.2.840.113556.1.4.803:=65536))"+
    #Filter out Users that are inactive for 90+ days based on lastLogon
    "(LastLogon>=$(((Get-Date).AddDays(-90)).ToFileTime()))"
)

$DSQUsers = @()

#Grab all of our users based on the parameters we set
foreach($OU in $OUs)
{
    #Check DSQuery and trim empty lines
    $DSQUsers += ((dsquery * $OU -filter "(&($PropsDSQ))" -limit 0 -attr distinguishedName | Select-Object -Skip 1) | ? {$_.trim() -ne ""})
}

#Count how many Users there are based on our parameters
Write-Output "Total Users in AD found to be added to the dynamic group: $($DSQUsers.Count)"

#Our Dynamic Group
$DynamicGroup = "GS_Office_Dynamic_Access"

#Get Current Users in the group
$CurrentGroupMembers = dsquery * -filter "(&(memberof=CN=$DynamicGroup,OU=Security Groups,OU=domain,DC=domain,DC=com))" -limit 0

#User count in group
Write-Output "Current total of users in the group: $($DSQUsers.Count)"

$TrimmedDSQueryUsers = [string[]]$DSQUsers.Trim()
$TrimmedCurrentGroupUsers = [string[]]($CurrentGroupMembers -replace '"','')

#========= Remove users from Group

$RemoveUsersArray = [String[]][Linq.Enumerable]::Except($TrimmedCurrentGroupUsers, $TrimmedDSQueryUsers)

if(([string]::IsNullOrEmpty($RemoveUsersArray)) -ne $True)
{
    Write-Output "`nRemoving Users:`n"
    $RemoveUsersArray

    #Remove Users
    Remove-ADGroupMember -Identity $DynamicGroup -Members $RemoveUsersArray -Confirm:$False #-WhatIf
}

#========= Add users to Group

$AddUsersArray = [String[]][Linq.Enumerable]::Except($TrimmedDSQueryUsers, $TrimmedCurrentGroupUsers)

if(([string]::IsNullOrEmpty($AddUsersArray)) -ne $True)
{
    Write-Output "`nAdding Users:`n"
    $AddUsersArray

    #Add User
    Add-ADGroupMember -Identity $DynamicGroup -Members $AddUsersArray -Confirm:$False #-WhatIf
}