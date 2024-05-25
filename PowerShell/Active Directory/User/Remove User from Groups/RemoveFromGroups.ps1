CLS

function Get-ModuleAD() 
{
    #Add the import and snapin in order to perform AD functions
    #Get Primary DNS
    $DNS = (Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IpEnabled='True'" | ForEach-Object {$_.DNSServerSearchOrder})[1]

    #Convert IP to hostname
    $hostname = ([System.Net.Dns]::gethostentry($DNS)).HostName

    #Add the necessary modules from the server
    Try
    {
        $session = New-PSSession -ComputerName $hostname -Authentication Kerberos -ErrorAction Stop
    }
    Catch
    {
        $Auth = Get-Credential
        $session = New-PSSession -ComputerName $hostname -Authentication $Auth
    }

    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

CLS

$path = "OU=ABC,OU=DEF,OU=GHI,DC=domain,DC=com"
$groups = "RoadNotes","U-JC-Domain Users","Outlook Users","HPNI - Patient Portal Users","HPNI - Patient Care Services Security","HPNI - Patient Care Services","HPNI - All Staff","HPNI - All Nurses","JC - Referral Team Schedulers","HNI Userlist","Care Coordination Public Folder Editors","JC - Referral Team Schedulers","G-CS-Field SW-Sapphire","Resource Nurses","HPNI - All Barrington Staff","HPNI - Sapphire Team","JC - Ruby Team","JC - Amber Team","GRP-Jet_Team","GRP-Jet_Team-1-1261626272","MDM-iOSApp-Turboscan"

$Users = Get-ADUser -SearchBase $path -Filter * -Properties memberOf | Select SamAccountName, name, memberOf

foreach($group in $groups)
{
    $CheckGrp = Get-ADGroup -Filter "Name -like '$group'"

    foreach($user in $users)
    {
        Try
        {
            Remove-ADGroupMember -identity $group -members $user.samaccountname -Confirm:$false #-ErrorAction Stop -Verbose
            #Remove-ADGroupMember -identity $group.name -members $user.samaccountname -Confirm:$false -ErrorAction Stop -Verbose
            #Write-Output "User $($User.samaccountname) removed from $($group.name)" | 
            #Out-File C:\Logs\XXLogs.txt -Append
        }
        Catch
        {
            Write-Warning "Error removing user $($User.samaccountname)"
        }
    }
}