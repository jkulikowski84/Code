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
        $credentials = Get-Credential
        $session = New-PSSession -ComputerName $hostname -Credential $credentials
    }

    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

$Export = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Sec-Grps.csv"

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

Get-ADGroup -filter {groupCategory -eq 'Security'} -Properties * |
Select-Object -Property CanonicalName,
CN,
Created,
createTimeStamp,
Deleted,
Description,
DisplayName,
DistinguishedName,
extensionAttribute13,
GroupCategory,
GroupScope,
groupType,
HomePage,
instanceType,
isDeleted,
LastKnownParent,
ManagedBy,
@{L = "member"; E = {$_.member -join";"}},
@{L = "MemberOf"; E = {$_.MemberOf -join";"}},
@{L = "Members"; E = {$_.Members -join";"}},
Modified,
modifyTimeStamp,
Name,
nTSecurityDescriptor,
ObjectCategory,
ObjectClass,
ObjectGUID,
objectSid,
ProtectedFromAccidentalDeletion,
SamAccountName,
sAMAccountType,
sDRightsEffective,
SID,
@{L = "SIDHistory"; E = {$_.SIDHistory -join";"}},
uSNChanged,
uSNCreated,
whenChanged,
whenCreated | export-csv $Export
