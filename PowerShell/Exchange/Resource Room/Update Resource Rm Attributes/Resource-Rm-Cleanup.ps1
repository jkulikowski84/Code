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

#Clear screen again
CLS

$SearchBase = "OU=Resource-Rooms,DC=domain,DC=com"
$ResourceRooms = Get-ADUser -Filter * -SearchScope OneLevel -SearchBase $SearchBase -Properties *

foreach($ResourceRoom in $ResourceRooms)
{
    $UPN = $ResourceRoom.UserPrincipalName
    $Email = $ResourceRoom.EmailAddress
    $Mail = $ResourceRoom.mail
    $DisplayName = $ResourceRoom.DisplayName
    $Proxies = $ResourceRoom.proxyAddresses
    $TargetSuffix = (($ResourceRoom.targetAddress).ToLower()) -replace "smtp:",""
    $SID = $ResourceRoom.SID.Value

    if($UPN -notlike $Email)
    {
        Set-ADUser $SID -EmailAddress $UPN 
        Set-ADUser $SID -Replace @{MailNickName = $UPN}

        foreach($Proxy in $Proxies)
        {
            $ProxyEmail = (($Proxy.ToLower()) -replace "smtp:", "")

            if($ProxyEmail -notlike "x500*")
            {
                if($ProxyEmail -like "*domainA.com*")
                {
                    Set-ADUser $SID -Remove @{proxyAddresses = $Proxy}
                }
                if($ProxyEmail -like "*domainB.com*")
                {
                    Set-ADUser $SID -Remove @{proxyAddresses = $Proxy}
                }
                if($ProxyEmail -like "*domainC.com*")
                {
                    Set-ADUser $SID -Remove @{proxyAddresses = $Proxy}
                }
            
                #Remove Primary SMTP
                if($Proxy -cmatch “^[A-Z]:*”)
                {
                    $OldPrimaryEmail = (($Proxy.ToLower()) -replace "smtp:", "")
                    
                    if($OldPrimaryEmail -like $Proxy)
                    {
                        Set-ADUser $SID -Remove @{proxyAddresses = "SMTP:$OldPrimaryEmail"}
                    }
                }
                
                if($UPN -like $ProxyEmail)
                {
                    Set-ADUser $SID -Remove @{proxyAddresses = "smtp:$UPN"}
                }

                #Remove current email if it already exists
                if($Email -like $ProxyEmail)
                {
                    Set-ADUser $SID -Remove @{proxyAddresses = "smtp:$Email"}
                }

                #Remove target email if it already exists
                if($TargetSuffix -like $ProxyEmail)
                {
                    Set-ADUser $SID -Remove @{proxyAddresses = "smtp:$TargetSuffix"}
                }
                
            }

            Set-ADUser $SID -Add @{proxyAddresses = "SMTP:$UPN"}
            Set-ADUser $SID -Add @{proxyAddresses = "smtp:$TargetSuffix"}
        }
    }
}

#Sync the changes
Invoke-Command -ComputerName "gle-aad-01.domain.com" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}