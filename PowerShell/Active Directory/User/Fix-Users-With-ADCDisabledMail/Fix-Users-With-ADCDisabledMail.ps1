CLS

$Users = @'
User1
User2
User3
'@.Split("`n").Trim()

#We targetted only the 3 users above
#We get their AD properties, and specifically look at the ProxyAddresses Configured
ForEach($User in $Users) 
{
    $UserInfo = Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(sAMAccountName=$User))" -Properties *

    $SAM = $UserInfo.sAMAccountName

    $Proxy = $UserInfo.proxyAddresses

    #We want to set the correct information to legacyExchangeDN, so grab the info from ProxyAddresses
    $NewData = (($Proxy | Select-Object | Where-Object { ($_ -like "*cn=[^a-zA-Z0-9]*-$($SAM)*") } ) -split(":"))[1]

    #Check if the current value is correct, if not, then update it so it's correct.
    if($UserInfo.legacyExchangeDN -ne $NewData)
    {
        Set-ADUser -Identity $SAM -Replace @{ legacyExchangeDN = $NewData } 
    }
}
