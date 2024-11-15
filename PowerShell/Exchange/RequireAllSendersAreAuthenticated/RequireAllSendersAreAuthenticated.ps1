CLS

$SessionsRunning = get-pssession

if($SessionsRunning.ComputerName -like "*ExchangeServer*")
{
    #If session is running we don't need to do anything
}
else
{
    $MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer.domain.com/PowerShell/ -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $MBXSession -AllowClobber -DisableNameChecking
}

#All enabled users
$Users = (dsquery * -filter "(&(objectClass=Person)(objectCategory=Person)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(msExchWhenMailboxCreated=*))" -limit 0 -attr Name | sort).trim()

#All mailboxes with "RequireAllSendersAreAuthenticated" enabled
$UsersWithEauthEnabledMbx = ((Get-Mailbox -Filter { (RequireAllSendersAreAuthenticated -eq $true) } -ResultSize Unlimited).Name | sort).trim()

#Cross reference
$FilteredResults = $Null

$FilteredResults = ForEach ($Item in $UsersWithEauthEnabledMbx) 
{
    If($Item -in $Users)
    {
        $Item
    }
}

$FilteredResults 
