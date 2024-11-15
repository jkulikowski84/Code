CLS

#============ Do not edit below this line ===============================

#Connect to Exchange server remotely if we're not already connected
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

$OAB = ((Get-OfflineAddressBook).Identity).replace("\","")

$OAB_GUIDS = ((Get-OfflineAddressBook).GUID).guid

$ServersInfo = Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}

foreach ($ServerInfo in $ServersInfo)
{
    #Clear Variables each iteration
    $done = $Path = $ServerName = $NULL

    $ServerName  = $ServerInfo.Name
    
    #write-Host -ForegroundColor Green $ServerName

    Foreach($OAB_GUID in $OAB_GUIDS)
    {
        $PathExists = $NULL
        $Path = "\\$($ServerName)\C$\Program Files\Microsoft\Exchange Server\V15\ClientAccess\OAB\$($OAB_GUID)"
        
        Try
        {
            if(Test-Path($Path))
            {
                if($NULL -eq $Done)
                {
                    $done = 1
                    write-Host -ForegroundColor Green "$ServerName`n"
                }
                Write-Host "`t OAB $($OAB_GUID) synchronized on: $((Get-Item $Path).LastWriteTime)"
            }
            else
            {
                if($NULL -eq $Done)
                {
                    $done = 1

                    write-Host -ForegroundColor Red "$ServerName`n"
                }
            }
        }
        Catch
        {}
    }

    write-Host "`n"
}

#Update-OfflineAddressBook -Force -identity $OAB
