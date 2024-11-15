CLS

$LoadModules = @'
SplitPipeline
'@.Split("`n").Trim()

ForEach($Module in $LoadModules) 
{
    #Check if Module is Installed
    if($NULL -eq (Get-Module -ListAvailable $Module))
    {
        #Install Module
        Install-Module -Name $Module -Scope CurrentUser -Confirm:$False -Force
        #Import Module
        Import-Module $Module
    }

    #Check if Module is Imported
    if($NULL -eq (Get-Module -Name $Module))
    {
        #Install Module
        Import-Module $Module
    }
}

#Variables
$Sender = "s33patel@domain.com"
$Recipient = "ajohn@domain.com"
$ExportToFile = "C:\temp\s33patel.csv"


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

$ServersInfo = Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}

$ServersInfo | split-pipeline -count 64 -Variable Sender, Recipient {
    Begin
    {
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
    }
    Process
    {
        $ServerInfo = $_

        $Start = "10/01/2024 12:00AM"
        $Now = ([DateTime]::Now.ToString("MM/dd/yyy H:mmtt")).replace(".","/")

        if($NULL -eq $End)
        {
            $End = $Now
        }

        if(($NULL -ne $Start) -AND ($NULL -ne $End))
        {
            Get-MessageTrackingLog -ResultSize Unlimited -Server $ServerInfo.Name -Start $Start -End $End -Sender $Sender | Where-Object { ( $_.Recipients -like "$($Recipient)" ) } | Select-Object Timestamp,ServerHostname,ClientHostname,ConnectorId,Source,EventId,Sender,ReturnPath,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject,MessageId | Select-Object Timestamp,ServerHostname,ClientHostname,ConnectorId,Source,EventId,Sender,ReturnPath,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject,MessageId | Sort-Object -Property Timestamp
        }
        else
        {
            Get-MessageTrackingLog -ResultSize Unlimited -Server $ServerInfo.Name -Sender $Sender | Where-Object { ( $_.Recipients -like "$($Recipient)" ) } | Select-Object Timestamp,ServerHostname,ClientHostname,ConnectorId,Source,EventId,Sender,ReturnPath,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject,MessageId | Select-Object Timestamp,ServerHostname,ClientHostname,ConnectorId,Source,EventId,Sender,ReturnPath,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject,MessageId | Sort-Object -Property Timestamp
        }
    }
} | sort Timestamp | Export-Csv $ExportToFile