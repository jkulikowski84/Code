CLS

#Variables
$Sender = "jholm1@domain.com"
$Recipient = "thomaspathiyildo@gmail.com"
$ExportToFile = "C:\temp\jholm1.csv"

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

foreach ($ServerInfo in $ServersInfo)
{
    Get-MessageTrackingLog -ResultSize Unlimited -Server $ServerInfo.Name -Sender $Sender | Where-Object { ( $_.Recipients -like "$($Recipient)" ) } | Select-Object Timestamp,ServerHostname,ClientHostname,ConnectorId,Source,EventId,Sender,ReturnPath,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject,MessageId | Select-Object Timestamp,ServerHostname,ClientHostname,ConnectorId,Source,EventId,Sender,ReturnPath,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject,MessageId | Sort-Object -Property Timestamp | Export-Csv $ExportToFile
}
