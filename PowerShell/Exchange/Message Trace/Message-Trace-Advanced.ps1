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

#============

##Source
#https://practical365.com/tell-transport-rule-applied-email-message/

#Find info about the email message based on the MessageID
$logs = (Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}) | ForEach-Object {Get-MessageTrackingLog -MessageId "<16278dae2cc04da98adc68494d1d14a9@domain.com>" -ResultSize Unlimited -Server $_.Name } | Sort-Object -Property Timestamp 

#Sort the log(s) found from the message trace above
#$logs | Sort timestamp | Select eventid,source,messagesubject

<#
EventId   Source MessageSubject                 
-------   ------ --------------                 
RECEIVE   SMTP   Emailing: www.myworkday.com.har
FAIL      AGENT  Emailing: www.myworkday.com.har
AGENTINFO AGENT  Emailing: www.myworkday.com.har
#>

#Get more info about the email. Usually the "EventData" field has the information needed for example ruleId shows which transport rule blocked the email.
($logs | where {$_.eventid -eq "AGENTINFO"}).EventData | fl
#($logs).EventData | fl

#Grab the transport rule Info from the ID in the EventData section
#Get-TransportRule -Identity 6c60b250-fbba-40c9-802f-2d52a023bc70 | select description | fl