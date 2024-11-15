CLS

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

#Export mailboxes with the content filter between 2 dates

New-MailboxExportRequest -Name HBadalExport1 -Mailbox "HBadal" -ContentFilter "((Body -like '*MARGIE*') -or (Subject -like '*MARGIE*')) -and ((Received -lt '04/20/2022') -and (Received -gt '09/09/2021')) -and ((Sent -lt '04/20/2022') -and (Sent -gt '09/09/2021'))" -FilePath "\\CIFS\ExchangeArchive\HBadal_ACTIVE_MARGIE.pst"
