CLS

#Variables
$username = "user@domain.com"
$message = "I am currently out of the office."

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

Try
{
    Set-MailboxAutoReplyConfiguration -Identity $username -AutoReplyState Enabled -InternalMessage $message -ErrorAction Stop
}
Catch
{
    Write-Host "The username $username doesn't exist"
}
