CLS

#============ Variables

$Sender = "user@domain.com"
$Subject = "Test - dont read"
$Recipients = @'
someone@domain.com
'@.Split("`n").Trim()

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

foreach($Recipient in $Recipients)
{
    $Result = $NULL

    Write-Host "Searching the $($Recipient) mailbox for an email with the subject:" -NoNewline
    Write-Host "$($Subject)" -ForegroundColor yellow

    if($NULL -eq $Result)
    {
        #Loop until we see the email (sometimes it can take a minute or so)
        do
        {
            #Check if we see the email in the mailbox (EstimateResultOnly parm prevents the log from sending each search to the target)
            #Remove the "-AsJob" switch to not run this as a job
            $Result = search-mailbox -EstimateResultOnly -identity $Recipient -searchquery "Subject:$($Subject)" -WarningAction silentlyContinue -AsJob

            #Wait 10 seconds before checking again
            sleep -Milliseconds 250

        }until(($Result.ChildJobs.output).ResultItemsCount -ne 0) #($Result.ResultItemsCount -ne 0)
    }

    if($Result)
    {
        #We found the email, let's try to delete it after confirming
        Get-Mailbox -Identity $Recipient -ResultSize Unlimited  -WarningAction silentlyContinue | Search-Mailbox -SearchQuery "Subject:$($Subject)" -loglevel full -TargetMailbox $Sender -TargetFolder Inbox -DeleteContent -WarningAction silentlyContinue
    }

    CLS
}
