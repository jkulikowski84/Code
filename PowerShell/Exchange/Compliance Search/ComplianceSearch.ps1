CLS

#Setup our Attributes
$date2 = "sent>={0:MM/dd/yyyy} AND" -f (get-date).AddMonths(-1)
$date3 = get-date -format "'sent'<=MM/dd/yyyy 'AND'"
$Spammer = "from:staffduty072@gmail.com"

#========================================

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

CLS

#========================================

#The Name of our Compliance Search
$ComplianceName = ("$Spammer - $(Get-Date -format MM/dd/yyy)").Replace("from:","")

#Configure our new compliance search
New-ComplianceSearch -Name $ComplianceName -ExchangeLocation all -ContentMatchQuery $date2, $date3, $Spammer

#Start the compliance search
Start-ComplianceSearch -Identity $ComplianceName

#Start of the search

#Check every 5 seconds to see if our search is completed
Do
{
    (Get-ComplianceSearch -Identity $ComplianceName).Status
    sleep -Seconds 5

} while(((Get-ComplianceSearch -Identity $ComplianceName).Status) -ne "Completed")

CLS

#Review the Results
(Get-ComplianceSearch -Identity $ComplianceName).SuccessResults

Pause

#Configure the compliance purge
New-ComplianceSearchAction -SearchName $ComplianceName -Purge -PurgeType SoftDelete

$PurgeName = $ComplianceName +"_Purge"

Do
{
    (Get-ComplianceSearchAction -Identity $PurgeName).status
    sleep -Seconds 5

} while(((Get-ComplianceSearchAction -Identity $PurgeName).Status) -ne "Completed")

(Get-ComplianceSearchAction -Identity $PurgeName).SuccessResults

Pause

Remove-ComplianceSearch -Identity $ComplianceName -Confirm:$false -ErrorAction SilentlyContinue
Remove-ComplianceSearchAction -Identity $PurgeName -Confirm:$false -ErrorAction SilentlyContinue
