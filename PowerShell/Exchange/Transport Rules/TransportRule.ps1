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

#All of our Distribution Groups
$DistributionGroups = ((Get-ADGroup -Filter { (GroupCategory -eq "Distribution") } -Properties DistinguishedName, Name, SamAccountName, Mail, SID, ObjectGUID | Where-Object { ($_.Mail -ne $null) -AND ($_.distinguishedName -notlike "*OU=Distribution List Disabled,OU=DO NOT REMOVE ACCTS,OU=domain,DC=domain,DC=com") } ) | sort Name)

#We will be grouping them up by 90 distribution lists per group
$groupCount = [Math]::Ceiling($DistributionGroups.Count / 90)

$BaseName = "Access to DL from A and B and C"
$LastPriority = (Get-TransportRule | Select-Object -Last 1).Priority

for ($i = 1; $i -le $groupCount; $i++) 
{
    #Reset variables each iteration
    $Group = $endIndex = $startIndex = $Priority = $groupName = $NULL

    $Priority = ($LastPriority + $i)

    # Create our start and end indexes
    $startIndex = ($i - 1) * 90
    $endIndex = [Math]::Min($startIndex + 89, $DistributionGroups.Count - 1)

    #Grab 90 groups
    $Group = ($DistributionGroups.ObjectGUID[$startIndex..$endIndex])
    $Group

    #$groupName = "$($BaseName + $i)"
    $groupName = "$($BaseName + " ('$($Group[$startIndex])' - '$($Group[$endIndex])')")"

    #Create the initial rule with the first member
    New-TransportRule -Name "$($groupName)" -Comments "To allow A and B employees to email into C DL." -Priority $Priority -Enabled $False -Mode "Enforce" -AnyOfToCcHeader $Group.guid -ExceptIfSenderDomainIs 'ABC.com', 'DEF.com', 'GHI.org', 'JKL.org', 'MNO.org', 'PQR.org' -RejectMessageReasonText 'This recipient only accepts e-mail from within the ABC organization.' -RejectMessageEnhancedStatusCode '5.7.1'
    
    #Set-TransportRule -Identity "$groupName" -AnyOfToCCHeader $Group
}
