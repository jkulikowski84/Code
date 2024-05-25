CLS

#Root Path
$rootPath = $pwd.ProviderPath #$PSScriptRoot #$pwd.ProviderPath

Write-Host "Importing all necessary modules."

#******************************************************************
#                            PREREQUISITES
#******************************************************************
#Nuget - Needed for O365 Module to work properly

if(!(Get-Module -ListAvailable -Name NuGet))
{
    #Install NuGet (Prerequisite) first
	Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$False
}

#******************************************************************

#Connect w\ Active Directory Module
& $rootPath\AD-Module\AD-module.ps1

#Load the O365 Module
& $rootPath\O365-Module\O365-module.ps1

#Load the on-prem Exchange Module
& $rootPath\AAD-Module\AAD-module.ps1

#Load the on-prem Exchange Module
& $rootPath\MBX-Module\MBX-module.ps1

#Clear screen after loading all the modules/sessions
CLS

#******************************************************************
#                         PUT CODE BELOW
#******************************************************************

#GLOBAL VARIABLES
$global:user = $ADUser = $NULL

function Main
{
    do
    {
	    #Get users input to get the username
        $user = Read-Host "Type in the username or type 'exit' to exit"
	    $mos = $user + "@domain.mail.onmicrosoft.com"
	    $useremail = $user + "@domain.com"

        if($user -eq "exit")
        {
            exit
        }

        if(([string]::IsNullOrWhiteSpace($user) -eq $true))
        {
            main
        }

	    #Don't run it if you type "exit"
	    if($user -ne "exit")
	    {

            while (-not $ADUser) 
            {
                try 
                {
                    $ADUser = Get-ADUser -Filter "SamAccountName -like '$user'" -Properties *
                }
                catch {Write-Output "User not found in AD yet. Checking again in 5 seconds"; sleep 5}
            }

            #$exist = [bool](Get-Mailbox $useremail -erroraction SilentlyContinue)
            $exist = [bool](Get-RemoteMailbox $user -erroraction SilentlyContinue)

            if($exist -eq $true)
            {
                CLS
                write-host "This account already has an email associated with it."
                Main
            }

			if($NULL -ne $ADUser.mail)
            {
                CLS
                write-host "An account with this email: $($ADUser.mail) already exists"
                Main
            }
			
            #Make sure the user doesn't already have any Exchange attributes in their profile
            #It's a new user, so All Exchange Attributes should be cleared out

            #$ADaccount = get-user $user
            $ADaccount = Get-ADUser $user -Properties *
            $FullDistinguishName = "LDAP://" + $ADaccount.distinguishedName

            #Lets make sure all Exchange Attributes are cleared out
            $AccountEntry = New-Object DirectoryServices.DirectoryEntry $FullDistinguishName 
            $AccountEntry.PutEx(1, "mail", $null) 
            $AccountEntry.PutEx(1, "HomeMDB", $null) 
            $AccountEntry.PutEx(1, "HomeMTA", $null) 
            $AccountEntry.PutEx(1, "legacyExchangeDN", $null) 
            $AccountEntry.PutEx(1, "msExchMailboxAuditEnable", $null) 
            $AccountEntry.PutEx(1, "msExchAddressBookFlags", $null) 
            $AccountEntry.PutEx(1, "msExchArchiveQuota", $null) 
            $AccountEntry.PutEx(1, "msExchArchiveWarnQuota", $null) 
            $AccountEntry.PutEx(1, "msExchBypassAudit", $null) 
            $AccountEntry.PutEx(1, "msExchDumpsterQuota", $null) 
            $AccountEntry.PutEx(1, "msExchDumpsterWarningQuota", $null)  
            $AccountEntry.PutEx(1, "msExchHomeServerName", $null) 
            $AccountEntry.PutEx(1, "msExchMailboxAuditEnable", $null) 
            $AccountEntry.PutEx(1, "msExchMailboxAuditLogAgeLimit", $null) 
            $AccountEntry.PutEx(1, "msExchMailboxGuid", $null) 
            $AccountEntry.PutEx(1, "msExchMDBRulesQuota", $null) 
            $AccountEntry.PutEx(1, "msExchModerationFlags", $null) 
            $AccountEntry.PutEx(1, "msExchPoliciesIncluded", $null) 
            $AccountEntry.PutEx(1, "msExchProvisioningFlags", $null) 
            $AccountEntry.PutEx(1, "msExchRBACPolicyLink", $null) 
            $AccountEntry.PutEx(1, "msExchRecipientDisplayType", $null) 
            $AccountEntry.PutEx(1, "msExchRecipientTypeDetails", $null) 
            $AccountEntry.PutEx(1, "msExchTransportRecipientSettingsFlags", $null) 
            $AccountEntry.PutEx(1, "msExchUMDtmfMap", $null) 
            $AccountEntry.PutEx(1, "msExchUMEnabledFlags2", $null) 
            $AccountEntry.PutEx(1, "msExchUserAccountControl", $null) 
            $AccountEntry.PutEx(1, "msExchVersion", $null)  
            $AccountEntry.PutEx(1, "proxyAddresses", $null)  
            $AccountEntry.PutEx(1, "showInAddressBook", $null)  
            $AccountEntry.PutEx(1, "mailNickname", $null) 
            $AccountEntry.SetInfo() | Out-Null

            Enable-RemoteMailbox $user -RemoteRoutingAddress $mos -Confirm:$false | Out-Null

            Write-host "Waiting for Email to Sync with O365 before we license. Please be patient this can take several minutes."

            #Before we start, we need to make sure the user has synced up with O365 already, otherwise there will be errors
	        try
            {
                while (-not $GetMailbox)
                {
                    Try 
                    {
                        if((Get-ADSyncScheduler).SyncCycleInProgress -eq "True")
                        {

                        }
                        else
                        {
                            Start-ADSyncSyncCycle delta | Out-Null
                        }
                        Sleep 10
                    
                        #$exist = [bool](Get-RemoteMailbox $user -erroraction SilentlyContinue)
                        $GetMailbox = Get-Mailbox $useremail -ErrorAction SilentlyContinue                
                        #$GetMailbox = Get-RemoteMailbox $user #| -ErrorAction SilentlyContinue 
                    }
                    Catch {sleep 1}
                }
            }
            catch {}

            #This is the license we will be assigning to the user. The "EnterprisePack" license if Office365 E3
            $license = (Get-MsolAccountSku).AccountSkuId | Where-Object {$_ -like "domain:ENTERPRISEPACK" }
            $LicenseOptions = New-MsolLicenseOptions -AccountSkuID $license

            #Run another delta Sync
            #Invoke-Command -ComputerName "gle-aad-01.domain.net" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
            #Sync our changes with AD if a sync isn't already running
            if((Get-ADSyncScheduler).SyncCycleInProgress -eq "True")
            {

            }
            else
            {
                Start-ADSyncSyncCycle delta | Out-Null
            }

            #Readd the license
            Set-MsolUser -UserPrincipalName $useremail -UsageLocation 'US' -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $useremail -AddLicenses $license -LicenseOptions $LicenseOptions -ErrorAction SilentlyContinue
        }
    }
    until ($user -eq "exit")
}

#******************************************************************

#Our main Function
Main

#Clean up Sessions after use

if($NULL -ne (Get-PSSession))
{
    Remove-PSSession *
}

[GC]::Collect()
