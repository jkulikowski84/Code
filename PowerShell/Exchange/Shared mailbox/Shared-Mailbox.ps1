CLS

$OldPref = $global:ErrorActionPreference
$global:ErrorActionPreference = 'Stop'

function Get-ModuleAD() 
{
    #Add the import and snapin in order to perform AD functions
    #Get Primary DNS
    $DNS = (Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IpEnabled='True'" | ForEach-Object {$_.DNSServerSearchOrder})[1]

    #Convert IP to hostname
    $hostname = ([System.Net.Dns]::gethostentry($DNS)).HostName

    #Add the necessary modules from the server
    Try
    {
        $session = New-PSSession -ComputerName $hostname -Authentication Kerberos -ErrorAction Stop
    }
    Catch
    {
        $credentials = Get-Credential
        $session = New-PSSession -ComputerName $hostname -Credential $credentials
    }

    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

function SyncWithO365
{
    ###### Connect to O365 ######
    $userUPN = "365admin@domain.onmicrosoft.com" 
    $AESKeyFilePath = (Split-Path $script:MyInvocation.MyCommand.Path) + "\AES.key"
    $SecurePwdFilePath =  (Split-Path $script:MyInvocation.MyCommand.Path) + "\AESpassword.txt"
    $AESKey = Get-Content -Path $AESKeyFilePath -Force
    $securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

    #create a new psCredential object with required username and password
    $UserCredential = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)

    $O365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $UserCredential -Authentication Basic -AllowRedirection

    Import-PSSession $O365 -AllowClobber

    Connect-MsolService -Credential $UserCredential

    CLS

    #Write-Output "Configuring the mailbox. This can take a few (about 5) minutes."
    Write-Host "Configuring the mailbox. This can take a few" -NoNewline;
    Write-Host " (about 5) " -ForegroundColor Red -NoNewline;
    Write-Host "minutes."

    try
    {
        while (-not $GetMailbox)
        {
            Try 
            {
                Invoke-Command -ComputerName "gle-aad-01.domain.com" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
                Sleep 30

                $GetMailbox = Get-Mailbox $User                     
            }
            Catch {sleep 1}
        }
    }
    catch {}

    Sleep 5

    Get-Mailbox -identity $User | set-mailbox -type "Shared"

    CLS
}
#----------------------------------------------------------------------------------------------------------------
<#
    The worksheet variable will need to be modified before running this script. 
    Whatever the name of the worksheet is that you want to import data from, type that in below.
#>

function CreateNewSharedMailbox
{
	$DN = "OU=Shared Mailbox,DC=domain,DC=com"
	
	#Get the User information
	do
	{
        #Dynamic variables will be set to $NULL initially
        $UPN = $RecipientType = $ADUser = $SamAccountName = $TargetAddress = $GetMailbox = $ValidateUser = $NULL

		#Get users input to get the username
		$user = Read-Host "Type in the Shared Mailbox Name or type 'exit' to exit"
		
		#Check for spaces
		if($user -like "* *")
		{
			$user = $user.Replace(" ","-")
		}
		
		$UPN = $user + "@domain.com"
		$SamAccountName = $user[0..19] -join ""
        $TargetAddress = $user + "@domain.mail.onmicrosoft.com"
					
		#Don't run it if you type "exit"
		if($user -ne "exit")
		{
			#####################  Create the Shared Account  #####################

            $SharedMbx=`
            @{
                Name=$User
                SamAccountName=$SamAccountName
                Surname=$User
                DisplayName=$User
                UserPrincipalName=$UPN
                Path=$DN
                Enabled=$False
                ChangePasswordAtLogon=$False
                PasswordNeverExpires=$True
            }

            New-ADUser @SharedMbx

            CLS

            Write-Host "Creating the Shared Account '$User'"

			#####################  Make sure account created successfully before continuing  #####################

            while (-not $ADUser) 
            {
                Try 
                {
                    $ADUser = Get-ADUser -Filter "DisplayName -like '$user'" -Properties * -ErrorAction Stop
                }
                Catch {sleep 1}
            }

            while ($RecipientType -ne "User")
            {
                Try 
                {
                    $RecipientType = (Get-User $SamAccountName).RecipientType
                }
                Catch {sleep 1}
            }
			
            #Enable the remote mailbox
            Enable-RemoteMailbox -Identity $UPN -RemoteRoutingAddress $TargetAddress

            #####################  Configure Permissions to the Shared Mailbox  #####################
            CLS

            ### ============== Load O365 Module and sync the new Account ============== ###

            SyncWithO365

			### ============== Full Access Permissions ============== ###
			do
			{
				Write-Host "Usernames must be 'pre Windows 2000'" -ForegroundColor DarkYellow

				Write-Host "Type in the usernames of the people you want to give" -NoNewline;
                Write-Host " Full Access " -ForegroundColor Red -NoNewline;
                Write-Host "to for the mailbox, or type 'exit' to exit (Hit Enter in between each name and when you're done, type exit)"

                $userPerm = Read-Host

                Try
                {
                    $ValidateUser = Get-ADUser $userPerm

                    if($NULL -ne $ValidateUser)
                    {
                        if($userPerm -ne "exit")
                        {
                            Add-MailboxPermission $user -User $userPerm -AccessRights FullAccess –InheritanceType all -Confirm:$False
                        }
                    }
                    else
                    {
                        if($userPerm -ne "exit")
                        {
                            Write-Host "Make sure you have typed in the username correctly"
                            Sleep 1
                        }
                    }
				}
                Catch
                {
                    if($userPerm -ne "exit")
                    {
                        Write-Host "Make sure you have typed in the username correctly"
                        Sleep 1
                    }
                }

                Sleep 1
                CLS
				
			}until ($userPerm -eq "exit")

			CLS
			
			### ============== SendAs Permissions ============== ###
			do
			{
				Write-Host "Usernames must be 'pre Windows 2000'" -ForegroundColor DarkYellow
				
                Write-Host "Type in the usernames of the people you want to give" -NoNewline;
                Write-Host " SendAs Access " -ForegroundColor Red -NoNewline;
                Write-Host "to for the mailbox, or type 'exit' to exit (Hit Enter in between each name and when you're done, type exit)"

                $userSendAsPerm = Read-Host

                Try
                {
                    $ValidateUserPerm = Get-ADUser $userSendAsPerm

                    if($NULL -ne $ValidateUser)
                    {
				        if($userSendAsPerm -ne "exit")
				        {
                            Add-RecipientPermission -Identity $user -Trustee $userSendAsPerm -AccessRights SendAs -Confirm:$False
				        }
                    }
                    else
                    {
                        if($userSendAsPerm -ne "exit")
                        {
                            Write-Host "Make sure you have typed in the username correctly"
                            Sleep 1
                        }
                    }
                }
                Catch
                {
                    if($userSendAsPerm -ne "exit")
                    {
                        Write-Host "Make sure you have typed in the username correctly"
                        Sleep 1
                    }
                }

                Sleep 1
                CLS
				
			}until ($userSendAsPerm -eq "exit")
		}
		CLS
	}
	until ($user -eq "exit")

    #####################  Close out all of our sessions  #####################
    Remove-PSSession *
}

#==============================================================================================================================================================================================

### ============== Remotely Load ActiveDirectory Module ============== ###

if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

### ============== Remotely Load On-Prem Exchange Module ============== ###

#Check if we have a session open right now
$SessionsRunning = get-pssession

if($SessionsRunning.ComputerName -like "*tst-mbx-01*")
{
    #If session is running we don't need to do anything
}
else
{
    $MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://tst-mbx-01.domain.com/PowerShell/ -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $MBXSession -AllowClobber -DisableNameChecking
}

### ============== Clear Screen ============== ###

CLS

### ============== Run our MAIN function ============== ###

CreateNewSharedMailbox
