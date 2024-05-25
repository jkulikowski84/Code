CLS

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
        $Auth = Get-Credential
        $session = New-PSSession -ComputerName $hostname -Authentication $Auth
    }

    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

#Install the module that will let us perform certain tasks in Excel
#Install PSExcel Module for powershell
if(!(Get-Module -ListAvailable -Name ImportExcel))
{
	#Install NuGet (Prerequisite) first
	Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$False
	
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$False
	Import-Module ImportExcel
} 

#Clear screen again
CLS

#----------------------------------------------------------------------------------------------------------------
<#
    The worksheet variable will need to be modified before running this script. 
    Whatever the name of the worksheet is that you want to import data from, type that in below.
#>
$worksheet = "Sheet1"

#The file we will be reading from
$ExcelFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Final.xlsx"

$Import = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

foreach ($Room in $Import)
{
    #Define our variables
    $DisplayName = ($Room."New Room Name").Trim()
    $OldRoom = ($Room."Old Room Name").Trim()
    $Email = ($Room."Email Address").Trim()
    $Office = ($Room."Office").Trim()
    $emailsuffix = ($Email.Substring(0, $Email.IndexOf('@'))).Trim()
    $SamAccountName = $emailsuffix[0..19] -join ""
    $TargetAddress = ("SMTP:" + $emailsuffix + "@domain.mail.onmicrosoft.com").Trim()    
    
    if($OldRoom -ne "New Resource")
    {
		$validateRoom = Get-ADUser -Filter "Name -like '$OldRoom'" -Properties *
        $SID = $validateRoom.SID.Value
        $DN = $validateRoom.DistinguishedName

        if($NULL -ne $validateRoom)
        {
            #####################  Clear Settings  #####################
            Try
            {
                #Clear out some attributes we will overwrite
                Set-ADUser $SID -Clear givenName, description #-WhatIf
            }
            Catch
            {
                <#
                Write-Output "Error for: $validateRoom.Name New name: $DisplayName"
                $_.Exception.Message 
                $_.Exception.ItemName 
                $_.InvocationInfo.MyCommand.Name 
                $_.ErrorDetails.Message
                $_.InvocationInfo.PositionMessage 
                $_.CategoryInfo.ToString()
                $_.FullyQualifiedErrorId 
                Write-Output "-----------------------------"
                Write-Output ""
                #>
            }
            #####################  Update Attributes from existing resource  #####################
            Try
            {
                #SamAccountName has a character limitation of 20 characters, so we need to make sure we trim the attribute value otherwise it will fail
                $SamAccountName = $emailsuffix[0..19] -join ""

                #Populate fields
                Rename-ADObject -Id $DN -NewName $DisplayName #-WhatIf
                Set-ADUser $SID -DisplayName $DisplayName -EmailAddress $Email -Office $Office -SamAccountName $SamAccountName -Surname $DisplayName -UserPrincipalName $Email #-WhatIf
                Set-ADUser $SID -Replace @{MailNickName = $emailsuffix; targetAddress = $TargetAddress} #-WhatIf
            }
            Catch
            {
                Write-Output "Error for: $validateRoom.Name New name: $DisplayName"
                $_.Exception.Message 
                $_.Exception.ItemName 
                $_.InvocationInfo.MyCommand.Name 
                $_.ErrorDetails.Message
                $_.InvocationInfo.PositionMessage 
                $_.CategoryInfo.ToString()
                $_.FullyQualifiedErrorId 
                Write-Output "-----------------------------"
                Write-Output ""
            }
            #####################  Populate ProxyAddresses  #####################
            Try
            {
                $RoomProxies = (get-aduser -Filter "Name -like '$DisplayName'" -Properties * | select-object @{"name"="proxyaddresses";"expression"={$_.proxyaddresses}}).proxyaddresses
                $EmailSMTP = "SMTP:" + $Email
                $TargetSuffix = ($TargetAddress.ToLower()) -replace "smtp:",""				
										
                $NewRoom = Get-ADUser -Filter "Name -like '$DisplayName'" -Properties *
                $NewRmSID = $NewRoom.SID.Value

                foreach($RoomProxy in $RoomProxies)
                {
                    $ProxyEmail = (($RoomProxy.ToLower()) -replace "smtp:", "")

                    #Remove Primary SMTP
                    if($RoomProxy -cmatch “^[A-Z]:*”)
                    {
                        $OldPrimaryEmail = (($RoomProxy.ToLower()) -replace "smtp:", "")

                        if($OldPrimaryEmail -like $ProxyEmail)
                        {
                            Set-ADUser $NewRmSID -Remove @{proxyAddresses = "SMTP:$OldPrimaryEmail"}
                        }
                    }
            
                    #Remove current email if it already exists
                    if($Email -like $ProxyEmail)
                    {
                        Set-ADUser $NewRmSID -Remove @{proxyAddresses = "smtp:$Email"}
                    }

                    #Remove target email if it already exists
                    if($TargetSuffix -like $ProxyEmail)
                    {
                        Set-ADUser $NewRmSID -Remove @{proxyAddresses = "smtp:$TargetSuffix"}
                    }
                }

                #Add the new Information into ProxyAddresses
                Set-ADUser $NewRmSID -Add @{proxyAddresses = "$EmailSMTP"}
                Set-ADUser $NewRmSID -Add @{proxyAddresses = "smtp:$OldPrimaryEmail"}
                Set-ADUser $NewRmSID -Add @{proxyAddresses = "smtp:$TargetSuffix"}
            }
            Catch
            {
                Write-Output "Error for: $validateRoom.Name New name: $DisplayName"
                $_.Exception.Message 
                $_.Exception.ItemName 
                $_.InvocationInfo.MyCommand.Name 
                $_.ErrorDetails.Message
                $_.InvocationInfo.PositionMessage 
                $_.CategoryInfo.ToString()
                $_.FullyQualifiedErrorId 
                Write-Output "-----------------------------"
                Write-Output ""
            }
        }
    }
    Else
    {
        #####################  New Resource Creation  #####################

        Write-Output "Creating a new Resource Room '$DisplayName'"
											
        $CopiedProfileUser = Get-ADUser -Filter "Name -like 'GLE Conf Rm B'" -Properties *
        $RmPassword = ConvertTo-SecureString "!@#$SomeSecurePassword*&^" -AsPlainText -Force
                                          
        Try
		{       
			New-RemoteMailbox -Name $DisplayName -Password $RmPassword -Room -UserPrincipalName $Email -Confirm:$False -DisplayName $DisplayName -LastName $DisplayName -OnPremisesOrganizationalUnit "OU=Resource-Rooms,DC=domain,DC=com" -PrimarySmtpAddress $Email -ResetPasswordOnNextLogon $False -SamAccountName $SamAccountName | Out-Null
		}
		Catch
		{
			Write-Output "Error for: $DisplayName"
			$_.Exception.Message 
			$_.Exception.ItemName 
			$_.InvocationInfo.MyCommand.Name 
			$_.ErrorDetails.Message
			$_.InvocationInfo.PositionMessage 
			$_.CategoryInfo.ToString()
			$_.FullyQualifiedErrorId 
			Write-Output "-----------------------------"
			Write-Output ""
		}

		#Try and make sure all variables are correct
		Try
		{
			#Let's sleep for a couple seconds to confirm the resource was created
			Start-Sleep -s 3 #(2 is fine as well)

			$NewRoom = Get-ADUser -Filter "Name -like '$DisplayName'" -Properties *
			Set-ADUser $NewRoom -Office $Office -PasswordNeverExpires $True #-WhatIf
		}
		Catch
		{
			Write-Output "Error for: $DisplayName"
			$_.Exception.Message 
			$_.Exception.ItemName 
			$_.InvocationInfo.MyCommand.Name 
			$_.ErrorDetails.Message
			$_.InvocationInfo.PositionMessage 
			$_.CategoryInfo.ToString()
			$_.FullyQualifiedErrorId 
			Write-Output "-----------------------------"
			Write-Output ""
		}
			
		#####################     Let's sync with O365      #####################

		#Start-ADSyncSyncCycle -PolicyType Delta
		Invoke-Command -ComputerName "test-aad-01.domain.com" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} 

		Write-Output "Starting an O365 Sync now. Please allow up to 15 mins before the changes sync in Office365."

		#Let's sleep for a minute to give the sync enough time
		Start-Sleep -s 60        
		
        #####################  Setup Conf Room Permissions  #####################	
		
		Try
        {
            $userUPN = "365admin@domain.onmicrosoft.com" 
			$AESKeyFilePath = (Split-Path $script:MyInvocation.MyCommand.Path) + "\AES.key"
			$SecurePwdFilePath =  (Split-Path $script:MyInvocation.MyCommand.Path) + "\AESpassword.txt"
			$AESKey = Get-Content -Path $AESKeyFilePath -Force
			$securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

			#create a new psCredential object with required username and password
			$UserCredential = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)

			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

			Import-PSSession $O365 -DisableNameChecking

			$NewRoom = Get-ADUser -Filter "Name -like '$DisplayName'" -Properties *
			$ResourcePermission = $NewRoom.name + ":\Calendar"
			
			### Check User Groups ###
			$users = ((Get-MailboxFolderPermission $ResourcePermission).user).displayname
			
			foreach ($user in $users)
			{
				if($user -eq "Default")
				{
					Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights LimitedDetails
				}

				if($user -eq "Anonymous")
				{
					Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights None
				}

				if($user -eq "Room Schedulers")
				{
					Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights Author
				}

				if($user -eq "Room Schedulers Editor")
				{
					Set-MailboxFolderPermission -Identity $ResourcePermission -User $user -AccessRights Editor
				}
			}
			
			if($users -notcontains "Default")
			{
				Add-MailboxFolderPermission -Identity $ResourcePermission -User "Default" -AccessRights LimitedDetails
			}

			if($users -notcontains "Anonymous")
			{
				Add-MailboxFolderPermission -Identity $ResourcePermission -User "Anonymous" -AccessRights None
			}

			if($users -notcontains "Room Schedulers")
			{
				Add-MailboxFolderPermission -Identity $ResourcePermission -User "Room Schedulers" -AccessRights Author
			}

			if($users -notcontains "Room Schedulers Editor")
			{
				Add-MailboxFolderPermission -Identity $ResourcePermission -User "Room Schedulers Editor" -AccessRights Editor
			}
        }
        Catch
        {
            Write-Output "Error for: $DisplayName"
            $_.Exception.Message 
            $_.Exception.ItemName 
            $_.InvocationInfo.MyCommand.Name 
            $_.ErrorDetails.Message
            $_.InvocationInfo.PositionMessage 
            $_.CategoryInfo.ToString()
            $_.FullyQualifiedErrorId 
            Write-Output "-----------------------------"
            Write-Output ""        
        }
    }
}

#Close out our sessions once we're done using it

if($null -ne $O365)
{
    Remove-PSSession -Session $O365
}

if($null -ne $MBXSession)
{
    Remove-PSSession -Session $MBXSession
}

if($null -ne $session)
{
    Remove-PSSession -Session $session
} 
