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
        $credentials = Get-Credential
        $session = New-PSSession -ComputerName $hostname -Credential $credentials
    }

    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

#----------------------------------------------------------------------------------------------------------------
<#
    The worksheet variable will need to be modified before running this script. 
    Whatever the name of the worksheet is that you want to import data from, type that in below.
#>

function CreateNewResourceRm
{
    $worksheet = "Sheet1"

    #The file we will be reading from
    $ExcelFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Final.xlsx"

    $Import = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

    foreach ($Room in $Import)
    {
		#Zero out our Variables
		DisplayName = Email = Office = emailsuffix = SamAccountName = validateRoom = ADUser = $NULL
		
        #Define our variables
        $DisplayName = ($Room."New Room Name").Trim()
        $Email = ($Room."Email Address").Trim()
        $Office = ($Room."Office").Trim()
        $emailsuffix = ($Email.Substring(0, $Email.IndexOf('@'))).Trim()
        $SamAccountName = $emailsuffix[0..19] -join ""   
    
	    #####################  New Resource Creation  #####################

        #To run this successfully, the account running must be part of the Mailbox Import/Export role on the On-Prem Exchange Server
        #New-ManagementRoleAssignment -Role "Mailbox Import Export" –User <domain\user>
		#Also need to add the service account to the Exchange MgmT Role "Organization Management" in AD

	    #First make sure the Resource Room doesn't already exist
	    $validateRoom = Get-ADUser -Filter "Name -like '$DisplayName'" -Properties *
	
	    if($NULL -eq $validateRoom)
	    {
		    Write-Output "Creating a new Resource Room '$DisplayName'"
			
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
				while (-not $ADUser) 
				{
					try 
					{
						$ADUser = Get-ADUser -Filter "DisplayName -like '$DisplayName'" -Properties *
					}
					catch {Write-Output "Room not found in AD yet. Checking again in 5 seconds"; sleep 5}
				}

				Set-ADUser $ADUser -Office $Office -PasswordNeverExpires $True 
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
			Invoke-Command -ComputerName "test-aad-01.domain.com" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} #| Out-NULL

			Write-Output "Starting an O365 Sync now. Please allow up to 15 mins before the changes sync in Office365."
			Write-Output "Make sure to run the script to update the permissions for the new Resource room after the resource syncs with Office 365."
			
            #Let's sleep for a minute to give the sync enough time
            Start-Sleep -s 60
        }
		Else
		{
			Write-Output "The Resource Room already exists!"
		}
    }

    #Close out our sessions once we're done using it
	if($null -ne $AADsession)
    {
        Remove-PSSession -Session $AADsession
    }
	
    if($null -ne $MBXSession)
    {
        Remove-PSSession -Session $MBXSession
    }

    if($null -ne $session)
    {
        Remove-PSSession -Session $session
    } 
}

#======================================
#Main

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

#Check if we have a session open right now
$SessionsRunning = get-pssession

if($SessionsRunning.ComputerName -like "*bar-mbx-01*")
{
    #If session is running we don't need to do anything
}
else
{
    $MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://test-mbx-01.domain.com/PowerShell/ -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $MBXSession -AllowClobber -DisableNameChecking
}

#Clear screen again
CLS

CreateNewResourceRm
