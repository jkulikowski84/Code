CLS
		
#Global Variables
$EmailFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\emails.txt"
$OldPref = $global:ErrorActionPreference
$global:ErrorActionPreference = 'Stop'

#Remove the email file if it already exists
if([System.IO.File]::Exists($EmailFile))
{
	remove-item $EmailFile -Force
}
	
function Get-ModuleAD
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

Function LicenseOfficeUser
{
    if(!(Get-Module -ListAvailable -Name MSOnline))
	{
		Install-Module -Name MSOnline -Scope CurrentUser -Force -Confirm:$False
	} 

    #Quick way to see if we are connected to the MSOL service is to run a simple query. If it doesn't return NULL, then we are fine and don't need to load it again
    if(!(Get-MsolUser -SearchString "Task Scheduler" -ErrorAction SilentlyContinue))
    {
        $userUPN = "365admin@domain.onmicrosoft.com" 
        $AESKeyFilePath = (Split-Path $script:MyInvocation.MyCommand.Path) + "\AES.key"
		$SecurePwdFilePath =  (Split-Path $script:MyInvocation.MyCommand.Path) + "\AESpassword.txt"
        $AESKey = Get-Content -Path $AESKeyFilePath -Force
		$securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

		#create a new psCredential object with required username and password
        $adminCreds = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)
        
		$O365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $adminCreds -Authentication Basic -AllowRedirection

		Import-PSSession $O365 -AllowClobber
	
        Connect-MsolService -Credential $adminCreds
		
		#Clear screen
		CLS
    }

	#This is the list of emails we will be importing from. This file gets created when you run Part1
	#$EmailFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\emails.txt"

	#This is the license we will be assigning to the user. The "EnterprisePack" license if Office365 E3
	$license = (Get-MsolAccountSku).AccountSkuId | Where-Object {$_ -like "domain:ENTERPRISEPACK" }
	$CheckLastUser = Get-Content $EmailFile -Tail 1

	#Before we start, we need to make sure the last user we exported has synced up with O365 already, otherwise there will be errors
	try
    {
        while (-not $GetMailbox)
        {
            Try 
            {
                Invoke-Command -ComputerName "gle-aad-01.domain.com" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
                Sleep 30

                $GetMailbox = Get-Mailbox $CheckLastUser                     
            }
            Catch {sleep 1}
        }
    }
    catch {}
	
    #Now lets read the emails from the emails.txt file and license the new users for Office365
    Get-Content $EmailFile | ForEach-Object {
        $useremail = $_
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuID $license

        Set-MsolUser -UserPrincipalName $useremail -UsageLocation 'US' -ErrorAction SilentlyContinue
        Set-MsolUserLicense -UserPrincipalName $useremail -AddLicenses $license -LicenseOptions $LicenseOptions -ErrorAction SilentlyContinue
    }
}

function CreateNewUser
{
    $worksheet = "Sheet1"

    #The file we will be reading from
    $ExcelFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Final.xlsx"

    $Import = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

    foreach ($User in $Import)
    {
        #$Set our Variables to NULL
        $FirstName = $LastName = $MiddleName = $DisplayName = $Email = $emailsuffix = $SamAccountName = $CopyFromProfile = $NULL

        #Define our variables
        $FirstName = ($User."First Name").Trim()
        $LastName = ($User."Last Name").Trim()
        $MiddleName = $User."Middle Name"

        if($NULL -ne $MiddleName)
        {
            $MiddleName = $MiddleName.Trim()
            $DisplayName = $FirstName + " " + $MiddleName + " " + $LastName
        }
        else
        {
            $DisplayName = $FirstName + " " + $LastName
        }

        $Email = ($User."Email Address").Trim()
        $emailsuffix = ($Email.Substring(0, $Email.IndexOf('@'))).Trim()
        $SamAccountName = $emailsuffix[0..19] -join ""

        $CopyFromProfile = $User."Copy From Profile (username)"

        if($NULL -ne $CopyFromProfile)
        {
            $CopyFromProfile = $CopyFromProfile.Trim()
        }
    
	    #####################  New User Creation  #####################

        #To run this successfully, the account running must be part of the Mailbox Import/Export role on the On-Prem Exchange Server
        #New-ManagementRoleAssignment -Role "Mailbox Import Export" –User <domain\user>
		#Also need to add the service account to the Exchange MgmT Role "Organization Management" in AD

        #Zero out our Variables
        $validateUser = $CopyProfile = $UserPassword = $Enabled = $ADUser = $NULL

	    #First make sure the Resource Room doesn't already exist
	    $validateUser = Get-ADUser -Filter "Name -like '$DisplayName'" -Properties *
	
	    if($NULL -eq $validateUser)
	    {
		    Write-Output "Creating a new User Account '$DisplayName'"
						
            if($NULL -ne $CopyFromProfile)
            {
                $CopyProfile = Get-ADUser -Filter "SamAccountName -like '$CopyFromProfile'" -Properties *
            }

            #Stock User Password
		    $UserPassword = ConvertTo-SecureString "Welcome1" -AsPlainText -Force
			$Enabled = $True
								  
		    Try
		    {   
                #####################  Variables From Another User Profile  #####################
                if($NULL -ne $CopyProfile)
                {
                    #Zero out our Variables
                    $DN = $Title = $Description = $Office = $City = $Zip = $Co = $Dept = $Company = $Mgr = $Manager = $AccountExpiration = $NULL

                    $DN = $CopyProfile.DistinguishedName -replace '^cn=.+?(?<!\\),'
                    $Title = $CopyProfile.title
                    $Description = $Title
                    $Office = $CopyProfile.physicalDeliveryOfficeName
                    $City = $CopyProfile.City
                    $State = $CopyProfile.State
                    $Zip = $CopyProfile.PostalCode
                    $Co = $CopyProfile.Co
                    $Dept = $CopyProfile.Department
                    $Company = $CopyProfile.Company
                    $Mgr = (($CopyProfile.Manager -split ',*..=')[1])
                    $Manager = (Get-ADUser -Filter "DisplayName -like '$Mgr'" -Properties *).SamAccountName

                    #If it's a Volunteer, set a 6 month expiration
                    if($Title -eq "Volunteer")
                    {
                        #There are approximately 180 days in 6 months
                        $AccountExpiration = (New-TimeSpan -Days 180)
                    }
                }
                #####################  Create a New User Profile  #####################
                else
                {
                    #Zero out our Variables
                    $DNInput = $LDAPPath = $seek = $Result = $DN = $NULL

                    #Default OU where New Users are placed if they're not copied from another user's profile
                    do
                    {
                        $DNInput = Read-Host -Prompt 'What OU do you want to put the new User in? (For example GLE-Users)'

                        $LDAPPath = "LDAP://dc=domain,dc=com"
                        $seek = [System.DirectoryServices.DirectorySearcher]$LDAPPath
                        $seek.Filter = “(&(name=$DNInput)(objectCategory=organizationalunit))”
                        $Result = (($seek.FindOne()).Path) -replace "LDAP://",""

                    }while(!$Result)

                    if($NULL -ne $Result)
                    {
                        $DN = $Result
                    }
                }

                #Middle name Initials if they have any
                if($NULL -ne $MiddleName)
                {
                    $Initial = $MiddleName[0] -join ""
                }
                else
                {
                    $Initial = ""
                }

                #Create the new User Account
			    New-RemoteMailbox -Name $DisplayName -Password $UserPassword -UserPrincipalName $Email -Confirm:$False -DisplayName $DisplayName -FirstName $FirstName -LastName $LastName `
                -Initials $Initial -OnPremisesOrganizationalUnit $DN -PrimarySmtpAddress $Email -ResetPasswordOnNextLogon $False -SamAccountName $SamAccountName
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
            
            #####################  Make sure account created successfully before continuing  #####################
            while (-not $ADUser) 
            {
                try 
                {
                    $ADUser = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties *
                }
                catch {Write-Output "User not found in AD yet. Checking again in 5 seconds"; sleep 5}
            }

            #####################  Populate Fields from our Copied Profile to the new Account  #####################

            if($NULL -ne $CopyProfile)
            {
		        Try
		        {
                    #Add the user group memberships from the copied profile
                    $CopyProfile.memberof | add-adgroupmember -members $ADUser.SamAccountName -ErrorAction SilentlyContinue

				    Set-ADUser $ADUser -Title $Title -Description $Description -Office $Office -City $City -State $State -PostalCode $Zip -Enabled $Enabled -Department $Dept -Company $Company -Manager $Manager
                    Set-ADUser -Identity $ADUser -Replace @{ c = "US"; co = "USA" }

                    #We expire Volunteer Accounts 6 months after creation
                    if($Title -eq "Volunteer")
                    {
                        #There are approximately 180 days in 6 months
                        Set-ADAccountExpiration $ADUser -TimeSpan $AccountExpiration
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
            if($Title -ne "Volunteer")
            {
                Set-ADUser -Identity $ADUser -Replace @{ extensionAttribute2 = "DynamicDistro" }
            }
			
			#Export the email to a file. We will use this later to license the new accounts
			$Email | Out-File -Append -FilePath $EmailFile
        }

        #####################  User Account already Exists  #####################

		Else
		{
			Write-Output "The User already exists!"
		}
<#
        #Clear all our variables
        $Vars =("FirstName","LastName","MiddleName","DisplayName","Email","emailsuffix","SamAccountName","CopyFromProfile","validateUser","CopyProfile","DN","Title","Description","Office","City","State","Zip","Co","Dept", `
        "$Company","Mgr","Manager","AccountExpiration","DNInput","LDAPPath","Result","ADUser","Initial")
        
        for ($i = 0; $i -le ($Vars.length - 1); $i++) 
        {
            Try
            {
                Clear-Variable $Vars[$i] -ErrorAction SilentlyContinue
            }
            Catch
            {

            }
        }
#>
    }

    #####################     Let's sync with O365      #####################
<#
    #Start-ADSyncSyncCycle -PolicyType Delta
    Invoke-Command -ComputerName "gle-aad-01.domain.com" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

    Write-Output "Starting an O365 Sync now. Please allow up to 15 mins before the changes sync in Office365."
	Write-Output ""
    #Let's sleep for a minute to give the sync enough time
    Start-Sleep -s 30
#>
	#####################     License Users      #####################
	LicenseOfficeUser

    #Close out our sessions once we're done using it
	Remove-PSSession *
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

if($SessionsRunning.ComputerName -like "*tst-mbx-01*")
{
    #If session is running we don't need to do anything
}
else
{
    $MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://tst-mbx-01.domain.com/PowerShell/ -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $MBXSession -AllowClobber -DisableNameChecking
}

#Clear screen again
CLS

CreateNewUser
