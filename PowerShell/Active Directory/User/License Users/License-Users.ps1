CLS

Function LicenseOfficeUser
{

    if(!(Get-Module -ListAvailable -Name MSOnline))
	{
		Install-Module -Name MSOnline -Scope CurrentUser -Force -Confirm:$False
	} 
<#
	if (Get-Module -ListAvailable -Name MSOnline) 
    {
            #Write-Host "Module exists"
    } 
    else 
    {
        Install-Module -Name MSOnline -Force 
    }
#>
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
        
        Connect-MsolService -Credential $adminCreds
    }

	#This is the list of emails we will be importing from. This file gets created when you run Part1
	$EmailFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\emails.txt"

	#This is the license we will be assigning to the user. The "EnterprisePack" license if Office365 E3
	$license = (Get-MsolAccountSku).AccountSkuId | Where-Object {$_ -like "domain:ENTERPRISEPACK" }

    #Now lets read the emails from the emails.txt file and license the new users for Office365
    Get-Content $EmailFile | ForEach-Object {
        $useremail = $_
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuID $license

        Set-MsolUser -UserPrincipalName $useremail -UsageLocation 'US' -ErrorAction SilentlyContinue
        Set-MsolUserLicense -UserPrincipalName $useremail -AddLicenses $license -LicenseOptions $LicenseOptions -ErrorAction SilentlyContinue
    }
}

#License the User in Office
LicenseOfficeUser