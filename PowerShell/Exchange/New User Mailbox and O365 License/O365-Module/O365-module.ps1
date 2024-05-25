CLS

function Get-ModuleMSOnline
{
    if(!(Get-Module -ListAvailable -Name MSOnline))
    {
        Install-Module -Name MSOnline -Scope CurrentUser -Force -Confirm:$False
    }

    if(!(Get-MsolUser -SearchString "Task Scheduler" -ErrorAction SilentlyContinue))
    {
        $userUPN = "365admin@domain.onmicrosoft.com" 
        $AESKeyFilePath = ($pwd.ProviderPath) + "\O365-Module\aeskey.txt"
        $SecurePwdFilePath =  ($pwd.ProviderPath) + "\O365-Module\password.txt"
        $AESKey = Get-Content -Path $AESKeyFilePath -Force
        $securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

        #create a new psCredential object with required username and password
        $adminCreds = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)
        
        #$O365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $adminCreds -Authentication Basic -AllowRedirection
        $O365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $adminCreds -Authentication Basic -AllowRedirection

        #Import-PSSession $O365 -AllowClobber -DisableNameChecking
	    Import-PSSession $O365 -DisableNameChecking | Out-Null

        Connect-MsolService -Credential $adminCreds
        Connect-AzureAD -Credential $adminCreds
    }
}

#Load Active Directory Module remotely if it's not already loaded
Get-ModuleMSOnline