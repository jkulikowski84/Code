CLS

#Load Active Directory Module remotely if it's not already loaded
$SessionsRunning = get-pssession

if($SessionsRunning.ComputerName -like "*bar-mbx-01*")
{
    #If session is running we don't need to do anything
}
else
{
    $userUPN = "SVC-AADConnect" 
    $AESKeyFilePath = ($pwd.ProviderPath) + "\MBX-Module\aeskey.txt"
    $SecurePwdFilePath =  ($pwd.ProviderPath) + "\MBX-Module\password.txt"
    $AESKey = Get-Content -Path $AESKeyFilePath -Force
    $securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

    #create a new psCredential object with required username and password
    $adminCreds = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)

    #$MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://tst-mbx-01.domain.com/PowerShell/ -Credential $adminCreds -ErrorAction Stop
    $MBXSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://tst-mbx-01.domain.com/PowerShell/ -Authentication Kerberos -Credential $adminCreds

    #Import-PSSession $MBXSession -AllowClobber -DisableNameChecking
    Import-PSSession $MBXSession -DisableNameChecking | Out-Null
	
	#Invoke-Command -Session $MBXSession -ScriptBlock {Get-Module ActiveDirectory} 
    #Import-Module -Name ActiveDirectory -PSSession $MBXSession
	
	#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
