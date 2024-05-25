CLS

function Get-ModuleADSync
{
    #Add the necessary modules from the server
    Try
    {
        $AADsession = New-PSSession -ComputerName "gle-aad-01.domain.com" -Authentication Kerberos -ErrorAction Stop
    }
    Catch
    {
        $userUPN = "SVC-AADConnect" 
        $AESKeyFilePath = ($pwd.ProviderPath) + "\AAD-Module\AESkey.txt"
        $SecurePwdFilePath =  ($pwd.ProviderPath) + "\AAD-Module\password.txt"
        $AESKey = Get-Content -Path $AESKeyFilePath -Force
        $securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

        #create a new psCredential object with required username and password
        $adminCreds = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)

        #$credentials = Get-Credential
        $AADsession = New-PSSession -ComputerName "gle-aad-01.domain.com" -Credential $adminCreds
    }

    Invoke-Command -Session $AADsession -ScriptBlock {Get-Module ADSync} 
    Import-Module -Name ADSync -PSSession $AADsession | Out-Null
}

if(!(Get-Module -ListAvailable -Name "ADSync"))
{
    Get-ModuleADSync
}