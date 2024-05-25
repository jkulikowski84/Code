CLS

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
        $ADsession = New-PSSession -ComputerName $hostname -Authentication Kerberos -ErrorAction Stop
    }
    Catch
    {
        $userUPN = "SVC-AADConnect" 
        $AESKeyFilePath = ($pwd.ProviderPath) + "\AD-Module\AESkey.txt"
        $SecurePwdFilePath =  ($pwd.ProviderPath) + "\AD-Module\password.txt"
        $AESKey = Get-Content -Path $AESKeyFilePath -Force
        $securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

        #create a new psCredential object with required username and password
        $adminCreds = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)

        #$credentials = Get-Credential
        $ADsession = New-PSSession -ComputerName $hostname -Credential $adminCreds
    }

    Invoke-Command -Session $ADsession -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $ADsession | Out-Null
}

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}