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

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

if(!(Get-Module -ListAvailable -Name MSOnline))
{
    Install-Module -Name MSOnline -Scope CurrentUser -Force -Confirm:$False
} 

#Quick way to see if we are connected to the MSOL service is to run a simple query. If it doesn't return NULL, then we are fine and don't need to load it again
if(!(Get-MsolUser -SearchString "Task Scheduler" -ErrorAction SilentlyContinue))
{
    $userUPN = "admin@domain.onmicrosoft.com" 
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

$SharedMBXs = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited

foreach($SharedMBX in $SharedMBXs)
{
    $UPN = $SharedMBX.UserPrincipalName
    $user = Get-ADUser -Filter "UserPrincipalName -like '$UPN'" -Properties *
    $DN = $user.DistinguishedName

    if($DN -notlike "*OU=Shared Mailbox*")
    {
        $DN

        Try
        {
            $DN | Move-ADObject -TargetPath "OU=Shared Mailbox,DC=domain,DC=com" -ErrorAction Stop
        }
        Catch
        {
            $UPN
        }
    }
}

#Close out our sessions once we're done using it
Remove-PSSession *
