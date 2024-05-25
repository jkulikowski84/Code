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

$Export = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Info.csv"

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

$path = "OU=DEF,OU=ABC,DC=domain,DC=com"

get-aduser -SearchBase $path -filter 'enabled -eq $true' -properties * | ? {$_.DistinguishedName -notlike "*,OU=xyz,*"} | Where-Object {$_.enabled -like "true" -and $_.mail -ne "something@domain.com" -and $_.mail -notlike "" -and $_.office -notlike "vendor" -and $_.title -notlike "service account" -and $_.company -notlike ""} | select-object name, givenname, surname, @{"name"="Apple Email";"expression"={(($_.ProxyAddresses | Where-Object ({($_ -like "*.*@domain.com") }) | Select-Object -First 1) -replace "smtp:", "") -join ';'} } | Export-Csv $Export

