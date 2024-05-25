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

#Clear screen again
CLS

# Variables relative to environment
$ComputersOuDn = 'OU=COMPUTERS,OU=COMPUTER-SYSTEMS,DC=domain,DC=com' # Distinguished Name for OU where computer objects to check are located
$TpmDevicesCnDn = 'CN=TPM Devices,DC=domain,DC=com' # Distinguished Name for Contoiner where TPM objects are located

# Get AD computers with non-null TPM attributes
$TpmComputers = Get-ADComputer -SearchBase $ComputersOuDn -Filter {msTPM-TpmInformationForComputer -ne "$null"} -Properties name,msTPM-TpmInformationForComputer |
    Select-Object -Property @{Name='Name';Expression={$_."name"}},@{Name='TpmInfo';Expression={$_."msTPM-TpmInformationForComputer"}}
    
# Get AD objects in the TPM Devices container
$TpmObjects = Get-ADObject -SearchBase $TpmDevicesCnDn -Filter {objectClass -eq "msTPM-InformationObject"} -Properties distinguishedName,name,msTPM-OwnerInformation | 
    Select-Object -Property distinguishedName,@{Name='Name';Expression={$_."name"}},@{Name='OwnerInfo';Expression={$_."msTPM-OwnerInformation"}}

# Iterate through computers and build a new Computers object with the required values
$Computers = @()
$OrphanComputers = @()
$OrphanTpmObjects = @()

ForEach ($TpmComputer in $TpmComputers)
{
    # Get matching objects in the TpmObjects array with attributes matching current computer in TpmComputers array
    $TpmObject = $TpmObjects.Where({$_.distinguishedName -eq $TpmComputer.TpmInfo})

    If (!$TpmObject)
		{$OrphanComputers += $TpmComputer}
    Else
    {
		# Build object properties from computer and TPM object properties
		$Properties = @{}
		$Properties.Name = $TpmComputer.Name;
		$Properties.TPMInfo = $TpmObject.Name
		$Properties.OwnerInfo = $TpmObject.OwnerInfo
			
		# Add new object to Computers array
		$Computers += New-Object -TypeName PSObject -Property $Properties
    }
}

# Find orphan TPM Objects
ForEach ($TpmObject in $TpmObjects)
{
    # Get matching objects in the TpmObjects array with attributes matching current computer in TpmComputers array
    $TpmComputer = $TpmComputers.Where({$_.TpmInfo -eq $TpmObject.distinguishedName})
    If (!$TpmComputer)
    {
        $OrphanTpmObjects += $TpmObject
        
        #Pause
        Remove-ADObject $TpmObject.distinguishedName -Confirm:$False
    }
}

# Outputs
If ($Computers.Count -gt 0)
{
    Write-Output "`r`nComputers with TPM information ($($Computers.Count)): "
    Write-Output $Computers | Select-Object Name,TPMInfo,OwnerInfo | Sort-Object Name
}

If ($OrphanComputers.Count -gt 0)
{
    Write-Output "`r`nOrphan computers ($($OrphanComputers.Count)): "
    Write-Output $OrphanComputers | Select-Object Name | Sort-Object Name
}

If ($OrphanTpmObjects.Count -gt 0)
{
    Write-Output "`r`nOrphan TPM objects ($($OrphanTpmObjects.Count)): "
    Write-Output $OrphanTpmObjects | Select-Object Name,OwnerInfo | Sort-Object Name
}