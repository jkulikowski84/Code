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

#-------------------------------------------------------------

#Global Variable(s)
$Processes = $NULL

#Get all Print Servers
$PrintServers = Get-ADComputer -Filter * -SearchBase "OU=Servers,DC=domain,DC=com" | Where-Object { ($_.Name -like "*WVS0*EP*S*00*") -AND ($_.Name -notlike "*EPRS*") -AND ($_.Name -notlike "TWVS01EEPS7002") } | Select -Property Name | Sort Name

foreach($PrintServer in $PrintServers.Name)
{
    #Check to see if the process(es) is/are running
	$PrinterInstallerClients = Get-Process -ComputerName $PrintServer PrinterInstallerClient* -ErrorAction SilentlyContinue

	#Close out of the processes if they are running
	if($PrinterInstallerClients)
	{
        Write-Host "Recycling services on $PrintServer"

		#First we need to format the process so it shows correctly
		$Processes = @()
		foreach($PrinterInstallerClient in $PrinterInstallerClients)
		{
			$Processes += ($PrinterInstallerClient.ProcessName).split('_')[0] + ".exe"
		} 

		#Sort the process in descending order so we can properly exit the processes starting with PrinterInstallerClientLauncher
		$Processes = ($Processes | sort -Descending)

		foreach($Process in $Processes)
		{
			taskkill /S $PrintServer /IM $Process /f | Out-Null
		}
	}

	#Start the PrinterInstallerClientLauncher
	sc.exe \\$PrintServer start PrinterInstallerLauncher | Out-Null

    Clear-Variable -Name "Processes"
}

CLS

Write-Host "All Servers had the PrinterLogic Service restarted."
Write-Host "Please give 1-5 mins for the new printer settings to update."
