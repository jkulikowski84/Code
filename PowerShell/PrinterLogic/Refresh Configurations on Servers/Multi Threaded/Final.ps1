CLS

#Start Timestamp
"[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)

#Global Variables
$Processes = $NULL

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

# Script to run in each thread.
[System.Management.Automation.ScriptBlock]$ScriptBlock = {
	Param ( $PrintServer )
	
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

function Invoke-AsyncJob
{
    $AllJobs = New-Object System.Collections.ArrayList

    $HostRunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(20,100,$Host)

    $HostRunspacePool.Open()

    $searcher = [adsisearcher]::new()

    #Sort in ascending by Name
    $searcher.Sort.PropertyName = "name"

    #Search root is the OU level we want to search at
    $searcher.SearchRoot = [adsi]"LDAP://OU=Servers,DC=domain,DC=com"

    #Make this any non zero value to expand the default result size.
    $searcher.PageSize = 100

    #Filter Computers only and enabled
    $searcher.Filter = "(&(objectCategory=computer)(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"

    #List the properties we are interested in
    $searcher.PropertiesToLoad.AddRange('name')

    #Now output the results with the exception of filtering out any machines in the citrix farms (we're not interested in those)
    $Servers = Foreach ($Computer in $searcher.FindAll() | where { (($_.properties["Name"] -like "*EEPS*") -OR ($_.properties["Name"] -like "*EPMS*")) -AND ($_.properties["Name"] -notlike "testserver0002") }){ ($Computer.Properties).name }

    #---------------

    foreach($Server in $Servers)
    {

        $asyncJob = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddArgument($Server)
        
        $asyncJob.RunspacePool = $HostRunspacePool

        $asyncJobObj = @{ JobHandle   = $asyncJob;
                          AsyncHandle = $asyncJob.BeginInvoke() }

        $AllJobs.Add($asyncJobObj) | Out-Null
    }
    $ProcessingJobs = $true

    Do 
    {
        $CompletedJobs = $AllJobs | Where-Object { $_.AsyncHandle.IsCompleted }

        if($null -ne $CompletedJobs)
        {
            foreach($job in $CompletedJobs)
            {
                $job.JobHandle.EndInvoke($job.AsyncHandle)

                $job.JobHandle.Dispose()

                $AllJobs.Remove($job)
            } 
        } 
        else 
        {
            if($AllJobs.Count -eq 0)
            {
                $ProcessingJobs = $false
            } 
            else 
            {
                Start-Sleep -Milliseconds 1000
            }
        }
    } 
    While ($ProcessingJobs)

    $HostRunspacePool.Close()
    $HostRunspacePool.Dispose()
} 

Write-Host "All Servers had the PrinterLogic Service restarted."
Write-Host "Please give 1-5 mins for the new printer settings to update."

Invoke-AsyncJob

Write-Host " "

#End Timestamp
"[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
