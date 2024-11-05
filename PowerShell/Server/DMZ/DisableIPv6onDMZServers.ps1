CLS

#List the modules we want to load per line
$LoadModules = @'
SplitPipeline
'@.Split("`n").Trim()

ForEach($Module in $LoadModules) 
{
    #Check if Module is Installed
    if($NULL -eq (Get-Module -ListAvailable $Module))
    {
        #Install Module
        Install-Module -Name $Module -Scope CurrentUser -Confirm:$False -Force
        #Import Module
        Import-Module $Module
    }

    #Check if Module is Imported
    if($NULL -eq (Get-Module -Name $Module))
    {
        #Install Module
        Import-Module $Module
    }
}

$DMZServers = @'
DMZ1.domain.local
DMZ2.domain.local
DMZ3.domain.local
DMZ4.domain.local
DMZ5.domain.local
DMZ6.domain.local
DMZ7.domain.local
DMZ8.domain.local
DMZ9.domain.local
'@.Split("`n").Trim()

$data = @{
    Count = $(($DMZServers | measure).count)
	Done = 0
}

if($NULL -eq $DMZCred)
{
    #Store DMZ credentials
    $DMZCred = Get-Credential( (whoami) -replace "domain","domain.local" )
}

$DMZServers | split-pipeline -count 64 -Variable data, DMZCred -ErrorAction SilentlyContinue {
    Process
    {
        Try
        {
            Invoke-Command -Credential $DMZCred  -ComputerName $_ -ScriptBlock { Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue | Disable-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue }
        }
        Catch
        {}
        
        # Count how far along we are
        $done = ++$data.Done

        # show progress
	    Write-Progress -Activity "Done $done" -Status Processing -PercentComplete (100*$done/$data.Count)
    }
}