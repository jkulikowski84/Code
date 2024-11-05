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

if($NULL -eq $DMZCred)
{
    #Store DMZ credentials
    $DMZCred = Get-Credential( (whoami) -replace "domain","domain.local" )
}


$DMZServers = Invoke-Command -ComputerName "DMZ-AD.domain.local" -Credential $DMZCred -ScriptBlock { ((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique }

## Used for Progress Bar
$data = @{
    Count = $(($DMZServers | measure).count)
	Done = 0
}

$DMZServers | split-pipeline -count 64 -Variable data, DMZCred -ErrorAction SilentlyContinue {
    Process
    {
        $DMZServerName = $_ + ".domain.local"

        Try
        {
            $ipv6Addresses = Invoke-Command -Credential $DMZCred  -ComputerName $DMZServerName -ScriptBlock { Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue }

            #if ($ipv6Addresses.Count -gt 0)
            if ($ipv6Addresses.Enabled -eq "True") 
            {
                $DMZServerName
            }
        }
        Catch
        {}

        # Count how far along we are
        $done = ++$data.Done

        # show progress
	    Write-Progress -Activity "Done $done" -Status Processing -PercentComplete (100*$done/$data.Count)
    }
}