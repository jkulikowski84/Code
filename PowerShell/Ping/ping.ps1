CLS


<#

https://stackoverflow.com/questions/37954004/how-to-multithread-powershell-ping-script
https://superuser.com/questions/805621/test-network-ports-faster-with-powershell
https://learn-powershell.net/2016/04/22/speedy-ping-using-powershell/
https://stackoverflow.com/questions/53618904/how-to-use-multithreading-script
https://learn.microsoft.com/en-us/dotnet/api/system.net.networkinformation.ping?view=net-8.0
https://github.com/jaydo1/Scripts/blob/master/PowershellScripts/BgPing.ps1
https://www.reddit.com/r/PowerShell/comments/6eyhpv/whats_the_quickest_way_to_ping_a_computer/
https://randombrainworks.com/2018/01/28/powershell-background-jobs-runspace-jobs-thread-jobs/


#>

#List the modules we want to load per line
<#
$LoadModules = @'
PoshRSJob
SplitPipeline
'@.Split("`n").Trim()
#>

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

Clear-Variable A, B, C, D, E, F, Servers -Force -Confirm:$False -ErrorAction SilentlyContinue

if($NULL -eq $Servers)
{
    $Servers = ((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique
}

<#
Measure-Command {
$A = $Servers | Start-RSJob -Throttle 100 -Batch "Test" -ScriptBlock {
    
    Param($Server)

    
    Try
    {
        if(([System.Net.Sockets.TcpClient]::new().ConnectAsync($Server, 3389).Wait(1000)) -eq "True")
        {
            $Server
        }
    }
    Catch
    {}

} | Wait-RSJob -ShowProgress | Receive-RSJob
}


Measure-Command {
$B = $Servers | Start-RSJob -Throttle 100 -Batch "Test" -ScriptBlock {
    
    Param($Server)

    Try
    {
        if($NULL -ne (Get-CimInstance -ErrorAction SilentlyContinue -ClassName Win32_PingStatus -Filter "Address='$($Server)' AND Timeout=1000").ResponseTime)
        {
            $Server
        }
    }
    Catch
    {}

} | Wait-RSJob -ShowProgress | Receive-RSJob
}

Measure-Command {
$C = $Servers | Start-RSJob -Throttle 100 -Batch "Test" -ScriptBlock {
    
    Param($Server)

    Try
    {
        if(([System.Net.NetworkInformation.ping]::new().SendPingAsync($Server).Wait(1000)) -eq "True")
        {
            $Server
        }
    }
    Catch
    {}

} | Wait-RSJob -ShowProgress | Receive-RSJob
}
#>

Measure-Command {
$D = $Servers | split-pipeline -count 64 {
    Process
    {
        Try
        {
            if(([System.Net.NetworkInformation.ping]::new().SendPingAsync($_).Wait(1000)) -eq "True")
            {
                $_
            }
        }
        Catch
        {}
    }
}
}

Measure-Command {
$E = $Servers | split-pipeline -count 64 {
    Process{
        Try
        {
            if(([System.Net.Sockets.TcpClient]::new().ConnectAsync($_, 3389).Wait(1000)) -eq "True")
            {
                $_
            }
        }
        Catch
        {}
    }
}
}
