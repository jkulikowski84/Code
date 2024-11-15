CLS

#SRC: https://github.com/proxb/AsyncFunctions/blob/master/Test-ConnectionAsync.ps1

Clear-Variable Result, Task, Object -Force -Confirm:$False -ErrorAction SilentlyContinue

#===================

#Set Variables
#$IP = "123.123.123.123"

#Ask user for an IP
Do
{
    CLS
    $IP = Read-Host "Type in an IP"

}until($IP -match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$" -and [bool]($IP -as [ipaddress]))

$IPRange = 1..254 | % {"$($IP.Substring(0, $IP.lastIndexOf('.'))).$_"}

#===================

$Task = ForEach ($IP in $IPRange) 
{
    [pscustomobject] @{
        IP = $IP
        Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($IP,100)
    }
}        

Try 
{
    [void][Threading.Tasks.Task]::WaitAll($Task.Task)
} 
Catch {}

$Task | ForEach {
    If ($_.Task.IsFaulted) 
    {
        $Result = $_.Task.Exception.InnerException.InnerException.Message
        $IPAddress = $Null
    } 
    Else 
    {
        $Result = $_.Task.Result.Status
        $IPAddress = $_.task.Result.Address.ToString()
    }

    $Object = [pscustomobject]@{
        IP = $_.IP
        IPAddress = $IPAddress
        Result = $Result
    }

    $Object.pstypenames.insert(0,'Net.AsyncPingResult')
    $Object | Where-Object { ($_.Result -ne "Success") } | Select-Object -Property IP -ExpandProperty IP
}
