CLS

$PrintServers = (dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(name=*EPMS*))" -limit 0 -attr Name | sort).trim()

foreach($PrintServer in $PrintServers)
{
    $ServiceFailure = sc.exe "\\$PrintServer" qfailure "Spooler"

    $RecoveryParams = [hashtable]@{
        RecoveryAction = [PSCustomObject[]]@()
    }

    $ServiceFailure | ForEach-Object {
        if ($_ -imatch '^\s*(?:FAILURE_ACTIONS\s+:\s+)?(?<RecoveryAction>RESTART|RUN PROCESS|REBOOT) -- Delay = (?<RecoveryDelay>\d+) milliseconds.$')
        {
            $RecoveryParams.RecoveryAction += "Server=$PrintServer Action=$($matches.RecoveryAction) Delay=$([TimeSpan]::FromMilliseconds($matches.RecoveryDelay));"
            <#
             switch ($matches.RecoveryAction) 
             {
                default { $RecoveryParams.RecoveryAction += "Action=$($matches.RecoveryAction) Delay=$([TimeSpan]::FromMilliseconds($matches.RecoveryDelay));"; break }
            }
            #>
        }
    }

    <#
    if(($RecoveryParams.Values).length -le 2)
    {
        Try
        {
            sc.exe "\\$PrintServer" failure "Spooler" reset= 30 actions= restart/5000
        }
        Catch
        {
            $PrintServer
        }
    }
    #>
    $RecoveryParams
    Clear-Variable RecoveryParams
}