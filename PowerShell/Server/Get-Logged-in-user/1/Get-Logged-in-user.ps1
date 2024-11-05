CLS

Function Get-LoggedInUser 
{ 
    [CmdletBinding()]
        param(
            [Parameter(
                Mandatory = $false,
                ValueFromPipeline = $true,
                ValueFromPipelineByPropertyName = $true,
                Position=0
            )]
            [string[]] $ComputerName = $env:COMPUTERNAME,
 
 
            [Parameter(
                Mandatory = $false
            )]
            [Alias("SamAccountName")]
            [string]   $UserName
        )
 
    BEGIN {}
 
    PROCESS 
    {
        foreach ($Computer in $ComputerName) 
        {
            try 
            {
                $Computer = $Computer.ToUpper()
                $SessionList = quser /Server:$Computer 2>$null

                if ($SessionList) 
                {
                    $UserInfo = foreach ($Session in ($SessionList | select -Skip 1)) {
                        
                        $Session = $Session.ToString().trim() -replace '\s+', ' ' -replace '>', ''
                        
                        if ($Session.Split(' ')[3] -eq 'Active') 
                        {
                            [PSCustomObject]@{
                                ComputerName = $Computer
                                UserName     = $session.Split(' ')[0]
                                SessionName  = $session.Split(' ')[1]
                                SessionID    = $Session.Split(' ')[2]
                                SessionState = $Session.Split(' ')[3]
                                IdleTime     = $Session.Split(' ')[4]
                                LogonTime    = $session.Split(' ')[5, 6, 7] -as [string] -as [datetime]
                            }
                        } 
                        else 
                        {
                            [PSCustomObject]@{
                                ComputerName = $Computer
                                UserName     = $session.Split(' ')[0]
                                SessionName  = $null
                                SessionID    = $Session.Split(' ')[1]
                                SessionState = 'Disconnected'
                                IdleTime     = $Session.Split(' ')[3]
                                LogonTime    = $session.Split(' ')[4, 5, 6] -as [string] -as [datetime]
                            }
                        }
                    }
 
                    if ($PSBoundParameters.ContainsKey('Username')) 
                    {
                        $UserInfo | Where-Object {$_.UserName -eq $UserName}
                    } 
                    else 
                    {
                        $UserInfo | Sort-Object LogonTime
                    }
                }
            } 
            catch 
            {
                Write-Error $_.Exception.Message
            }
        }
    }
 
    END {}
}

$PC = "n070982"

Get-LoggedInUser -ComputerName $PC