CLS

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

$Servers = Get-Content -Path "$Path\Servers.txt"

$password = ConvertTo-SecureString 'password123' -AsPlainText -Force

$results = @()
$credential = $NULL

foreach($Server in $Servers)
{
    $ServerName = $Server.split("\")[0]
    $User = $Server.split("\")[1]

    if($User -eq "Administrator")
    {
        if($NULL -eq $credential)
        {
            $credential = New-Object System.Management.Automation.PSCredential($User, $password)
        }

        Try 
        { 
            $session = New-RDPSession -ComputerName $ServerName -Credential $credential 
            
            # If the connection is successful, add a "Success" result to the results array 
            $results += New-Object PSObject -Property @{ Server = $ServerName; Result = "Success" }
            
            # Disconnect the RDP session 
            Remove-RDPSession -Session $session 
        } 
        Catch 
        { 
            # If the connection fails, add a "Failed" result to the results array 
            $results += New-Object PSObject -Property @{ Server = $ServerName; Result = "Failed" }
        }
    }
}

$results
