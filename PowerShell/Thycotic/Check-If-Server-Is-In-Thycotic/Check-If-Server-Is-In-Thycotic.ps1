CLS

$LoadModules = @'
SplitPipeline
Thycotic.SecretServer
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

#===================  Connect to Thycotic

if($NULL -eq $Thysession)
{
    $token = "tokencode"
    
    Try
    {
        $Thysession = New-TssSession https://thycotic.domain.com/SecretServer -AccessToken $token -WarningAction Stop
    }
    Catch
    {
        #Write-Output "Token expired. Replace your token."
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        $MessageBody = "Token expired. Replace your token."
        $MessageTitle = "Token Expired"
        $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

        if($NULL -ne $Result)
        {
            Exit
        }
    }
}

#===================

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

$Servers = Get-Content -Path "$Path\Servers.txt"

$results = @()

#User for progress
$data = @{
    Count = $(($Servers | measure).count)
	Done = 0
}

foreach($Server in $Servers)
{
    #Clear variable every iteration    
    $Secret = $NULL

    $ServerName = $Server.split("\")[0]
    $User = $Server.split("\")[1]

    #$Secret = (Search-TssSecret -TssSession $Thysession -SearchText $ServerName)
    $Secret = (Search-TssSecret -TssSession $Thysession -Field "Machine" -SearchText $ServerName -ExactMatch)

    if($NULL -ne $Secret)
    {
        $results += New-Object PSObject -Property @{ Server = $ServerName; Result = "Success" }
    }

    $done = ++$data.Done

    # show progress
    Write-Progress -Activity "Done $done" -Status Processing -PercentComplete (100*$done/$data.Count)
}


<#
$Servers | split-pipeline -count 64 -Variable results, Thysession {
    Process
    {
        $Secret = $NULL

        $ServerName = $_.split("\")[0]

        $Secret = (Search-TssSecret -TssSession $Thysession -Field "Machine" -SearchText $ServerName -ExactMatch)

        if($NULL -ne $Secret)
        {
            $results += New-Object PSObject -Property @{ Server = $ServerName; Result = "Success" }
        }
    }
}
#>
$results
