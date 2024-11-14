CLS

Try
{
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
}
Catch
{
    Throw "Could not load assembly: $_"
}

#===================

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

#===================

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

$Servers = Get-Content -Path "$Path\Servers.txt"

$password = ConvertTo-SecureString 'Password123' -AsPlainText -Force

#$results = @()
$credential = $NULL

#User for progress
$data = @{
    Count = $(($Servers | measure).count)
	Done = 0
}

$results = $Servers | split-pipeline -count 64 -Variable results, password, data {
    Process
    {
        $results = $DS = $Test = $NULL
        $ServerName = $_.split("\")[0]
        $User = $_.split("\")[1]

        if($User -eq "Administrator")
        {
            if($NULL -eq $credential)
            {
                $credential = New-Object System.Management.Automation.PSCredential($User, $password)
            }

            Try 
            { 
                $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine, $ServerName)
            }
            Catch{}

            Try
            {
                $Test = $DS.ValidateCredentials($credential.UserName, $credential.GetNetworkCredential().password)
            }
            Catch
            {}

            if($Test -eq "True")
            {
                #$results += New-Object PSObject -Property @{ Server = $ServerName; Result = "Success" }
                New-Object PSObject -Property @{ Server = $ServerName; Result = "Success" }
            }
            else
            { 
                # If the connection fails, add a "Failed" result to the results array 
                #$results += New-Object PSObject -Property @{ Server = $ServerName; Result = "Failed" }
                New-Object PSObject -Property @{ Server = $ServerName; Result = "Failed" }
            }

            $DS.Dispose()
        }

        $done = ++$data.Done

        # show progress
        Write-Progress -Activity "Done $done" -Status Processing -PercentComplete (100*$done/$data.Count)
    }
}

$results | Export-Csv "Servers.csv" -NoTypeInformation
