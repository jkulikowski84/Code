CLS

#List the modules we want to load per line
$LoadModules = @'
VMware.PowerCLI
'@.Split("`n").Trim()

ForEach($Module in $LoadModules) 
{
    if($NULL -eq (Get-Module -ListAvailable $Module))
    {
        Import-Module $Module
    }
}

#================= Configure PowerCLI

if(-NOT (Get-PowerCLIConfiguration).Scope -eq "User" -AND ((Get-PowerCLIConfiguration).InvalidCertificateAction -eq "Ignore") -AND ((Get-PowerCLIConfiguration).DefaultVIServerMode -eq "Multiple"))
{
    Set-PowerCLIConfiguration -Scope User -DefaultVIServerMode Multiple -ParticipateInCEIP $false -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $False -Confirm:$False | Out-NULL
}

#===================  Connect to vCenter

if($NULL -eq ($global:DefaultVIServers.Name))
{
    $cred = (Get-Credential (whoami))
    
    Connect-VIServer "pvsa01vcsa0001" -Protocol https -Credential $cred -AllLinked -WarningAction 0 | Out-NULL
    Connect-VIServer "pvsa02vcsa0001" -Protocol https -Credential $cred -AllLinked -WarningAction 0 | Out-NULL
}

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

$PoweredOffVMs = ((get-vm *) | Where-Object { $_.PowerState -eq "PoweredOff" }).Name
$AllVMs = (get-vm *).Name

$ServersList = [System.IO.File]::ReadAllLines("$Path\PoweredOff.txt")

ForEach ($Item in $ServersList)
{
    If (($item -in $PoweredOffVMs) -OR ($Item -notin $AllVMs))
    {
        $item
    }
}