CLS

#===================  Load Module(s)

$LoadModules = @'
VMware.VimAutomation.Core
ImportExcel
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

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

$ErrorLog = "$path\Error.txt"
$ReplicationErrorLog = "$path\ReplicationErrorLog.txt"
$ReplicationLog = "$path\Replication-$((get-date).ToString("MM-dd-yyyy-HH-mm-ss")).txt"
$DC = "DC01"

#=================== vCenter Servers

$vCenters = @(
"vs01"
"vs02"
)

#===================  Connect to vCenter

if($NULL -eq ($global:DefaultVIServers.Name))
{
    $Connection = Connect-VIServer $vCenters -WarningAction SilentlyContinue -Credential $cred -ErrorAction Stop | Out-Null
}

#===================  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
Try
{
    $worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ErrorAction Stop -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name
}
Catch
{
    #Try loading our DLL and try again
    $DLL = "$((gmo -list importexcel).Path | split-path -Parent)\EPPlus.dll"
    Add-Type -Path $DLL
    $worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name
}

if($NULL -eq $ServerMigrations)
{
    $ServerMigrations = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1
}

#===================

if($DC -eq "DC01")
{
    ##DC01
    $vCenter = "vs02"
    $vmDNS = "1.2.3.4","2.3.4.5"
}
if($DC -eq "DC02")
{
    ##DC02
    $vCenter = "vs02"
    $vmDNS = "3.4.5.6","4.5.6.7"
}

#=================== Variables

if((Test-Path $ErrorLog) -eq $True)
{
    Remove-Item $ErrorLog
}

if((Test-Path $ReplicationErrorLog) -eq $True)
{
    Remove-Item $ReplicationErrorLog
}

$VMS = Get-VM -Server $vCenter | Get-SpbmEntityConfiguration -Server $vCenter | Where-Object StoragePolicy -like "VVol No Requirements Policy" #| Select-Object -First 100

foreach($VM in $VMS)
{
    foreach($ServerMigration in $ServerMigrations)
    {
        if($($VM.Entity).Name -eq $ServerMigration.VM)
        {
            $Server = $($VM.Entity).Name
            $vVolTarget = $($ServerMigration.'vVol Target')
            $Tier = (((Get-VM -Server $vCenter -Name $Server -ErrorAction Stop | Get-Tagassignment -Server $vCenter -ErrorAction Stop).Tag).Name | Where-Object {$_ -like "*Tier*"})

            #Remove spaces from the tier so we can use this to search as a filter
            if($NULL -ne $Tier)
            {
                $TierFilter = $Tier -replace " ",""
            }

            #Skip Tier 5+ because they don't have a replication Group
            if(($NULL -ne $TierFilter) -AND ($TierFilter -ne "Tier5") -OR ($TierFilter -ne "Tier6") -OR ($TierFilter -ne "Tier7"))
            {
                #Get vVol Policy
                if($NULL -ne $TierFilter)
                {
                    $vVolPolicy = (Get-SpbmStoragePolicy -Server $vCenter | Where-Object { ($_.Name -like "*$($TierFilter)*") -AND ($_.Name -like "*$($vVolTarget)*") }).Name
                    $vVolPolicyObj = (Get-SpbmStoragePolicy -Server $vCenter | Where-Object { ($_.Name -like "*$($TierFilter)*") -AND ($_.Name -like "*$($vVolTarget)*") })
                }

                #Get Datastore
                if($NULL -ne $vVolPolicy)
                {
                    $PUREDataStore = (Get-SpbmCompatibleStorage -Server $vCenter -StoragePolicy $vVolPolicy -ErrorAction Stop | Where-Object { $_.Name -like "*vVol*"}).Name
                }

                #Get the Current VM and Drive Info
                $VMConfig = ((Get-vm $Server -Server $vCenter), ((Get-vm $Server -Server $vCenter) | get-harddisk -Server $vCenter)) | Get-SpbmEntityConfiguration -Server $vCenter

                #Get the Replication Group
                if($NULL -ne $PUREDataStore)
                {
                    $ReplicationGroup = Get-SpbmReplicationGroup -Server $vCenter -Datastore $PUREDataStore -StoragePolicy $vVolPolicy
                }

                #DEBUG INFO
                
                $Server
                $vVolTarget
                $Tier
                $vVolPolicy
                $PUREDataStore
                $VMConfig
                $ReplicationGroup
                Write-Output ""
                Write-Output "--------------------"
                Write-Output ""

                if($NULL -ne $ReplicationGroup)
                {
                    Try
                    {
                        $VMConfig | Set-SpbmEntityConfiguration -StoragePolicy $vVolPolicyObj -ReplicationGroup $ReplicationGroup | Out-File $ReplicationLog -Append -force
                    }
                    Catch
                    {
                        #Log the info
                        "$Server Unable to update Replication Group" | Out-File $ReplicationErrorLog -Append -force
                    }
                }
                else
                {
                    #Log the info
                    "$Server doesn't have a tier listed or doesn't exist" | Out-File $ErrorLog -Append -force
                }

            }
            #Write-Output "$Server has Target: $vVolTarget and Tier: $Tier"
        }
    }
}