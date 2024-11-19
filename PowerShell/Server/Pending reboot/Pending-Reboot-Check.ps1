CLS

#------------- PoshRSJob (multitasking)

if((Get-Module -ListAvailable -Name "PoshRSJob") -or (Get-Module -Name "PoshRSJob"))
{
        Import-Module PoshRSJob
}
else
{	
    Install-Module -Name PoshRSJob -Scope CurrentUser -Force -Confirm:$False
	Import-Module PoshRSJob
}


#------------- ImportExcel Module 

if((Get-Module -ListAvailable -Name "ImportExcel") -or (Get-Module -Name "ImportExcel"))
{
        Import-Module ImportExcel
}
else
{
    #Install NuGet (Prerequisite) first
	Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$False
	
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$False
	Import-Module ImportExcel
}

#Clear screen again
CLS

#----------------------------------------------------------------------------------------------------------------

#Start Timestamp
$Start = Get-Date

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.csv"

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

#------------------------------------ Populate our variable with data from spreadsheet

$ExcelServersList = foreach($ExcelServer in $ExcelServers) {
    $ExcelServer | Select-Object @{Name="ServerName";Expression={$_.Child}}, "Primary", @{Name="PatchWindow";Expression={$_."Patch Window"}}, @{Name="TestServer";Expression={$_."Test Server"}}, "DMZ" 
}

#------------------------------------ Remove Duplicate entries

$SortedExcelServersList = ($ExcelServersList | Sort-Object -Property ServerName -Unique)

#------------------------------------ Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if($($SortedExcelServerList.DMZ) -eq $true)
    {
        $SortedExcelServerList.ServerName = [System.String]::Concat("$($SortedExcelServerList.ServerName)",".dmz.com")
    }

    $SortedExcelServerList
}

#------------------------------------ Grab all servers from AD so we can use to compare against our list - also trimany whitespaces from output

$Servers = (dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0 -attr Name | sort).trim()

#------------------------------------ Compare our list to servers in AD and filter out appliances


$FilteredServersResult = $Null

$FilteredServersResult = ForEach ($Item in $FilteredServers) 
{
    If (($item.servername -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

#------------------------------------ Multithreading Magic

$FilteredServersResult | Start-RSJob -Throttle 25 -ScriptBlock {

    Param($Server)

    $PendingReboot = $False
    $operationSource = $Operations = $trueOperationsCount = $operationDestination = $trueRenames = $NULL

    #Ping servers to make sure they're responsive
    if($NULL -ne (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($Server.servername)' AND Timeout=100").ResponseTime)
    { 
        Try
        {
            if ([bool]([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$($Server.servername)).OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing').GetValue('RebootPending')) -eq $True) { $PendingReboot = $true }
            if ([bool]([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$($Server.servername)).OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update').GetValue('RebootRequired')) -eq $True) { $PendingReboot =  $true }

            $Operations = ([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$($Server.servername)).OpenSubKey('SYSTEM\CurrentControlSet\Control\Session Manager').GetValue('PendingFileRenameOperations'))

            if ($null -eq $Operations) 
            {
                $PendingReboot = $false
            } 
            else 
            {
                $trueOperationsCount = $operations.Length / 2
                $trueRenames = [System.Collections.Generic.Dictionary[string, string]]::new($trueOperationsCount)

                for ($i = 0; $i -ne $trueOperationsCount; $i++) 
                {
                    $operationSource = $operations[$i * 2]
                    $operationDestination = $operations[$i * 2 + 1]

                    if ($operationDestination.Length -eq 0) 
                    {
                        Write-Verbose "Ignoring pending file delete '$operationSource'"
                    } 
                    else 
                    {
                        Write-Host "Found a true pending file rename (as opposed to delete). Source '$operationSource'; Dest '$operationDestination'"
                        $trueRenames[$operationSource] = $operationDestination
                    }
                }
                $PendingReboot = ($trueRenames.Count -gt 0)
            }

            Try 
            { 
                $util = ([wmiclass]"\\$($Server.servername)\root\ccm\clientsdk:CCM_ClientUtilities")
                $status = $util.DetermineIfRebootPending()
                if (($status -ne $null) -and $status.RebootPending) 
                {
                    $PendingReboot = $true
                }
            }
            Catch { }
        }
        Catch
        {           
            if($Error.exception -like "*")
            {
                ($Server | Add-Member -NotePropertyMembers @{"Error" = [string]$Error} -PassThru) | Export-Csv -Path $using:ErrorFile -NoTypeInformation -Force -Append
            }
        }
    }

    if($PendingReboot -eq $True)
    {
        ($Server | Add-Member -NotePropertyMembers @{"Pending Reboot Reason" = [string]"Found a true pending file rename (as opposed to delete). Source '$operationSource'; Dest '$operationDestination'"} -PassThru)
    }

} | Wait-RSJob -ShowProgress | Receive-RSJob | Export-Csv -Path "$Path\Results.csv" -NoTypeInformation -Force

$End =  (Get-Date)

$End - $Start
