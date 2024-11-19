CLS

#=================== Load Remote Modules

#List the remote module you want to load and the server you want to load it from
$LoadRemoteModules = @'
STProtect, IvantiServer
'@.Split("`n").Trim()

ForEach($RemoteModuleIfo in $LoadRemoteModules)
{
    #Seperate the Module from the Server
    $RemoteModule = ($RemoteModuleIfo -split ",").Trim()[0]
    $RemoteModuleServer = ($RemoteModuleIfo -split ",").Trim()[1]

    if(($NULL -eq ((Get-PSSession).ComputerName)) -OR (((Get-PSSession).ComputerName) -ne $RemoteModuleServer))
    {
        $Session = New-PSSession -ComputerName $RemoteModuleServer -Authentication Kerberos
        Invoke-Command -Session $Session -ScriptBlock { Import-Module STProtect }
    }

    if(!((((Get-Module).ExportedCommands).Values).Name -like "*Add-MachineGroup*"))
    {
        Import-PSSession -Session $Session -Module STProtect -AllowClobber | Out-Null
    }
}

#=================== Import Modules

#List the modules we want to load per line
$LoadModules = @'
PoshRSJob
ImportExcel
'@.Split("`n").Trim()

ForEach($Module in $LoadModules) 
{
    if($NULL -eq (Get-Module -ListAvailable $Module))
    {
        Import-Module $Module
    }
}

#=================== Global Variables

#Start Timestamp
$Start = (Get-Date)

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

#Set our script location to the path:
if((Get-Location).Path -ne $Path)
{
    Set-Location $Path
}
#=================== Remove Old Spreadsheets

$AllSpreadsheets = [System.IO.Directory]::EnumerateFiles("$Path","*.Xlsx*","AllDirectories") -notlike "*Filtered-Spreadsheet-MASTER*"
$AllSpreadsheets | Remove-Item -Force -Confirm:$False

#===================  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Load our Worksheet
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

Try
{
    $ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1 -AsDate "Most recent discovery","Last Reboot Date" -ErrorAction Stop
}
Catch
{
    $ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1
}

#=================== Remove Duplicate entries and sort by Name

$SortedExcelServersList = ($ExcelServers | Sort-Object -Property Child -Unique)

#=================== Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if(($($SortedExcelServerList.DMZ) -eq $true) -AND ($($SortedExcelServerList.child) -notlike "*.dmz.com"))
    {
        $SortedExcelServerList.child = [System.String]::Concat("$($SortedExcelServerList.child)",".dmz.com")
    }

    $SortedExcelServerList
}

#=================== Grab all servers from AD so we can use to compare against our list - also trim any whitespaces from output

$Servers = ((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique

#=================== Compare our list to servers in AD and filter out appliances

$FilteredServersResult = ForEach ($Item in $FilteredServers) 
{
    If (($item.child -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

#=================== Create our spreadsheet

## All Servers
$FilteredServersResult | Export-Excel -Path "FilteredSpreadsheet.xlsx" -WorksheetName "All Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== Ivanti Stuff Starts here

#Main Group
$MainGroup = "zzPatching"

#=================== REMOVE ALL MACHINE GROUPS UNDER THE MAIN GROUP TO MAKE SURE WE DON'T HAVE STALE DATA

if((Get-MachineGroup).Path -like "$MainGroup*")
{
    ((Get-MachineGroup | Select-Object Name,Path) | Where-Object { $_.path -like "$MainGroup*" }).Name | Remove-MachineGroup -Confirm:$False -ErrorAction SilentlyContinue | Out-NULL
}

#=================== Recreate all the Ivanti Patch Groups

#=================== Static Groups

##Autos
Add-MachineGroup -Name "Auto Test Servers" -Path "$MainGroup\1. TEST\1. 10AM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Test EPIC Servers" -Path "$MainGroup\1. TEST\2. 6PM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Prod TUE 10AM Servers" -Path "$MainGroup\2. PROD\1. TUE\1. 10AM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Prod TUE 6PM Servers" -Path "$MainGroup\2. PROD\1. TUE\2. 6PM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Prod TUE 9PM Servers" -Path "$MainGroup\2. PROD\1. TUE\3. 9PM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Prod THR 10AM Servers" -Path "$MainGroup\2. PROD\2. THR\1. 10AM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Prod THR 6PM Servers" -Path "$MainGroup\2. PROD\2. THR\2. 6PM\AUTO" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Auto Prod THR 9PM Servers" -Path "$MainGroup\2. PROD\2. THR\3. 9PM\AUTO" -ErrorAction SilentlyContinue | Out-Null

#Manuals
Add-MachineGroup -Name "Manual Test Lab" -Path "$MainGroup\1. TEST\1. 10AM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Manual Test EPIC" -Path "$MainGroup\1. TEST\1. 10AM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Manual Test Provisioning Workstations" -Path "$MainGroup\1. TEST\1. 10AM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Manual Prod Tue 6PM Stryker" -Path "$MainGroup\2. PROD\1. TUE\2. 6PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Manual Prod THR 6PM EPIC" -Path "$MainGroup\2. PROD\2. THR\2. 6PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
Add-MachineGroup -Name "Manual Prod THR 9PM OneContent" -Path "$MainGroup\2. PROD\2. THR\3. 9PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null

#=================== Dynamic Groups

$TestLabs = @("eClinicalWorks","QML","TEG","Novanet","Data Innovations")
$ProdLabs = @("Biomerieux","Novanet","BioRad","TEG","Voicebrook","Data Innovations","QML","Softlab","WAM","Synapsys")

$TESTServerListManualReboot10AM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "TESTServerListManualReboot") 
    { 
        if(($FilteredServer.Parent -notlike "*EPIC*") -AND ($FilteredServer.Parent -notlike "*AD Manager*") -AND (($FilteredServer.Parent -notmatch "($(($TestLabs|ForEach{[RegEx]::Escape($_)}) -join '|'))")))
        {
            $(($FilteredServer.'Parent').Split(" ")[0])
        }
    }
}) | Sort -Unique

foreach($Path in $TESTServerListManualReboot10AM)
{
    Add-MachineGroup -Name "Manual Test $($Path)" -Path "$MainGroup\1. TEST\1. 10AM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

$TuesdayManualRebootLAB10AM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 10 am") 
    { 
        if(($FilteredServer.Parent -match "($(($ProdLabs|ForEach{[RegEx]::Escape($_)}) -join '|'))"))
        {
            $($FilteredServer.'Parent')
        }
    }
}) | Sort -Unique

foreach($Path in $TuesdayManualRebootLAB10AM)
{
    Add-MachineGroup -Name "Manual Prod Tue 10AM $($Path)" -Path "$MainGroup\2. PROD\1. TUE\1. 10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-Null
}

$TuesdayManualReboot10AM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if(($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 10 am") -AND (($FilteredServer.Parent -notmatch "($(($TuesdayManualRebootLAB10AM|ForEach{[RegEx]::Escape($_)}) -join '|'))")))
    { 
        $($FilteredServer.'Parent')
    }
}) | Sort -Unique

foreach($Path in $TuesdayManualReboot10AM)
{
    Add-MachineGroup -Name "Manual Prod Tue 10AM $($Path)" -Path "$MainGroup\2. PROD\1. TUE\1. 10AM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

$TuesdayManualReboot6PM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 6 pm") 
    { 
        if($FilteredServer.Parent -notlike "*Stryker*")
        { 
            $(($FilteredServer.'Parent').Split(" ")[0])
        } 
    }
}) | Sort -Unique

foreach($Path in $TuesdayManualReboot6PM)
{
    Add-MachineGroup -Name "Manual Prod Tue 6PM $($Path)" -Path "$MainGroup\2. PROD\1. TUE\2. 6PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

$TuesdayManualReboot9PM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 9 pm") 
    { 
        $($FilteredServer.'Parent')
    }
}) | Sort -Unique

foreach($Path in $TuesdayManualReboot9PM)
{
    Add-MachineGroup -Name "Manual Prod Tue 9PM $($Path)" -Path "$MainGroup\2. PROD\1. TUE\3. 9PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

$ThursdayManualReboot10AM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "ThursdayManualReboot - 10 am") 
    { 
        $(($FilteredServer.'Parent').Split(" ")[0])
    }
}) | Sort -Unique

foreach($Path in $ThursdayManualReboot10AM)
{
    Add-MachineGroup -Name "Manual Prod THR 10AM $($Path)" -Path "$MainGroup\2. PROD\2. THR\1. 10AM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

$ThursdayManualReboot6PM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "ThursdayManualReboot - 6 pm") 
    { 
        if($FilteredServer.Parent -notlike "*EPIC*")
        { 
            $FilteredServer.Parent
        } 
    }
}) | Sort -Unique

foreach($Path in $ThursdayManualReboot6PM)
{
    Add-MachineGroup -Name "Manual Prod THR 6PM $($Path)" -Path "$MainGroup\2. PROD\2. THR\2. 6PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

$ThursdayManualReboot9PM = @(foreach($FilteredServer in $FilteredServersResult)
{
    if($FilteredServer.'Patch Window' -eq "ThursdayManualReboot - 9 pm") 
    { 
        if($FilteredServer.Parent -notlike "*ROI*")
        { 
            $FilteredServer.Parent
        } 
    }
}) | Sort -Unique

foreach($Path in $ThursdayManualReboot9PM)
{
    Add-MachineGroup -Name "Manual Prod THR 9PM $($Path)" -Path "$MainGroup\2. PROD\2. THR\3. 9PM\MANUAL" -ErrorAction SilentlyContinue | Out-Null
}

#=================== Import Computers into our Groups

#======================================= TEST =======================================

$TestAuto = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TestServerList") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Test Servers" -EndpointNames $TestAuto -ErrorAction Stop } Catch { }

#Populate 10AM Test Manual LAB Servers
$TestLabServers = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TESTServerListManualReboot") { if($FilteredServer.Parent -match "($(($TestLabs|ForEach{[RegEx]::Escape($_)}) -join '|'))") { $FilteredServer.Child } } }
Try { Add-MachineGroupItem -Name "Manual Test Lab" -EndpointNames $TestLabServers -ErrorAction Stop } Catch { }

#Populate 10AM Test Manual EPIC Servers
$TESTManualEPIC = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TESTServerListManualReboot") { If($FilteredServer.'Parent' -like "*EPIC*") { $FilteredServer.Child } } }
Try { Add-MachineGroupItem -Name "Manual Test EPIC" -EndpointNames $TESTManualEPIC -ErrorAction Stop } Catch { }

#Populate 10AM Test Manual AD Manager/Provisioning Servers/Workstations
$TESTManualADM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TESTServerListManualReboot") { If($FilteredServer.'Parent' -like "*AD Manager*") { $FilteredServer.Child } } }
Try { Add-MachineGroupItem -Name "Manual Test Provisioning Workstations" -EndpointNames $TESTManualADM -ErrorAction Stop } Catch { }

#Populate 10AM All other Manual Servers that aren't in any of the above categories
foreach($FilteredServer in $FilteredServersResult) 
{
    if(($FilteredServer.'Patch Window' -eq "TESTServerListManualReboot") -AND ($FilteredServer.'Parent' -notlike "*AD Manager*") -AND ($FilteredServer.'Parent' -notlike "*EPIC*") -AND ($FilteredServer.Parent -notmatch "($(($TestLabs|ForEach{[RegEx]::Escape($_)}) -join '|'))"))
    {
        #$FilteredServer.Child
        Try { Add-MachineGroupItem -Name "Manual Test $(($FilteredServer.'Parent').Split(" ")[0])" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
    }
}
#Populate 6PM Test Auto EPIC Servers
$TestAutoEPIC = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "Test Servers Epic") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Test EPIC Servers" -EndpointNames $TestAutoEPIC -ErrorAction Stop } Catch { }

#======================================= PROD =======================================

        #==================== Tuesday

$TuesdayAutoReboot10AM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TuesdayAutoReboot - 10 am") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Prod TUE 10AM Servers" -EndpointNames $TuesdayAutoReboot10AM -ErrorAction Stop } Catch { }

foreach($FilteredServer in $FilteredServersResult) 
{ 
    if($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 10 am") 
    { 
        if($FilteredServer.Parent -match "($(($ProdLabs|ForEach{[RegEx]::Escape($_)}) -join '|'))") 
        { 
            #$FilteredServer.Child
            Try { Add-MachineGroupItem -Name "Manual Prod Tue 10AM $($FilteredServer.'Parent')" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
        else
        {
            Try { Add-MachineGroupItem -Name "Manual Prod Tue 10AM $($FilteredServer.'Parent')" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
    } 
}

$TuesdayAutoReboot6PM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TuesdayAutoReboot - 6 pm") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Prod TUE 6PM Servers" -EndpointNames $TuesdayAutoReboot6PM -ErrorAction Stop } Catch { }

foreach($FilteredServer in $FilteredServersResult) 
{ 
    if($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 6 pm") 
    { 
        if($FilteredServer.Parent -like "*Stryker*") 
        { 
            #$FilteredServer.Child
            Try { Add-MachineGroupItem -Name "Manual Prod Tue 6PM Stryker" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
        else
        {
            Try { Add-MachineGroupItem -Name "Manual Prod Tue 6PM $(($FilteredServer.'Parent').Split(" ")[0])" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
    } 
}

$TuesdayAutoReboot9PM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "TuesdayAutoReboot - 9 pm") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Prod TUE 9PM Servers" -EndpointNames $TuesdayAutoReboot9PM -ErrorAction Stop } Catch { }

foreach($FilteredServer in $FilteredServersResult) 
{ 
    if($FilteredServer.'Patch Window' -eq "TuesdayManualReboot - 9 pm") 
    { 
        Try { Add-MachineGroupItem -Name "Manual Prod Tue 9PM $($FilteredServer.'Parent')" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
    } 
}

        #==================== Thursday

$ThursdayAutoReboot10AM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "ThursdayAutoReboot - 10 am") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Prod THR 10AM Servers" -EndpointNames $ThursdayAutoReboot10AM -ErrorAction Stop } Catch { }

foreach($FilteredServer in $FilteredServersResult) 
{ 
    if($FilteredServer.'Patch Window' -eq "ThursdayManualReboot - 10 am") 
    { 
        Try { Add-MachineGroupItem -Name "Manual Prod THR 10AM $(($FilteredServer.'Parent').Split(" ")[0])" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
    } 
}

$ThursdayAutoReboot6PM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "ThursdayAutoReboot - 6 pm") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Prod THR 6PM Servers" -EndpointNames $ThursdayAutoReboot6PM -ErrorAction Stop } Catch { }

foreach($FilteredServer in $FilteredServersResult) 
{ 
    if($FilteredServer.'Patch Window' -eq "ThursdayManualReboot - 6 pm") 
    {
        if($FilteredServer.'Parent' -like "*EPIC*" )
        {
            Try { Add-MachineGroupItem -Name "Manual Prod THR 6PM EPIC" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
        else
        {
            Try { Add-MachineGroupItem -Name "Manual Prod THR 6PM $($FilteredServer.'Parent')" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
    } 
}

$ThursdayAutoReboot9PM = foreach($FilteredServer in $FilteredServersResult) { if($FilteredServer.'Patch Window' -eq "ThursdayAutoReboot - 9 pm") { $FilteredServer.Child } }
Try { Add-MachineGroupItem -Name "Auto Prod THR 9PM Servers" -EndpointNames $ThursdayAutoReboot9PM -ErrorAction Stop } Catch { }

foreach($FilteredServer in $FilteredServersResult) 
{ 
    if($FilteredServer.'Patch Window' -eq "ThursdayManualReboot - 9 pm") 
    {
        if($FilteredServer.'Parent' -like "*ROI*" )
        {
            Try { Add-MachineGroupItem -Name "Manual Prod THR 9PM OneContent" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
        else
        {
            Try { Add-MachineGroupItem -Name "Manual Prod THR 9PM $($FilteredServer.'Parent')" -EndpointNames $FilteredServer.Child -ErrorAction Stop } Catch { }
        }
    } 
}

#####################=======================================================================================================================================##### 

##============ Updating the Main Business Services

$Shavlik_Business_Service_Group = "Business Services"

#=================== REMOVE ALL MACHINE GROUPS UNDER THE MAIN GROUP TO MAKE SURE WE DON'T HAVE STALE DATA

if((Get-MachineGroup).Path -like "$Shavlik_Business_Service_Group*")
{
    ((Get-MachineGroup | Select-Object Name,Path) | Where-Object { ($_.path -like "$Shavlik_Business_Service_Group*") -AND  ($_.Name -notlike "*Workstations*") }).Name | Remove-MachineGroup -Confirm:$False -ErrorAction SilentlyContinue | Out-NULL
}

$End = (Get-Date)
$End - $Start
