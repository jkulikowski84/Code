CLS

#=================== Load Ivanti Shavlik Module Remotely into our current session

$IvantiServer = "IvantiServer"
$SessionsRunning = get-pssession

if(!($SessionsRunning.ComputerName -like $IvantiServer))
{
    $Session = New-PSSession -ComputerName $IvantiServer -Authentication Kerberos
    Invoke-Command -Session $Session -ScriptBlock { Import-Module STProtect }
}

$CheckModules = get-module

#Import the module from our Server session if it's not already imported
if(!(((($CheckModules).ExportedCommands).Values).Name -like "*Add-MachineGroup*"))
{
    Import-PSSession -Session $Session -Module STProtect -AllowClobber | Out-Null
}

#=================== PoshRSJob (multitasking)

if((Get-Module -ListAvailable -Name "PoshRSJob") -or (Get-Module -Name "PoshRSJob"))
{
        Import-Module PoshRSJob
}
else
{	
    Install-Module -Name PoshRSJob -Scope CurrentUser -Force -Confirm:$False
	Import-Module PoshRSJob
}

#=================== ImportExcel Module 

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

#===================

#Start Timestamp
$Start = Get-Date

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.csv"

#=================== Remove Old Spreadsheets

$FailedPingTestCSV = (Split-Path $script:MyInvocation.MyCommand.Path) + "\FailedPingTestCSV.csv"

if((Test-Path $FailedPingTestCSV) -eq $True)
{
    Remove-Item $FailedPingTestCSV
}

$FailedPingTest = (Split-Path $script:MyInvocation.MyCommand.Path) + "\FailedPingTest.xlsx"

if((Test-Path $FailedPingTest) -eq $True)
{
    Remove-Item $FailedPingTest
}

$TEST_MAN_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\TEST\TEST_MAN_SERVER_GROUPS.xlsx"

if((Test-Path $TEST_MAN_SERVER_GROUPS) -eq $True)
{
    Remove-Item $TEST_MAN_SERVER_GROUPS
}

$TEST_AUTO_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\TEST\TEST_AUTO_SERVER_GROUPS.xlsx"

if((Test-Path $TEST_AUTO_SERVER_GROUPS) -eq $True)
{
    Remove-Item $TEST_AUTO_SERVER_GROUPS
}

$TEST_AUTO_EPIC_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\TEST\TEST_AUTO_EPIC_SERVER_GROUPS.xlsx"

if((Test-Path $TEST_AUTO_EPIC_SERVER_GROUPS) -eq $True)
{
    Remove-Item $TEST_AUTO_EPIC_SERVER_GROUPS
}

$PROD_MAN_TUE_10AM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\PROD\TUE\PROD_MAN_TUE_10AM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_TUE_10AM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_TUE_10AM_SERVER_GROUPS
}

$PROD_MAN_TUE_6PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\PROD\TUE\PROD_MAN_TUE_6PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_TUE_6PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_TUE_6PM_SERVER_GROUPS
}

$PROD_MAN_TUE_9PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\PROD\TUE\PROD_MAN_TUE_9PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_TUE_9PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_TUE_9PM_SERVER_GROUPS
}

$PROD_MAN_THR_10AM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\PROD\THR\PROD_MAN_THR_10AM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_THR_10AM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_THR_10AM_SERVER_GROUPS
}

$PROD_MAN_THR_6PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\PROD\THR\PROD_MAN_THR_6PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_THR_6PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_THR_6PM_SERVER_GROUPS
}

$PROD_MAN_THR_9PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\PROD\THR\PROD_MAN_THR_9PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_THR_9PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_THR_9PM_SERVER_GROUPS
}

$FilteredSpreadsheet = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Filtered-Spreadsheet.xlsx"

if((Test-Path $FilteredSpreadsheet) -eq $True)
{
    Remove-Item $FilteredSpreadsheet
}

#===================  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

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

#=================== Grab all servers from AD so we can use to compare against our list - also trimany whitespaces from output

$Servers = ((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique

#=================== Compare our list to servers in AD and filter out appliances

$FilteredServersResult = $FilteredServersResultA = $Null

$FilteredServersResultA = ForEach ($Item in $FilteredServers) 
{
    If (($item.child -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

$FilteredServersResult = $FilteredServersResultA

#=================== Ping our Servers to make sure they're online, If they're not then filter them out and output to a file *** Added as an extra measure ***

<#
$ping = New-Object System.Net.NetworkInformation.Ping

$FilteredServersResult = $FilteredServersResultA | Start-RSJob -Throttle 50 -Batch "Test" -ScriptBlock {
    Param($PingItem)

    if($NULL -ne (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($PingItem.Child)' AND Timeout=1000").ResponseTime)
    {
        $PingItem
    }
} | Wait-RSJob -ShowProgress -Timeout 30 | Receive-RSJob

$FailedPingServers = ForEach ($Item in $FilteredServersResultA) 
{
    If ($item.child -notin $FilteredServersResult.Child)
    {
        $Item
    }
}
#>

#=================== Create our grouped tabs in the spreadsheet

## Servers that failed Ping test
#$FailedPingServers | Export-Excel -Path $FailedPingTest -WorksheetName "Failed Ping Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

## All Servers
$FilteredServersResult | Export-Excel -Path $FilteredSpreadsheet -WorksheetName "All Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== TEST Servers

##Test Auto 10AM
$TestAutoServerList = ForEach($AutoTestServer in $FilteredServersResult) {

    #if((("$AutoTestServer.Patch Window") -like "*TestServerList*") -AND (("$AutoTestServer.Patch Window") -notlike "*Manual*"))
    if(("$AutoTestServer.Patch Window") -match "\bTestServerList\b")
    {
        $AutoTestServer
    }
}

##Test Auto EPIC 6PM
$TestAutoServersEPIC = ForEach($AutoTestEPICServer in $FilteredServersResult) {

    if((("$AutoTestEPICServer.Patch Window") -like "*Test Servers Epic*"))
    {
        $AutoTestEPICServer
    }
}

#Group up all our Manual Test Servers
$Uncategorized_Test_Server = $Test_Corepoint = $Test_ECW = $Test_Epic = $Test_Lab = $Test_Prov = $Test_Mids = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*TESTServerListManualReboot*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*AD Manager*" { $Test_Prov += $UncategorizedServer }
            "*Corepoint*" { $Test_Corepoint += $UncategorizedServer }
            "*Data Innovations*" { $Test_Lab += $UncategorizedServer }
            "*eClinicalWorks*" { $Test_Lab += $UncategorizedServer } #$Test_ECW
            "*Epic*" { $Test_Epic += $UncategorizedServer }
            "*Midas*" { $Test_Mids += $UncategorizedServer }
            "*Novanet*" { $Test_Lab += $UncategorizedServer }
            "*QML*" { $Test_Lab += $UncategorizedServer }
            "*TEG*" { $Test_Lab += $UncategorizedServer }
            Default  { $Uncategorized_Test_Server += $UncategorizedServer }
        }
    }
}

#=================== Generate our New Spreadsheet

$Test_Corepoint | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "CorePoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_ECW | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "ECW" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Epic | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "EPIC" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Lab | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "LAB" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Prov | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "PROV Workstations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Mids | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "Midas" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$TestAutoServerList | Export-Excel -Path "$TEST_AUTO_SERVER_GROUPS" -WorksheetName "Auto Reboot Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$TestAutoServersEPIC | Export-Excel -Path "$TEST_AUTO_EPIC_SERVER_GROUPS" -WorksheetName "Auto EPIC Reboot Servers (6PM)" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== PROD Servers TUE 10AM

$Prod_TUE_10AM_Auto = ForEach($Prod_TUE_10AM_AutoServer in $FilteredServersResult) {

    if((("$Prod_TUE_10AM_AutoServer.Patch Window") -like "*TuesdayAutoReboot - 10 am*"))
    {
        $Prod_TUE_10AM_AutoServer
    }
}

#Group up all our Manual Servers
Clear-Variable UncategorizedServer

$Prod_TUE_10AM_AutomatedITJobs = $Prod_TUE_10AM_Capsule = $Prod_TUE_10AM_DoseEdge = $Prod_TUE_10AM_Embla_Rembrandt = $Prod_TUE_10AM_LAN_CoreInfrastructure = $Prod_TUE_10AM_NurseCall = $Prod_TUE_10AM_SleepWorks = $Prod_TUE_10AM_BioMearieux = $Prod_TUE_10AM_BioRad = $Prod_TUE_10AM_Synapsys = $Prod_TUE_10AM_DataInnovations = $Prod_TUE_10AM_Novanet = $Prod_TUE_10AM_QML = $Prod_TUE_10AM_SCCSoftlab = $Prod_TUE_10AM_TEGManager = $Prod_TUE_10AM_VoiceBrook = $Prod_TUE_10AM_WAM = $Prod_TUE_10AM_Uncategorized = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*TuesdayManualReboot - 10 am*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*Automated IT Jobs*" { $Prod_TUE_10AM_AutomatedITJobs += $UncategorizedServer }
            "*Capsule*" { $Prod_TUE_10AM_Capsule += $UncategorizedServer }
            "*DoseEdge*" { $Prod_TUE_10AM_DoseEdge += $UncategorizedServer }
            "*Embla*" { $Prod_TUE_10AM_Embla_Rembrandt += $UncategorizedServer }
            "*Rembrandt*" { $Prod_TUE_10AM_Embla_Rembrandt += $UncategorizedServer }
            "*LAN/Core Infrastructure*" { $Prod_TUE_10AM_LAN_CoreInfrastructure += $UncategorizedServer }
            "*Nurse Call*" { $Prod_TUE_10AM_NurseCall += $UncategorizedServer }
            "*SleepWorks*" { $Prod_TUE_10AM_SleepWorks += $UncategorizedServer }
            "*Biomerieux*" { $Prod_TUE_10AM_BioMearieux += $UncategorizedServer }
            "*BioRad*" { $Prod_TUE_10AM_BioRad += $UncategorizedServer }
            "*Data Innovations*" { $Prod_TUE_10AM_DataInnovations += $UncategorizedServer }
            "*Novanet*" { $Prod_TUE_10AM_Novanet += $UncategorizedServer }
            "*QML*" { $Prod_TUE_10AM_QML += $UncategorizedServer }
            "*SCC Softlab*" { $Prod_TUE_10AM_SCCSoftlab += $UncategorizedServer }
			"*Synapsys*" { $Prod_TUE_10AM_Synapsys += $UncategorizedServer }
            "*TEG Manager*" { $Prod_TUE_10AM_TEGManager += $UncategorizedServer }
            "*Voicebrook*" { $Prod_TUE_10AM_VoiceBrook += $UncategorizedServer }
            "*WAM*" { $Prod_TUE_10AM_WAM += $UncategorizedServer }
            Default  { $Prod_TUE_10AM_Uncategorized += $UncategorizedServer }
        }
    }
}

$Prod_TUE_10AM_AutomatedITJobs | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "AutomatedITJobs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Capsule | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Capsule" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_DoseEdge | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "DoseEdge" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Embla_Rembrandt | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Embla-Rembrandt" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_LAN_CoreInfrastructure | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "LAN-CoreInfrastructure" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_NurseCall | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "NurseCall" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_SleepWorks | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "SleepWorks" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_BioMearieux | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "BioMearieux" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_BioRad | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "BioRad" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_DataInnovations | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "DataInnovations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Novanet | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Novanet" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_QML | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "QML" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_SCCSoftlab | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "SCCSoftlab" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Synapsys | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Synapsys" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_TEGManager | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "TEGManager" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_VoiceBrook | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "VoiceBrook" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_WAM | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "WAM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== PROD Servers TUE 6PM

##PROD AUTO 6PM
$Prod_TUE_6PM_AUTO = ForEach($Prod_TUE_6PM_AUTOServer in $FilteredServersResult) {

    if(((("$Prod_TUE_6PM_AUTOServer.Patch Window") -like "*TuesdayAutoReboot - 6 pm*") -OR (("$Prod_TUE_6PM_AUTOServer.Patch Window") -like "*TuesdayAutoReboot - Citrix*")) -AND ($Prod_TUE_6PM_AUTOServer.DMZ -eq $False))
    {
        $Prod_TUE_6PM_AUTOServer
    }
}

##PROD AUTO 6PM DMZ Servers
$Prod_TUE_6PM_DMZ_AUTO = ForEach($Prod_TUE_6PM_DMZ_AUTOServer in $FilteredServersResult) {

    if((("$Prod_TUE_6PM_DMZ_AUTOServer.Patch Window") -like "*TuesdayAutoReboot - 6 pm*") -AND ($Prod_TUE_6PM_DMZ_AUTOServer.DMZ -eq $True))
    {
        $Prod_TUE_6PM_DMZ_AUTOServer
    }
}

#Group up all our Manual Servers

Clear-Variable UncategorizedServer

$Prod_TUE_6PM_Capsule = $Prod_TUE_6PM_Corepoint = $Prod_TUE_6PM_Exchange = $Prod_TUE_6PM_Midas = $Prod_TUE_6PM_Spacelabs = $Prod_TUE_6PM_Stryker = $Prod_TUE_6PM_SynaptiveMedical_ImageDrive = $Prod_TUE_6PM_Uncategorized = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*Capsule*" { $Prod_TUE_6PM_Capsule += $UncategorizedServer }
            "*Corepoint*" { $Prod_TUE_6PM_Corepoint += $UncategorizedServer }
            "*Exchange*" { $Prod_TUE_6PM_Exchange += $UncategorizedServer }
            "*Midas*" { $Prod_TUE_6PM_Midas += $UncategorizedServer }
            "*Spacelabs*" { $Prod_TUE_6PM_Spacelabs += $UncategorizedServer }
            "*Stryker*" { $Prod_TUE_6PM_Stryker += $UncategorizedServer }
            "*Synaptive Medical*" { $Prod_TUE_6PM_SynaptiveMedical_ImageDrive += $UncategorizedServer }
            "*Image Drive*" { $Prod_TUE_6PM_SynaptiveMedical_ImageDrive += $UncategorizedServer }
            Default  { $Prod_TUE_6PM_Uncategorized += $UncategorizedServer }
        }
    }
}

$Prod_TUE_6PM_Capsule | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Capsule" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Corepoint | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Corepoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Exchange | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Exchange" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Midas | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Midas" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Spacelabs | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Spacelabs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Stryker | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Stryker" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_SynaptiveMedical_ImageDrive | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "SynaptiveMedical-ImageDrive" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== PROD Servers TUE 9PM

##Prod Tue 9PM Auto Servers
$Prod_TUE_9PM_Auto = ForEach($Prod_TUE_9PM_AutoServer in $FilteredServersResult) {

    if((("$Prod_TUE_9PM_AutoServer.Patch Window") -like "*TuesdayAutoReboot - 9 pm*"))
    {
        $Prod_TUE_9PM_AutoServer
    }
}

#Group up all our Manual Servers

Clear-Variable UncategorizedServer

$Prod_TUE_9PM_AtlasQA = $Prod_TUE_9PM_AVST = $Prod_TUE_9PM_CBORD_Aramark = $Prod_TUE_9PM_CRad = $Prod_TUE_9PM_MIM = $Prod_TUE_9PM_Nuance = $Prod_TUE_9PM_Obix = $Prod_TUE_9PM_PaceArt = $Prod_TUE_9PM_RayStation = $Prod_TUE_9PM_Syngo  = $Prod_TUE_9PM_SunCheck = $Prod_TUE_9PM_Vitrea = $Prod_TUE_9PM_Uncategorized = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*Atlas*" { $Prod_TUE_9PM_AtlasQA += $UncategorizedServer }
            "*AVST*" { $Prod_TUE_9PM_AVST += $UncategorizedServer }
            "*CBORD*" { $Prod_TUE_9PM_CBORD_Aramark += $UncategorizedServer }
            "*Aramark*" { $Prod_TUE_9PM_CBORD_Aramark += $UncategorizedServer }
            "*C-RAD*" { $Prod_TUE_9PM_CRad += $UncategorizedServer }
            "*MIM*" { $Prod_TUE_9PM_MIM += $UncategorizedServer }
            "*Nuance*" { $Prod_TUE_9PM_Nuance += $UncategorizedServer }
            "*Obix*" { $Prod_TUE_9PM_Obix += $UncategorizedServer }
            "*PaceArt*" { $Prod_TUE_9PM_PaceArt += $UncategorizedServer }
            "*Ray Station*" { $Prod_TUE_9PM_RayStation += $UncategorizedServer }
            "*Syngo*" { $Prod_TUE_9PM_Syngo += $UncategorizedServer }
            "*SunCheck*" { $Prod_TUE_9PM_SunCheck += $UncategorizedServer }
            "*Vitrea*" { $Prod_TUE_9PM_Vitrea += $UncategorizedServer }
            Default  { $Prod_TUE_9PM_Uncategorized += $UncategorizedServer }
        }
    }
}

$Prod_TUE_9PM_AtlasQA | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "AtlasQA" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_AVST | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Avaya-AVST" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_CBORD_Aramark | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "CBORD-Aramark" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_CRad | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "C-RAD" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_MIM | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "MIM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Nuance | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Nuance" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Obix | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Obix" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_PaceArt | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "PaceArt" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_RayStation | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "RayStation" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Syngo | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Syngo" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_SunCheck | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "SunCheck" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Vitrea | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Vitrea" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== PROD Servers THR 10AM

##Prod THR 10AM Auto Servers
$Prod_THR_10AM_Auto = ForEach($Prod_THR_10AM_AutoServer in $FilteredServersResult) {

    if((("$Prod_THR_10AM_AutoServer.Patch Window") -like "*ThursdayAutoReboot - 10 am*"))
    {
        $Prod_THR_10AM_AutoServer
    }
}

#Group up all our Manual Servers

Clear-Variable UncategorizedServer

$Prod_THR_10AM_Kronos = $Prod_THR_10AM_Spok = $Prod_THR_10AM_Uncategorized = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*ThursdayManualReboot - 10 am*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*Kronos*" { $Prod_THR_10AM_Kronos += $UncategorizedServer }
            "*Spok*" { $Prod_THR_10AM_Spok += $UncategorizedServer }
            Default  { $Prod_THR_10AM_Uncategorized += $UncategorizedServer }
        }
    }
}

$Prod_THR_10AM_Kronos | Export-Excel -Path "$PROD_MAN_THR_10AM_SERVER_GROUPS" -WorksheetName "Kronos" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_10AM_Spok | Export-Excel -Path "$PROD_MAN_THR_10AM_SERVER_GROUPS" -WorksheetName "Spok" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== PROD Servers THR 6PM

##Prod THR 6PM Auto Servers
$Prod_THR_6PM_Auto = ForEach($Prod_THR_6PM_AUTOServer in $FilteredServersResult) {

    if(((("$Prod_THR_6PM_AUTOServer.Patch Window") -like "*ThursdayAutoReboot - 6 pm*") -OR (("$Prod_THR_6PM_AUTOServer.Patch Window") -like "*ThursdayAutoReboot - Citrix*")) -AND ($Prod_THR_6PM_AUTOServer.DMZ -eq $False))
    {
        $Prod_THR_6PM_AUTOServer
    }
}

##PROD AUTO 6PM DMZ Servers
$Prod_THR_6PM_DMZ_AUTO = ForEach($Prod_THR_6PM_DMZ_AUTOServer in $FilteredServersResult) {

    if((("$Prod_THR_6PM_DMZ_AUTOServer.Patch Window") -like "*ThursdayAutoReboot - 6 pm*") -AND ($Prod_THR_6PM_DMZ_AUTOServer.DMZ -eq $True))
    {
        $Prod_THR_6PM_DMZ_AUTOServer
    }
}

#Group up all our Manual Servers

Clear-Variable UncategorizedServer

$Prod_THR_6PM_Axis_Pats = $Prod_THR_6PM_Elekta_Mosaiq = $Prod_THR_6PM_Epic = $Prod_THR_6PM_Exchange = $Prod_THR_6PM_GEPACS = $Prod_THR_6PM_Radcalc = $Prod_THR_6PM_Varian = $Prod_THR_6PM_Uncategorized = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*Axis*" { $Prod_THR_6PM_Axis_Pats += $UncategorizedServer }
            "*Elekta Mosaiq*" { $Prod_THR_6PM_Elekta_Mosaiq += $UncategorizedServer }
            "*Epic*" { $Prod_THR_6PM_Epic += $UncategorizedServer }
            "*Exchange*" { $Prod_THR_6PM_Exchange += $UncategorizedServer }
            "*GE PACS*" { $Prod_THR_6PM_GEPACS += $UncategorizedServer }
            "*Radcalc*" { $Prod_THR_6PM_Radcalc += $UncategorizedServer }
            "*RCA*" { $Prod_THR_6PM_Radcalc += $UncategorizedServer }
            "*Varian*" { $Prod_THR_6PM_Varian += $UncategorizedServer }
            Default  { $Prod_THR_6PM_Uncategorized += $UncategorizedServer }
        }
    }
}

$Prod_THR_6PM_Axis_Pats | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Axis_Pats" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_6PM_Elekta_Mosaiq | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Elekta_Mosaiq" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_6PM_Epic | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "EPIC" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_6PM_Exchange | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Exchange" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_6PM_GEPACS | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "GE-Pacs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_6PM_Radcalc | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Radcalc" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_6PM_Varian | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Varian" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#=================== PROD Servers THR 9PM

##Prod THR 9PM Auto Servers
$Prod_THR_9PM_Auto = ForEach($Prod_THR_9PM_AUTOServer in $FilteredServersResult) {

    if((("$Prod_THR_9PM_AUTOServer.Patch Window") -like "*ThursdayAutoReboot - 9 pm*"))
    {
        $Prod_THR_9PM_AUTOServer
    }
}

#Group up all our Manual Servers

Clear-Variable UncategorizedServer

$Prod_THR_9PM_3M = $Prod_THR_9PM_Dexa = $Prod_THR_9PM_GE_Pacs = $Prod_THR_9PM_OneContent_ROI = $Prod_THR_9PM_Quick_Charge = $Prod_THR_9PM_Uncategorized = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    if(("$UncategorizedServer.Patch Window") -like "*ThursdayManualReboot - 9 pm*")
    {
        Switch -Wildcard ($UncategorizedServer.Parent)
        {
            "*3M*" { $Prod_THR_9PM_3M += $UncategorizedServer }
            "*Dexa*" { $Prod_THR_9PM_Dexa += $UncategorizedServer }
            "*GE Pacs*" { $Prod_THR_9PM_GE_Pacs += $UncategorizedServer }
            "*OneContent*" { $Prod_THR_9PM_OneContent_ROI += $UncategorizedServer }
            "*ROI*" { $Prod_THR_9PM_OneContent_ROI += $UncategorizedServer }
            "*Quick Charge*" { $Prod_THR_9PM_Quick_Charge += $UncategorizedServer }
            Default  { $Prod_THR_9PM_Uncategorized += $UncategorizedServer }
        }
    }
}

$Prod_THR_9PM_3M | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "3M" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_9PM_Dexa | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "Dexa" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_9PM_GE_Pacs | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "GE-Pacs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_9PM_OneContent_ROI | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "OneContent-ROI" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_THR_9PM_Quick_Charge | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "Quick-Charge" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#========================================================= Automatically pre-populate our data into Shavlik

#Main Group
$MainGroup = "zzPatching"

#=================== REMOVE ALL MACHINE GROUPS UNDER THE MAIN GROUP TO MAKE SURE WE DON'T HAVE STALE DATA

foreach($MachineGroup in (Get-MachineGroup))
{
    if($MachineGroup.Path -like "$MainGroup*")
    {
        #$MachineGroup.Name
        Remove-MachineGroup -Name $MachineGroup.Name -Confirm:$False -ErrorAction SilentlyContinue | Out-NULL
    }
}

####=========================================
####=========================================
####==========      TEST
####=========================================
####=========================================

########## TEST GROUP VARIABLES
$TST_AUTO = "Auto Test Servers"

$TST_COREPOINT = "Manual Test CorePoint"
$TST_ECW = "Manual Test ECW"
$TST_MAN_EPIC = "Manual Test EPIC"
$TST_LAB = "Manual Test LAB"
$TST_MIDAS = "Manual Test Midas"
$TST_PROVISIONING = "Manual Test Provisioning Workstations"
$TST_Uncategorized = "Manual Test Servers"

$TST_AUTO_EPIC = "Auto Test EPIC"

#=================== CREATE OUR GROUPS IF THEY DON'T ALREADY EXIST; FOR TEST PATCHING

#Server Paths
$TEST10AM = "$MainGroup\1. TEST\01 - 10AM"
$TEST6PM = "$MainGroup\1. TEST\02 - 6PM"

#========= 10AM AUTO

Add-MachineGroup -Name $TST_AUTO -Path "$TEST10AM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 10AM MANUAL

Add-MachineGroup -Name $TST_COREPOINT -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
#Add-MachineGroup -Name $TST_ECW -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL (Part of the TST_LAB Group)
Add-MachineGroup -Name $TST_MAN_EPIC -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $TST_LAB -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $TST_PROVISIONING -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $TST_MIDAS -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $TST_Uncategorized -Path "$TEST10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#========= 6PM AUTO

Add-MachineGroup -Name $TST_AUTO_EPIC -Path "$TEST6PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#=================== IMPORT OUR COMPUTERS FROM THE SPREADSHEET

#========= 10AM AUTO

Try { Add-MachineGroupItem -Name $TST_AUTO -EndpointNames $TestAutoServerList.Child -ErrorAction Stop } Catch { }

#========= 10AM MANUAL

Try { Add-MachineGroupItem -Name $TST_COREPOINT -EndpointNames $Test_Corepoint.Child -ErrorAction Stop } Catch { }
#Try { Add-MachineGroupItem -Name $TST_ECW -EndpointNames $Test_ECW.Child -ErrorAction Stop } Catch { } (Part of the TST_LAB Group)
Try { Add-MachineGroupItem -Name $TST_MAN_EPIC -EndpointNames $Test_Epic.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $TST_LAB -EndpointNames $Test_Lab.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $TST_PROVISIONING -EndpointNames $Test_Prov.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $TST_MIDAS -EndpointNames $Test_Mids.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $TST_Uncategorized -EndpointNames $Uncategorized_Test_Server.Child -ErrorAction Stop } Catch { }

#========= 6PM AUTO

Try { Add-MachineGroupItem -Name $TST_AUTO_EPIC -EndpointNames $TestAutoServersEPIC.Child -ErrorAction Stop } Catch { }

####=========================================
####=========================================
####==========      PROD
####=========================================
####=========================================

########################## TUESDAY

########## PROD TUE GROUP VARIABLES

##10AM
$PROD_TUE_10AM_AUTO_SERVERS = "Auto Prod Tue 10AM Servers"
$PROD_TUE_10AM_MAN_AutomatedITJobs = "Manual Prod Tue 10AM AutomatedITJobs"
$PROD_TUE_10AM_MAN_Capsule = "Manual Prod Tue 10AM Capsule"
$PROD_TUE_10AM_MAN_DoseEdge = "Manual Prod Tue 10AM DoseEdge"
$PROD_TUE_10AM_MAN_Embla_Rembrandt = "Manual Prod Tue 10AM Embla/Rembrandt"
$PROD_TUE_10AM_MAN_LAN_CoreInfrastructure = "Manual Prod Tue 10AM LAN/CoreInfrastructure"
$PROD_TUE_10AM_MAN_NurseCall = "Manual Prod Tue 10AM NurseCall"
$PROD_TUE_10AM_MAN_SleepWorks = "Manual Prod Tue 10AM SleepWorks"
$PROD_TUE_10AM_MAN_BioMearieux = "Manual Prod Tue 10AM BioMearieux"
$PROD_TUE_10AM_MAN_BioRad = "Manual Prod Tue 10AM BioRad"
$PROD_TUE_10AM_MAN_DataInnovations = "Manual Prod Tue 10AM DataInnovations"
$PROD_TUE_10AM_MAN_Novanet = "Manual Prod Tue 10AM Novanet"
$PROD_TUE_10AM_MAN_QML = "Manual Prod Tue 10AM QML"
$PROD_TUE_10AM_MAN_SCCSoftlab = "Manual Prod Tue 10AM SCCSoftlab"
$PROD_TUE_10AM_MAN_Synapsys = "Manual Prod Tue 10AM Synapsys"
$PROD_TUE_10AM_MAN_TEGManager = "Manual Prod Tue 10AM TEGManager"
$PROD_TUE_10AM_MAN_VoiceBrook = "Manual Prod Tue 10AM VoiceBrook"
$PROD_TUE_10AM_MAN_WAM = "Manual Prod Tue 10AM WAM"
$PROD_TUE_10AM_MAN_Uncategorized = "Manual Prod Tue 10AM Uncategorized"

##6PM
$PROD_TUE_6PM_AUTO_SERVERS = "Auto Prod Tue 6PM Auto/Citrix Servers"
$PROD_TUE_6PM_DMZ_AUTO_SERVERS = "Auto Prod Tue 6PM DMZ Servers"
$PROD_TUE_6PM_MAN_Capsule = "Manual Prod Tue 6PM Capsule"
$PROD_TUE_6PM_MAN_CorePoint = "Manual Prod Tue 6PM CorePoint"
$PROD_TUE_6PM_MAN_Exchange = "Manual Prod Tue 6PM Exchange"
$PROD_TUE_6PM_MAN_Midas = "Manual Prod Tue 6PM Midas"
$PROD_TUE_6PM_MAN_SpaceLabs = "Manual Prod Tue 6PM SpaceLabs"
$PROD_TUE_6PM_MAN_Stryker = "Manual Prod Tue 6PM Stryker"
$PROD_TUE_6PM_MAN_SynaptiveMedical_ImageDrive = "Manual Prod Tue 6PM SynaptiveMedical-ImageDrive"
$PROD_TUE_6PM_MAN_Uncategorized = "Manual Prod Tue 6PM Uncategorized"

##Anomaly
$PROD_TUE_6PM_MAN_AVST = "Manual Prod Tue 6PM AVST"

##9PM
$PROD_TUE_9PM_AUTO_SERVERS = "Auto Prod Tue 9PM Servers"
$PROD_TUE_9PM_MAN_AtlasQA = "Manual Prod Tue 9PM AtlasQA"
$PROD_TUE_9PM_MAN_AVST = "Manual Prod Tue 9PM AVST"
$PROD_TUE_9PM_MAN_CBORD_Aramark = "Manual Prod Tue 9PM CBORD-Aramark"
$PROD_TUE_9PM_MAN_C_RAD = "Manual Prod Tue 9PM C-RAD"
$PROD_TUE_9PM_MAN_MIM = "Manual Prod Tue 9PM MIM"
$PROD_TUE_9PM_MAN_Nuance = "Manual Prod Tue 9PM Nuance"
$PROD_TUE_9PM_MAN_Obix = "Manual Prod Tue 9PM Obix"
$PROD_TUE_9PM_MAN_PaceArt = "Manual Prod Tue 9PM PaceArt"
$PROD_TUE_9PM_MAN_RayStation = "Manual Prod Tue 9PM RayStation"
$PROD_TUE_9PM_MAN_Syngo= "Manual Prod Tue 9PM Syngo"
$PROD_TUE_9PM_MAN_SunCheck = "Manual Prod Tue 9PM SunCheck"
$PROD_TUE_9PM_MAN_Vitrea = "Manual Prod Tue 9PM Vitrea"
$PROD_TUE_9PM_MAN_Uncategorized = "Manual Prod Tue 9PM Uncategorized"

#=================== CREATE OUR GROUPS IF THEY DON'T ALREADY EXIST; FOR TEST PATCHING

#Server Paths
$PRODTUE10AM = "$MainGroup\2. PROD\1. TUE\1. - 10AM"
$PRODTUE6PM = "$MainGroup\2. PROD\1. TUE\2. - 6PM"
$PRODTUE9PM = "$MainGroup\2. PROD\1. TUE\3. - 9PM"

#========= 10AM AUTO

Add-MachineGroup -Name $PROD_TUE_10AM_AUTO_SERVERS -Path "$PRODTUE10AM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 10AM MANUAL

Add-MachineGroup -Name $PROD_TUE_10AM_MAN_AutomatedITJobs -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_Capsule -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_DoseEdge -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_Embla_Rembrandt -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_LAN_CoreInfrastructure -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_NurseCall -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_SleepWorks -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_BioMearieux -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_BioRad -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_DataInnovations -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_Novanet -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_QML -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_SCCSoftlab -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_Synapsys -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_TEGManager -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_VoiceBrook -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_WAM -Path "$PRODTUE10AM\MANUAL\Lab Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_10AM_MAN_Uncategorized -Path "$PRODTUE10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#========= 6PM AUTO

Add-MachineGroup -Name $PROD_TUE_6PM_AUTO_SERVERS -Path "$PRODTUE6PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_DMZ_AUTO_SERVERS -Path "$PRODTUE6PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 6PM MANUAL

Add-MachineGroup -Name $PROD_TUE_6PM_MAN_Capsule -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_CorePoint -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_Exchange -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_Midas -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_SpaceLabs -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_Stryker -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_SynaptiveMedical_ImageDrive -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_Uncategorized -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

##Anomaly
Add-MachineGroup -Name $PROD_TUE_6PM_MAN_AVST -Path "$PRODTUE6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#========= 9PM AUTO

Add-MachineGroup -Name $PROD_TUE_9PM_AUTO_SERVERS -Path "$PRODTUE9PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 9PM MANUAL

Add-MachineGroup -Name $PROD_TUE_9PM_MAN_AtlasQA -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_AVST -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_CBORD_Aramark -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_C_RAD -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_MIM -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_Nuance -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_Obix -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_PaceArt -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_RayStation -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_Syngo -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_SunCheck -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_Vitrea -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_TUE_9PM_MAN_Uncategorized -Path "$PRODTUE9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#=================== IMPORT OUR COMPUTERS FROM THE SPREADSHEET

#========= 10AM AUTO

Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_AUTO_SERVERS -EndpointNames $Prod_TUE_10AM_Auto.Child -ErrorAction Stop } Catch { }

#========= 10AM MANUAL

Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_AutomatedITJobs -EndpointNames $Prod_TUE_10AM_AutomatedITJobs.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_Capsule -EndpointNames $Prod_TUE_10AM_Capsule.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_DoseEdge -EndpointNames $Prod_TUE_10AM_DoseEdge.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_Embla_Rembrandt -EndpointNames $Prod_TUE_10AM_Embla_Rembrandt.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_LAN_CoreInfrastructure -EndpointNames $Prod_TUE_10AM_LAN_CoreInfrastructure.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_NurseCall -EndpointNames $Prod_TUE_10AM_NurseCall.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_SleepWorks -EndpointNames $Prod_TUE_10AM_SleepWorks.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_BioMearieux -EndpointNames $Prod_TUE_10AM_BioMearieux.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_BioRad -EndpointNames $Prod_TUE_10AM_BioRad.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_DataInnovations -EndpointNames $Prod_TUE_10AM_DataInnovations.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_Novanet -EndpointNames $Prod_TUE_10AM_Novanet.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_QML -EndpointNames $Prod_TUE_10AM_QML.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_SCCSoftlab -EndpointNames $Prod_TUE_10AM_SCCSoftlab.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $Prod_TUE_10AM_MAN_Synapsys -EndpointNames $Prod_TUE_10AM_Synapsys.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_TEGManager -EndpointNames $Prod_TUE_10AM_TEGManager.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_VoiceBrook -EndpointNames $Prod_TUE_10AM_VoiceBrook.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_WAM -EndpointNames $Prod_TUE_10AM_WAM.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_10AM_MAN_Uncategorized -EndpointNames $Prod_TUE_10AM_Uncategorized.Child -ErrorAction Stop } Catch { }

#========= 6PM AUTO

Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_AUTO_SERVERS -EndpointNames $Prod_TUE_6PM_AUTO.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_DMZ_AUTO_SERVERS -EndpointNames $Prod_TUE_6PM_DMZ_AUTO.Child -ErrorAction Stop } Catch { }

#========= 6PM MANUAL

Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_Capsule -EndpointNames $Prod_TUE_6PM_Capsule.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_CorePoint -EndpointNames $Prod_TUE_6PM_Corepoint.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_Exchange -EndpointNames $Prod_TUE_6PM_Exchange.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_Midas -EndpointNames $Prod_TUE_6PM_Midas.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_SpaceLabs -EndpointNames $Prod_TUE_6PM_Spacelabs.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_Stryker -EndpointNames $Prod_TUE_6PM_Stryker.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_SynaptiveMedical_ImageDrive -EndpointNames $Prod_TUE_6PM_SynaptiveMedical_ImageDrive.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_Uncategorized -EndpointNames $Prod_TUE_6PM_Uncategorized.Child -ErrorAction Stop } Catch { }

##Anomaly
Try { Add-MachineGroupItem -Name $PROD_TUE_6PM_MAN_AVST -EndpointNames $Prod_TUE_9PM_AVST.Child -ErrorAction Stop } Catch { }

#========= 9PM AUTO

Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_AUTO_SERVERS -EndpointNames $Prod_TUE_9PM_Auto.Child -ErrorAction Stop } Catch { }

#========= 9PM MANUAL

Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_AtlasQA -EndpointNames $Prod_TUE_9PM_AtlasQA.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_AVST -EndpointNames $Prod_TUE_9PM_AVST.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_CBORD_Aramark -EndpointNames $Prod_TUE_9PM_CBORD_Aramark.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_C_RAD -EndpointNames $Prod_TUE_9PM_CRad.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_MIM -EndpointNames $Prod_TUE_9PM_MIM.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_Nuance -EndpointNames $Prod_TUE_9PM_Nuance.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_Obix -EndpointNames $Prod_TUE_9PM_Obix.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_PaceArt -EndpointNames $Prod_TUE_9PM_PaceArt.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_RayStation -EndpointNames $Prod_TUE_9PM_RayStation.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_Syngo -EndpointNames $Prod_TUE_9PM_Syngo.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_SunCheck -EndpointNames $Prod_TUE_9PM_SunCheck.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_Vitrea -EndpointNames $Prod_TUE_9PM_Vitrea.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_TUE_9PM_MAN_Uncategorized -EndpointNames $Prod_TUE_9PM_Uncategorized.Child -ErrorAction Stop } Catch { }

#####################=======================================================================================================================================##### THURSDAY

########## PROD THR GROUP VARIABLES

##10AM
$PROD_THR_10AM_AUTO_SERVERS = "Auto Prod Thr 10AM Servers"
$PROD_THR_10AM_MAN_Kronos = "Manual Prod Thr 10AM Kronos"
$PROD_THR_10AM_MAN_Spok = "Manual Prod Thr 10AM Spok"
$PROD_THR_10AM_MAN_Uncategorized = "Manual Prod Thr 10AM Uncategorized"

##6PM
$PROD_THR_6PM_AUTO_SERVERS = "Auto Prod Thr 6PM Auto/Citrix Servers"
$PROD_THR_6PM_DMZ_AUTO_SERVERS = "Auto Prod Thr 6PM DMZ Servers"
$PROD_THR_6PM_MAN_Axis_Pats = "Manual Prod Thr 6PM Axis_Pats"
$PROD_THR_6PM_MAN_Elekta_Mosaiq = "Manual Prod Thr 6PM Elekta_Mosaiq"
$PROD_THR_6PM_MAN_EPIC = "Manual Prod Thr 6PM EPIC"
$PROD_THR_6PM_MAN_Exchange = "Manual Prod Thr 6PM Exchange"
$PROD_THR_6PM_MAN_GE_Pacs = "Manual Prod Thr 6PM GE-Pacs"
$PROD_THR_6PM_MAN_Radcalc = "Manual Prod Thr 6PM Radcalc"
$PROD_THR_6PM_MAN_Varian = "Manual Prod Thr 6PM Varian"
$PROD_THR_6PM_MAN_Uncategorized = "Manual Prod Thr 6PM Uncategorized"
$PROD_THR_6PM_MAN_AVST = "Manual Prod Thr 6PM AVST"

##Anomaly
$PROD_THR_9PM_MAN_Obix = "Manual Prod Thr 9PM Obix"

##9PM
$PROD_THR_9PM_AUTO_SERVERS = "Auto Prod Thr 9PM Servers"
$PROD_THR_9PM_MAN_3M = "Manual Prod Thr 9PM 3M"
$PROD_THR_9PM_MAN_Dexa = "Manual Prod Thr 9PM Dexa"
$PROD_THR_9PM_MAN_GE_Pacs = "Manual Prod Thr 9PM GE-Pacs"
$PROD_THR_9PM_MAN_OneContent_ROI = "Manual Prod Thr 9PM OneContent-ROI"
$PROD_THR_9PM_MAN_Quick_Charge = "Manual Prod Thr 9PM Quick-Charge"
$PROD_THR_9PM_MAN_Uncategorized = "Manual Prod Thr 9PM Uncategorized"
$PROD_THR_9PM_MAN_AVST = "Manual Prod Thr 9PM AVST"

#=================== CREATE OUR GROUPS IF THEY DON'T ALREADY EXIST; FOR TEST PATCHING

#Server Paths
$PRODTHR10AM = "$MainGroup\2. PROD\2. THR\1. - 10AM"
$PRODTHR6PM = "$MainGroup\2. PROD\2. THR\2. - 6PM"
$PRODTHR9PM = "$MainGroup\2. PROD\2. THR\3. - 9PM"

#========= 10AM AUTO

Add-MachineGroup -Name $PROD_THR_10AM_AUTO_SERVERS -Path "$PRODTHR10AM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 10AM MANUAL

Add-MachineGroup -Name $PROD_THR_10AM_MAN_Kronos -Path "$PRODTHR10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_10AM_MAN_Spok -Path "$PRODTHR10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_10AM_MAN_Uncategorized -Path "$PRODTHR10AM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#========= 6PM AUTO

Add-MachineGroup -Name $PROD_THR_6PM_AUTO_SERVERS -Path "$PRODTHR6PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_DMZ_AUTO_SERVERS -Path "$PRODTHR6PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 6PM MANUAL

Add-MachineGroup -Name $PROD_THR_6PM_MAN_Axis_Pats -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_Elekta_Mosaiq -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_EPIC -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_Exchange -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_GE_Pacs -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_Radcalc -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_Varian -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_6PM_MAN_Uncategorized -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

##Anomaly
Add-MachineGroup -Name $PROD_THR_6PM_MAN_AVST -Path "$PRODTHR6PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#========= 9PM AUTO

Add-MachineGroup -Name $PROD_THR_9PM_AUTO_SERVERS -Path "$PRODTHR9PM\AUTO" -ErrorAction SilentlyContinue | Out-NULL

#========= 9PM MANUAL

Add-MachineGroup -Name $PROD_THR_9PM_MAN_3M -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_9PM_MAN_Dexa -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_9PM_MAN_GE_Pacs -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_9PM_MAN_OneContent_ROI -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_9PM_MAN_Quick_Charge -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_9PM_MAN_Uncategorized -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

##Anomaly
Add-MachineGroup -Name $PROD_THR_9PM_MAN_AVST -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $PROD_THR_9PM_MAN_Obix -Path "$PRODTHR9PM\MANUAL" -ErrorAction SilentlyContinue | Out-NULL

#=================== IMPORT OUR COMPUTERS FROM THE SPREADSHEET

#========= 10AM AUTO

Try { Add-MachineGroupItem -Name $PROD_THR_10AM_AUTO_SERVERS -EndpointNames $Prod_THR_10AM_Auto.Child -ErrorAction Stop } Catch { }

#========= 10AM MANUAL

Try { Add-MachineGroupItem -Name $PROD_THR_10AM_MAN_Kronos -EndpointNames $Prod_THR_10AM_Kronos.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_10AM_MAN_Spok -EndpointNames $Prod_THR_10AM_Spok.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_10AM_MAN_Uncategorized -EndpointNames $Prod_THR_10AM_Uncategorized.Child -ErrorAction Stop } Catch { }

#========= 6PM AUTO

Try { Add-MachineGroupItem -Name $PROD_THR_6PM_AUTO_SERVERS -EndpointNames $Prod_THR_6PM_Auto.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_DMZ_AUTO_SERVERS -EndpointNames $Prod_THR_6PM_DMZ_AUTO.Child -ErrorAction Stop } Catch { }

#========= 6PM MANUAL

Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_Axis_Pats -EndpointNames $Prod_THR_6PM_Axis_Pats.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_Elekta_Mosaiq -EndpointNames $Prod_THR_6PM_Elekta_Mosaiq.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_EPIC -EndpointNames $Prod_THR_6PM_Epic.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_Exchange -EndpointNames $Prod_THR_6PM_Exchange.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_GE_Pacs -EndpointNames $Prod_THR_6PM_GEPACS.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_Radcalc -EndpointNames $Prod_THR_6PM_Radcalc.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_Varian -EndpointNames $Prod_THR_6PM_Varian.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_AVST -EndpointNames $Prod_TUE_9PM_AVST.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_6PM_MAN_Uncategorized -EndpointNames $Prod_THR_6PM_Uncategorized.Child -ErrorAction Stop } Catch { }

#========= 9PM AUTO

Try { Add-MachineGroupItem -Name $PROD_THR_9PM_AUTO_SERVERS -EndpointNames $Prod_THR_9PM_Auto.Child -ErrorAction Stop } Catch { }

#========= 9PM MANUAL

Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_3M -EndpointNames $Prod_THR_9PM_3M.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_Dexa -EndpointNames $Prod_THR_9PM_Dexa.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_GE_Pacs -EndpointNames $Prod_THR_9PM_GE_Pacs.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_GE_Pacs -EndpointNames $Prod_THR_9PM_GE_Pacs.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_OneContent_ROI -EndpointNames $Prod_THR_9PM_OneContent_ROI.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_Quick_Charge -EndpointNames $Prod_THR_9PM_Quick_Charge.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_AVST -EndpointNames $Prod_TUE_9PM_AVST.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_Obix -EndpointNames $Prod_TUE_9PM_Obix.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $PROD_THR_9PM_MAN_Uncategorized -EndpointNames $Prod_THR_9PM_Uncategorized.Child -ErrorAction Stop } Catch { }

#####################=======================================================================================================================================##### 

##============ Updating the Main Business Services

#$Shavlik_Business_Service_Group = "Business Services (Used for Critical Updates to specific Business Services ONLY)"
$Shavlik_Business_Service_Group = "Business Services"

#=================== REMOVE ALL MACHINE GROUPS UNDER THE MAIN GROUP TO MAKE SURE WE DON'T HAVE STALE DATA

foreach($MachineGroup in (Get-MachineGroup))
{
    if($MachineGroup.Path -like "*$Shavlik_Business_Service_Group*")
    {
        #Don't remove the Workstations Group, because this is custom
        if($MachineGroup.Name -notlike "*Workstations*")
        {
            #$MachineGroup.Name
            Remove-MachineGroup -Name $MachineGroup.Name -Confirm:$False -ErrorAction SilentlyContinue | Out-NULL
        }
    }
}

#=================== Configure our Machine Groups

Clear-Variable UncategorizedServer

$All_AVST_Servers = $All_Active_Directory_Servers = $All_AD_Manager_Servers = $All_CorePoint_Capsule_SpaceLabs_Servers = $All_EPIC_Servers = $All_EPIC_MyChart_Servers = $All_Exchange_Servers = $All_Citrix_Servers = $All_LanDesk_Servers = $All_SNOW_Servers = $All_Solarwinds_Servers = $All_Spok_Servers = $All_Thycotic_Servers = @()

ForEach($UncategorizedServer in $FilteredServersResult) 
{
    Switch -Wildcard ($UncategorizedServer.Parent)
    {
        "*AVST*" { $All_AVST_Servers += $UncategorizedServer }
        "*Active Directory*" { $All_Active_Directory_Servers += $UncategorizedServer }
        "*AD Manager*" { $All_AD_Manager_Servers += $UncategorizedServer }
        "*SpaceLabs*" { $All_CorePoint_Capsule_SpaceLabs_Servers += $UncategorizedServer }
        "*Corepoint*" { $All_CorePoint_Capsule_SpaceLabs_Servers += $UncategorizedServer }
        "*Capsule*" { $All_CorePoint_Capsule_SpaceLabs_Servers += $UncategorizedServer }
        "*EPIC*" { $All_EPIC_Servers += $UncategorizedServer }
        "*Epic MyChart*" { $All_EPIC_MyChart_Servers += $UncategorizedServer }
        "*Exchange*" { $All_Exchange_Servers += $UncategorizedServer }
        "*Citrix*" { $All_Citrix_Servers += $UncategorizedServer }
        "*LanDesk Management Suite & Data Analytics*" { $All_LanDesk_Servers += $UncategorizedServer }
        "*LanDesk Service Desk*" { $All_LanDesk_Servers += $UncategorizedServer }
        "*ServiceNow*" { $All_SNOW_Servers += $UncategorizedServer }
        "*SolarWinds*" { $All_Solarwinds_Servers += $UncategorizedServer }
        "*Spok*" { $All_Spok_Servers += $UncategorizedServer }
        "*Thycotic*" { $All_Thycotic_Servers += $UncategorizedServer }
        #Default  { $Uncategorized_Test_Server += $UncategorizedServer }
    }
}

$All_EPIC_Print_Servers = ForEach($All_EPIC_Print_Server in $FilteredServersResult) {

    if(((("$All_EPIC_Print_Server.Parent") -like "*EPIC*") -AND ((("$All_EPIC_Print_Server.Child") -like "*EEPS*") -OR (("$All_EPIC_Print_Server.Child") -like "*EPMS*"))) -AND (("$All_EPIC_Print_Server.Child") -notlike "TWVS01EEPS7002"))
    {
        $All_EPIC_Print_Server
    }
}

########## PROD TUE GROUP VARIABLES

$All_AVST_Title = "All AVST Servers"
$All_ACTIVE_DIRECTORY_Title = "All Active Directory Servers"
$All_AD_Manager_Title = "All AD Manager Servers"
$All_CorePoint_Capsule_SpaceLabs_Title = "All CorePoint/Capsule/SpaceLabs Servers"
$All_EPIC_Title = "All EPIC Servers"
$All_EPIC_Print_Title = "All EPIC Print Servers"
$All_EPIC_MyChart_Title = "All EPIC MyChart Servers"
$All_Exchange_Title = "All Exchange Servers"
$All_Citrix_Title = "All Citrix Servers"
$All_LanDesk_Title = "All LanDesk Servers"
$All_SNOW_Title = "All SNOW Servers"
$All_SolarWinds_Title = "All SolarWinds Servers"
$All_Spok_Title = "All Spok Servers"
$All_Thycotic_Title = "All Thycotic Servers"

#=================== CREATE OUR GROUP IN IVANTI

#Server Paths
$Business_Service_Grp = "$Shavlik_Business_Service_Group"

Add-MachineGroup -Name $All_AVST_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_ACTIVE_DIRECTORY_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_AD_Manager_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_CorePoint_Capsule_SpaceLabs_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_EPIC_Title -Path "$Business_Service_Grp\All EPIC Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_EPIC_Print_Title -Path "$Business_Service_Grp\All EPIC Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_EPIC_MyChart_Title -Path "$Business_Service_Grp\All EPIC Servers" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_Exchange_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_Citrix_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_LanDesk_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_SNOW_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_SolarWinds_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_Spok_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL
Add-MachineGroup -Name $All_Thycotic_Title -Path "$Business_Service_Grp" -ErrorAction SilentlyContinue | Out-NULL

#=================== IMPORT OUR COMPUTERS FROM THE SPREADSHEET

Try { Add-MachineGroupItem -Name $All_AVST_Title -EndpointNames $All_AVST_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_ACTIVE_DIRECTORY_Title -EndpointNames $All_Active_Directory_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_AD_Manager_Title -EndpointNames $All_AD_Manager_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_CorePoint_Capsule_SpaceLabs_Title -EndpointNames $All_CorePoint_Capsule_SpaceLabs_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_EPIC_Title -EndpointNames $All_EPIC_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_EPIC_Print_Title -EndpointNames $All_EPIC_Print_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_EPIC_MyChart_Title -EndpointNames $All_EPIC_MyChart_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_Exchange_Title -EndpointNames $All_Exchange_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_Citrix_Title -EndpointNames $All_Citrix_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_LanDesk_Title -EndpointNames $All_LanDesk_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_SNOW_Title -EndpointNames $All_SNOW_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_SolarWinds_Title -EndpointNames $All_Solarwinds_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_Spok_Title -EndpointNames $All_Spok_Servers.Child -ErrorAction Stop } Catch { }
Try { Add-MachineGroupItem -Name $All_Thycotic_Title -EndpointNames $All_Thycotic_Servers.Child -ErrorAction Stop } Catch { }

#####################=======================================================================================================================================#####

# Generate our Compare Machine Group for the Month that will be used for Validation (IGNORE ALL LEGACY SERVERS (2003/2008))

#$DateTitle =  Get-Date -UFormat "(%B %Y)"
#$Shavlik_Validation_Report = "Shavlik Compare - Windows Servers " + $DateTitle
$Shavlik_Validation_Report = "Shavlik Compare - Windows Servers"

#Remove our Group if it already exists
Remove-MachineGroup -Name $Shavlik_Validation_Report -Confirm:$False -ErrorAction SilentlyContinue | Out-NULL

#Create the Group we need of all servers that aren't 2003/2008
$Shavlik_Validation_Servers = ForEach($Shavlik_Validation_Server in $FilteredServersResult) {

    if((("$Shavlik_Validation_Server.Operating System") -notlike "*2003*") -AND (("$Shavlik_Validation_Server.Operating System") -notlike "*2008*"))
    {
        $Shavlik_Validation_Server
    }
}

#Create our group in Ivanti
Add-MachineGroup -Name $Shavlik_Validation_Report -ErrorAction SilentlyContinue | Out-NULL

#Populate our report with all the servers we want
Try { Add-MachineGroupItem -Name $Shavlik_Validation_Report -EndpointNames $Shavlik_Validation_Servers.Child -ErrorAction Stop } Catch { }

$End =  (Get-Date)

$End - $Start
