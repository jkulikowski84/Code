CLS

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

#------------------------------------ Remove Old Filtered Spreadsheet

$NewFilteredWorkbook = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Filtered-Spreadsheet.xlsx"

if((Test-Path $NewFilteredWorkbook) -eq $True)
{
    Remove-Item $NewFilteredWorkbook
}

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName
#$ExcelFile = (Get-ChildItem -Path "$Path\Test Servers April 2023.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1 -AsDate "Most recent discovery","Last Reboot Date"

#------------------------------------ Remove Duplicate entries and sort by Name

$SortedExcelServersList = ($ExcelServers | Sort-Object -Property Child -Unique)

#------------------------------------ Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if(($($SortedExcelServerList.DMZ) -eq $true) -AND ($($SortedExcelServerList.child) -notlike "*.dmz.com"))
    {
        $SortedExcelServerList.child = [System.String]::Concat("$($SortedExcelServerList.child)",".dmz.com")
    }

    $SortedExcelServerList
}

#------------------------------------ Grab all servers from AD so we can use to compare against our list - also trim any whitespaces from output

$Servers = (dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0 -attr Name | sort).trim()

#------------------------------------ Compare our list to servers in AD and filter out appliances

$FilteredServersResult = $Null

$FilteredServersResult = ForEach ($Item in $FilteredServers) 
{
    If (($item.child -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

#------------------------------------ Get All servers that aren't DMZ or Citrix

$AllServers = ForEach($FilteredServerResult in $FilteredServersResult) {

    #Don't include DMZ or Citrix servers in our list
    if(($($FilteredServerResult.DMZ) -eq $false) -AND ($($FilteredServerResult.Parent) -notlike "*Citrix*"))
    {
        $FilteredServerResult
    }
}

#------------------------------------ Get Servers patched on Tuesday

$AllServersTues = ForEach($AllServer in $AllServers) {

    if(("$AllServer.Patch Window") -like "*Tuesday*")
    {
        $AllServer
    }
}

#------------------------------------ Get Servers patched on Thursday

$AllServersThur = ForEach($AllServer in $AllServers) {

    if(("$AllServer.Patch Window") -like "*Thursday*")
    {
        $AllServer
    }
}

#------------------------------------ Seperate Servers from DMZ Servers

$DMZServers = ForEach($FilteredServerResult in $FilteredServersResult) {

    if($($FilteredServerResult.DMZ) -eq $true)
    {
        if($($FilteredServerResult.child) -notlike "*.dmz.com")
        {
            $FilteredServerResult.Child = [System.String]::Concat("$($FilteredServerResult.Child)",".dmz.com")
        }
        $FilteredServerResult
    }
}

#------------------------------------ Get DMZ Servers patched on Tuesday

$DMZServersTues = ForEach($DMZServer in $DMZServers) {

    if(("$DMZServer.Patch Window") -like "*Tuesday*")
    {
        $DMZServer
    }
}

#------------------------------------ Get DMZ Servers patched on Thursday

$DMZServersThur = ForEach($DMZServer in $DMZServers) {

    if(("$DMZServer.Patch Window") -like "*Thursday*")
    {
        $DMZServer
    }
}

#------------------------------------ Seperate Servers from Citrix Servers

$CitrixServers = ForEach($FilteredServerResult in $FilteredServersResult) {

    if($($FilteredServerResult.Parent) -like "*Citrix*")
    {
        $FilteredServerResult
    }
}

#------------------------------------ Get Citrix Servers patched on Tuesday

$CitrixServersTues = ForEach($CitrixServer in $CitrixServers) {

    if(("$CitrixServer.Patch Window") -like "*Tuesday*")
    {
        $CitrixServer
    }
}

#------------------------------------ Get Citrix Servers patched on Thursday

$CitrixServersThur = ForEach($CitrixServer in $CitrixServers) {

    if(("$CitrixServer.Patch Window") -like "*Thursday*")
    {
        $CitrixServer
    }
}

#Create our new formatted Spreadsheet 

$AllServers | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "All Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$AllServersTues | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "Servers Tues" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$AllServersThur | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "Servers Thur" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#---

$DMZServers | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "All DMZ Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$DMZServersTues | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "DMZ Servers Tues" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$DMZServersThur | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "DMZ Servers Thur" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#---

$CitrixServers | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "All Citrix Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$CitrixServersTues | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "Citrix Servers Tues" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$CitrixServersThur | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "Citrix Servers Thur" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"


$End =  (Get-Date)

$End - $Start
