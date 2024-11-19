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

$DMZServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if($($SortedExcelServerList.DMZ) -eq $true)
    {
        $SortedExcelServerList.Child = [System.String]::Concat("$($SortedExcelServerList.Child)",".dmz.com")
        $SortedExcelServerList
    }
}

#------------------------------------ Seperate Servers from Citrix Servers

$CitrixServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if($($SortedExcelServerList.Parent) -like "Citrix GenApp")
    {
        $SortedExcelServerList
    }
}

$SortedExcelServersList | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "All Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$DMZServers | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "DMZ Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$CitrixServers | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "Citrix Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$End =  (Get-Date)

$End - $Start
