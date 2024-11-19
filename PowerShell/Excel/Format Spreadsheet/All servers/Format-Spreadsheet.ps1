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

#------------------------------------ Populate our variable with data from spreadsheet

<#
$ExcelServersList = foreach($ExcelServer in $ExcelServers) {
    $ExcelServer | Select-Object @{Name="ServerName";Expression={$_.Child}}, "Primary", @{Name="PatchWindow";Expression={$_."Patch Window"}}, @{Name="TestServer";Expression={$_."Test Server"}}, "DMZ" 
}
#>
#------------------------------------ Remove Duplicate entries

$SortedExcelServersList = ($ExcelServers | Sort-Object -Property Child -Unique)

#------------------------------------ Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if($($SortedExcelServerList.DMZ) -eq $true)
    {
        $SortedExcelServerList.child = [System.String]::Concat("$($SortedExcelServerList.child)",".dmz.com")
    }

    $SortedExcelServerList
}

#------------------------------------ Grab all servers from AD so we can use to compare against our list - also trimany whitespaces from output

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

#Create our new formatted Spreadsheet 

$FilteredServersResult | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "All Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$End =  (Get-Date)

$End - $Start
