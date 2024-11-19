CLS

#=================== Import Modules

#List the modules we want to load per line
$LoadModules = @'
ImportExcel
'@.Split("`n").Trim()

ForEach($Module in $LoadModules) 
{
    if($NULL -eq (Get-Module -ListAvailable $Module))
    {
        Import-Module $Module
    }
}

#===================

#Start Timestamp
$Start = Get-Date

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.csv"

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
<#
$TueSoftware10AM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Tuesday* - 10 am"}).Parent | sort -Unique
$TueSoftware6PM = ($SortedExcelServersList | Where-Object { ($_."Patch Window" -like "Tuesday* - 6 pm") -OR ($_."Patch Window" -like "TuesdayAutoReboot - Citrix") } ).Parent | sort -Unique
$TueSoftware9PM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Tuesday* - 9 pm"}).Parent | sort -Unique
$ThrSoftware10AM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Thursday* - 10 am"}).Parent | sort -Unique
$ThrSoftware6PM = ($SortedExcelServersList | Where-Object { ($_."Patch Window" -like "Thursday* - 6 pm") -OR ($_."Patch Window" -like "ThursdayAutoReboot - Citrix") } ).Parent | sort -Unique
$ThrSoftware9PM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Thursday* - 9 pm"}).Parent | sort -Unique
#>
Clear-Variable SoftwareGroups -Force -Confirm:$False -ErrorAction SilentlyContinue

<#
$SoftwareGroups = [PSCustomObject]@{
#'Tuesday 10 AM - 12PM' = "$($TueSoftware10AM)`n"
'Tuesday 10 AM - 12PM' = $TueSoftware10AM
'Tuesday 6 PM - 8 PM' = $TueSoftware6PM
'Tuesday 9 PM - 11 PM' = $TueSoftware9PM
'Thursday 10 AM - 12PM' = $ThrSoftware10AM
'Thursday 6 PM - 8 PM' = $ThrSoftware6PM
'Thursday 9 PM - 11 PM' = $ThrSoftware9PM
}
#>


$SoftwareGroups = ForEach($SoftwareGrp in $SortedExcelServersList) 
{
    Switch -Wildcard ($SoftwareGrp."Patch Window")
    {
        "Tuesday* - 10 am" { [PSCustomObject] @{ "Tuesday 10 am – 12 pm" = $SoftwareGrp.Parent}  }
        "Tuesday* - 6 pm" { [PSCustomObject] @{ "Tuesday 6 pm – 8 pm" = $SoftwareGrp.Parent } }
        "TuesdayAutoReboot - Citrix" {}
        "Tuesday* - 9 pm" {}
        "Thursday* - 10 am" {}
        "Thursday* - 6 pm" {}
        "ThursdayAutoReboot - Citrix" {}
        "Thursday* - 9 pm" {}
    }
}

$SoftwareGroups | ft

<#
$TueSoftware10AM | ForEach-Object {
    $SoftwareGroups += [PSCustomObject] @{
        "Tuesday 10 am – 12 pm" = "$_"
    }
}

$TueSoftware6PM | ForEach-Object {
    $SoftwareGroups += [PSCustomObject] @{
        "Tuesday 6 pm – 8 pm" = "$_"
    }
}

$SoftwareGroups | ft

#$SoftwareGroups | ft

#$SoftwareGroups | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "Software Groups" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle "Light1"

<#
$TueSoftware10AM | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "TUE 10AM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$TueSoftware6PM | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "TUE 6PM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$TueSoftware9PM | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "TUE 9PMM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$ThrSoftware10AM | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "THR 10AM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$ThrSoftware6PM | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "THR 6PM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$ThrSoftware9PM | Export-Excel -Path "SoftwareGroups.xlsx" -WorksheetName "THR 9PM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
#>

$End =  (Get-Date)

$End - $Start
