CLS

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

#Clear our Variables
Clear-Variable FilteredOutPrinters, FilteredPrinters, FailedPingedPrinters, FailedPingZebraPrinters -Force -Confirm:$False -ErrorAction SilentlyContinue


$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

#------------------------------------  Convert our CSV to a compatible XLSX file

$CSVFile = (Get-ChildItem -Path "$Path\*.csv").FullName

$ImportedData = Import-Csv -Path $CSVFile

#Skip if the file already exists
if((Test-Path "$($CSVFile -replace ".csv",".xlsx")") -eq $False)
{
    $ImportedData | Export-Excel "$($CSVFile -replace ".csv",".xlsx")"
}

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelPrinters = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

#------------------------------------ Populate our variable with data from spreadsheet

<#
$ExcelPrinterList = foreach($ExcelPrinter in $ExcelPrinters) 
{
    $ExcelPrinter | Select-Object "Port Address"
}

$ExcelPrinterList
#>

#------------------------------------ Filter out Zebra printers that are not in the _EMR section (Display those printers)

$FilteredOutPrinters = ForEach ($Item in $ExcelPrinters) 
{
    If(($item."Printer Name" -like "*zs*") -and ($item."Printer Folder" -notlike "*_EMR*"))
    {
        $Item
    }
}

#------------------------------------ Filter out Zebra printers that are not in the _EMR section (Display all printers except those)

$FilteredPrinters = ForEach ($Item in $ExcelPrinters) 
{
    If(-not (($item."Printer Name" -like "*zs*") -and ($item."Printer Folder" -notlike "*_EMR*")))
    {
        $Item
    }
}

#------------------------------------ Ping the printers to see if they are active/online

$ping = New-Object System.Net.NetworkInformation.Ping

$FailedPingedPrinters = $FilteredPrinters | Start-RSJob -Throttle 50 -Batch "Test" -ScriptBlock {
    Param($PingItem)

    #if($NULL -ne (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($PingItem."Port Address")' AND Timeout=1000").ResponseTime)
    if($NULL -eq (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($PingItem."Port Address")' AND Timeout=1000").ResponseTime)
    {
        $PingItem
    }
} | Wait-RSJob -ShowProgress -Timeout 30 | Receive-RSJob

#Output Failed Pinged printers by name
#$FailedPingedPrinters."Printer Name"

#Export list to Spreadsheet
$FailedPingedPrinters | Export-Excel -Path "$Path\Exports\FailedPingedPrinters.xlsx" -WorksheetName "Unresponsive Printers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#------------------------------------ Grab all zebras that failed the Ping test from the list above

$FailedPingZebraPrinters = ForEach ($Item in $FailedPingedPrinters) 
{
    If($item."Printer Name" -like "*zs*")
    {
        $Item
    }
}

#Output Failed Pinged Zebra Printers by name
#$FailedPingZebraPrinters | Export-Excel -Path "$Path\Exports\FailedPingZebraPrinters.xlsx" -WorksheetName "Unresponsive Zebra Printers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
