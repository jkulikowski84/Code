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

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.csv"

#===================  Remove our Software List if it exists

Remove-Item -Path "$Path\SoftwareList.xlsx" -Force -Confirm:$False -ErrorAction SilentlyContinue

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

#====================== Calculate dates for the invite
    
#Get current Month and Year
$Month = (get-Date).month
$Year = (get-Date).year

#Get the first day of the month
$FirstDayofMonth = [datetime] ([string]$Month + "/1/" + [string]$Year)

#Grab the Wednesday after Windows Patch Tuesday (2nd Tuesay of the month) for Test Patching
$Wed = (0..30 | % {$firstdayofmonth.adddays($_) } | ? {$_.dayofweek -like "Tue*"})[1].AddDays(1)

#Prod Patching *USUALLY* takes place on the 3rd week (after Test Patching) on Tuesday and Thursday
$Tue = ($Wed).AddDays(6)

#Check if the date for Tue Prod Patching is correct.
Write-output "Tue Prod Patching is scheduled for $($Tue.ToString('MM/dd/yyyy'))"
$ProdDay1 = Read-Host -Prompt "If the above date is correct, hit 'Enter' to continue, otherwise type in the number of the day test patching will take place on."

if([string]::IsNullOrEmpty($ProdDay1))
{
    #Do nothing, just continue
}
else
{
    $Tue = Get-Date -Date "$($year)-$($month)-$($ProdDay1)T00:00:00"
}

CLS

$Thr = ($Tue).AddDays(2)

#Check if the date for Thr Prod Patching is correct.
Write-output "Thr Prod Patching is scheduled for $($Thr.ToString('MM/dd/yyyy'))"
$ProdDay2 = Read-Host -Prompt "If the above date is correct, hit 'Enter' to continue, otherwise type in the number of the day test patching will take place on."

if([string]::IsNullOrEmpty($ProdDay2))
{
    #Do nothing, just continue
}
else
{
    $Thr = Get-Date -Date "$($year)-$($month)-$($ProdDay2)T00:00:00"
}

CLS

if($Tue -ge $(Get-Date))
{
    Write-Output "Tue Prod Patching will take place on: $($Tue.ToString('MM/dd/yyyy'))"
}

if($Thr -ge $(Get-Date))
{
    Write-Output "Thr Prod Patching will take place on: $($Thr.ToString('MM/dd/yyyy'))"
}

#=================== Create our Table

$TueSoftware10AM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Tuesday* - 10 am"}).Parent | sort -Unique
$TueSoftware6PM = ($SortedExcelServersList | Where-Object { ($_."Patch Window" -like "Tuesday* - 6 pm") -OR ($_."Patch Window" -like "TuesdayAutoReboot - Citrix") } ).Parent | sort -Unique
$TueSoftware9PM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Tuesday* - 9 pm"}).Parent | sort -Unique
$ThrSoftware10AM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Thursday* - 10 am"}).Parent | sort -Unique
$ThrSoftware6PM = ($SortedExcelServersList | Where-Object { ($_."Patch Window" -like "Thursday* - 6 pm") -OR ($_."Patch Window" -like "ThursdayAutoReboot - Citrix") } ).Parent | sort -Unique
$ThrSoftware9PM = ($SortedExcelServersList | Where-Object { $_."Patch Window" -like "Thursday* - 9 pm"}).Parent | sort -Unique

#=================== Export our data into the spreadsheet

##first Column are just row headings
"Day"         | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 1 -StartRow 1 -AutoSize -BoldTopRow -FreezeFirstColumn
"Date"        | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 1 -StartRow 2 -AutoSize -BoldTopRow
"Time"        | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 1 -StartRow 3 -AutoSize -BoldTopRow
"Applications" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 1 -StartRow 4 -AutoSize -BoldTopRow

##Column Headings
"Tuesday"          | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 2 -StartRow 1 -AutoSize -FreezeTopRow -BoldTopRow
"Tuesday"          | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 3 -StartRow 1 -AutoSize -FreezeTopRow -BoldTopRow
"Tuesday"          | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 4 -StartRow 1 -AutoSize -FreezeTopRow -BoldTopRow
"Thursday"         | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 5 -StartRow 1 -AutoSize -FreezeTopRow -BoldTopRow
"Thursday"         | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 6 -StartRow 1 -AutoSize -FreezeTopRow -BoldTopRow
"Thursday"         | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 7 -StartRow 1 -AutoSize -FreezeTopRow -BoldTopRow

"$((Get-Date -Date $Tue).Day)-$((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName((Get-Date).Month))" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 2 -StartRow 2 -AutoSize -BoldTopRow
"$((Get-Date -Date $Tue).Day)-$((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName((Get-Date).Month))" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 3 -StartRow 2 -AutoSize -BoldTopRow
"$((Get-Date -Date $Tue).Day)-$((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName((Get-Date).Month))" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 4 -StartRow 2 -AutoSize -BoldTopRow
"$((Get-Date -Date $Thr).Day)-$((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName((Get-Date).Month))" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 5 -StartRow 2 -AutoSize -BoldTopRow
"$((Get-Date -Date $Thr).Day)-$((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName((Get-Date).Month))" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 6 -StartRow 2 -AutoSize -BoldTopRow
"$((Get-Date -Date $Thr).Day)-$((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName((Get-Date).Month))" | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 7 -StartRow 2 -AutoSize -BoldTopRow

($TueSoftware10AM | Select-Object @{ Name='10 AM - 12 PM';  Expression={ "•         $_" } }) | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 2 -StartRow 3 -AutoSize -BoldTopRow
($TueSoftware6PM | Select-Object  @{ Name='6 PM - 8 PM';    Expression={ "•         $_" } }) | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 3 -StartRow 3 -AutoSize -BoldTopRow
($TueSoftware9PM | Select-Object  @{ Name='9 PM - 11 PM';   Expression={ "•         $_" } }) | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 4 -StartRow 3 -AutoSize -BoldTopRow
($ThrSoftware10AM | Select-Object @{ Name='10 AM - 12 PM';  Expression={ "•         $_" } }) | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 5 -StartRow 3 -AutoSize -BoldTopRow
($ThrSoftware6PM | Select-Object  @{ Name='6 PM - 8 PM';    Expression={ "•         $_" } }) | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 6 -StartRow 3 -AutoSize -BoldTopRow
($ThrSoftware9PM | Select-Object  @{ Name='9 PM - 11 PM';   Expression={ "•         $_" } }) | Export-Excel -Path "SoftwareList.xlsx" -WorksheetName "Software Groups" -StartColumn 7 -StartRow 3 -AutoSize -BoldTopRow

#=================== Manipulate data from our spreadsheet

$SpreadsheetGrpFile = (Get-ChildItem -Path "$Path\SoftwareList.xlsx").FullName

$SpreadsheetGrpOpen = Open-ExcelPackage -Path $SpreadsheetGrpFile

$SpreadsheetGrpWorksheet = $SpreadsheetGrpOpen.Workbook.Worksheets[0]



#$(($SpreadsheetGrpWorksheet.Columns[2].Range | Select-Object -first 1).LocalAddress)

##Center Alignment A1-A3 : G1-G3
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "A1:G1" -HorizontalAlignment Center -FontName "Times New Roman" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "A2:G2" -HorizontalAlignment Center -FontName "Times New Roman" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "A3:G3" -HorizontalAlignment Center -FontName "Times New Roman" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))

#Make font white in first column
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "A1:A4" -HorizontalAlignment Center -FontColor "White" -FontName "Times New Roman" 

##Draw borders
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "A1:A3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "B1:B3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "C1:C3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "D1:D3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "E1:E3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "F1:F3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "G1:G3" -BorderAround Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "A4:G$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BorderRight Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "B$($SpreadsheetGrpWorksheet.Dimension.End.Row):G$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BorderBottom Thin -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#ffc000"))

#Color our Cells
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[1].Range | Select-Object -first 1).LocalAddress):A$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#70ad47"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[2].Range | Select-Object -first 1).LocalAddress):B$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#e6eed5"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[3].Range | Select-Object -first 1).LocalAddress):C$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[4].Range | Select-Object -first 1).LocalAddress):D$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#e6eed5"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[5].Range | Select-Object -first 1).LocalAddress):E$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[6].Range | Select-Object -first 1).LocalAddress):F$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#e6eed5"))
Set-Format -WorkSheet $SpreadsheetGrpWorksheet -Range "$(($SpreadsheetGrpWorksheet.Columns[7].Range | Select-Object -first 1).LocalAddress):G$($SpreadsheetGrpWorksheet.Dimension.End.Row)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))

## Close out the file once we're done modifying spreadsheet

Close-ExcelPackage $SpreadsheetGrpOpen -Show
