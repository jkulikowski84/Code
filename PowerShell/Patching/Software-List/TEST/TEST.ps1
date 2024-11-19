CLS

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

#Remove our test spreadsheet if it already exists
Remove-Item -Path "$Path\Test.xlsx" -Force -Confirm:$False -ErrorAction SilentlyContinue

#Get our main spreadsheet
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#make a copy of our spreadsheet
Copy-Item -Path $ExcelFile -Destination "$Path\Test.xlsx"

#Get our test copy
$TestFile = (Get-ChildItem -Path "$Path\Test.xlsx").FullName

$excel = Open-ExcelPackage -Path $TestFile

$Worksheet = $Excel.Workbook.Worksheets[0]

#Set Color for each column

Set-Format -WorkSheet $Worksheet -Range "$(($Worksheet.Columns[1].Range | Select-Object -first 1).LocalAddress):$(($Worksheet.Columns[1].Range | Select-Object -Last 1).LocalAddress)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#e6eed5"))
Set-Format -WorkSheet $Worksheet -Range "$(($Worksheet.Columns[2].Range | Select-Object -first 1).LocalAddress):$(($Worksheet.Columns[2].Range | Select-Object -Last 1).LocalAddress)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
Set-Format -WorkSheet $Worksheet -Range "$(($Worksheet.Columns[3].Range | Select-Object -first 1).LocalAddress):$(($Worksheet.Columns[3].Range | Select-Object -Last 1).LocalAddress)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#e6eed5"))
Set-Format -WorkSheet $Worksheet -Range "$(($Worksheet.Columns[4].Range | Select-Object -first 1).LocalAddress):$(($Worksheet.Columns[4].Range | Select-Object -Last 1).LocalAddress)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
Set-Format -WorkSheet $Worksheet -Range "$(($Worksheet.Columns[5].Range | Select-Object -first 1).LocalAddress):$(($Worksheet.Columns[5].Range | Select-Object -Last 1).LocalAddress)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#e6eed5"))
Set-Format -WorkSheet $Worksheet -Range "$(($Worksheet.Columns[6].Range | Select-Object -first 1).LocalAddress):$(($Worksheet.Columns[6].Range | Select-Object -Last 1).LocalAddress)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))

Close-ExcelPackage $excel -Show
