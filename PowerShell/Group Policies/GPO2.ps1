CLS

$GPOs = Get-GPO -All
$OutputFile = "C:\TEMP\GPOList.txt"

"Name;LinkPath;ComputerEnabled;UserEnabled;WmiFilter" | Out-File $OutputFile

$GPOs | % {
     [xml]$Report = $_ | Get-GPOReport -ReportType XML
     $Links = $Report.GPO.LinksTo

     ForEach($Link In $Links)
     {
         $Output = $Report.GPO.Name + ";" + $Link.SOMPath + ";" + $Report.GPO.Computer.Enabled + ";" + $Report.GPO.User.Enabled + ";" + $_.WmiFilter.Name
         $Output | Out-File $OutputFile -Append
     }
}
