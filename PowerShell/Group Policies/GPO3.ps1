CLS

#Get-GPResultantSetOfPolicy -ReportType XML -Path "C\temp\report.xml"

$AllGPOS = Get-GPO -All

$Results = @()

Foreach($GPO in $AllGPOS)
{
    #$Name = ($GPO.DisplayName).Trim()
    #$GUID = $GPO.ID

    If ( $GPO | Get-GPOReport -ReportType XML | Select-String -NotMatch "<LinksTo>" ) 
    {
        $Results += $GPO
    }

   # Pause

    #$Name | sort
}

$Results | Sort DisplayName