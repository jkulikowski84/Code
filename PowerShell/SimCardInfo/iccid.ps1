$outfile="$env:USERPROFILE\Desktop\MEID_ICCID.txt"
$get_meid=cmd /c "netsh mbn sh interface"
$get_iccid=cmd /c "netsh mbn sh read i=*"

If($get_meid -eq "There is no Mobile Broadband interface")
{
    Write-Host "No Mobile Broadband Interface Found"
    Read-Host 'Press Enter to Continue' | Out-Null
    Exit
}

"GOBI Information for $env:COMPUTERNAME" | Out-File $outfile

Function Hide_and_Seek_MEID
{
    Foreach($entry in $get_meid)
    {
        If($entry.Contains($args[0]))
        {
            $entry | Out-File $outfile -Append
        }
    }
}

Function Hide_and_Seek_ICCID
{
    Foreach($entry in $get_iccid)
    {
        If($entry.Contains($args[0]))
        {
            $entry | Out-File $outfile -Append
        }
    }
}

"" | Out-File $outfile -Append
"MEID Info:" | Out-File $outfile -Append

Hide_and_Seek_MEID "Description"
Hide_and_Seek_MEID "Physical Address"
Hide_and_Seek_MEID "Cellular Class"
Hide_and_Seek_MEID "Device Id"
Hide_and_Seek_MEID "Manufacturer"
Hide_and_Seek_MEID "Model"
Hide_and_Seek_MEID "Firmware Version"

"" | Out-File $outfile -Append
"ICCID Info:" | Out-File $outfile -Append

Hide_and_Seek_ICCID "Subscriber Id"
Hide_and_Seek_ICCID "SIM ICC Id"