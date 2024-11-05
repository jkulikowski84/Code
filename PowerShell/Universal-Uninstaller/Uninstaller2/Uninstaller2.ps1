CLS

$Software = "ABC"
$Filter = "*" + $Software + "*"
$Program = $ProgUninstall = $NULL

try 
{
    if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node") 
    {
        $programs = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction Stop
    }

    $programs += Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction Stop
    $programs += Get-ItemProperty -Path "Registry::\HKEY_USERS\*\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
} 
catch 
{
    Write-Error $_
    break
}

foreach($Program in $Programs)
{
    $ProgDisplayName = $Program.DisplayName
    $ProgUninstall = $Program.UninstallString

    if(($ProgDisplayName -like $Filter) -and ($NULL -ne $ProgUninstall))
    {

        $aux = $ProgUninstall -split @('\.exe'),2,[System.StringSplitOptions]::None
        $Uninstaller = (cmd /c echo $($aux[0].TrimStart('"').TrimStart("'") + '.exe')).Trim('"')
        $UninsParams = $aux[1].TrimStart('"').TrimStart("'").Trim().split(' ',[System.StringSplitOptions]::RemoveEmptyEntries)

        #Debug
        #$Uninstaller
        #$UninsParams

        . $Uninstaller $UninsParams | Where-Object { $_ -notlike "param 0 = *" }
    }
}
