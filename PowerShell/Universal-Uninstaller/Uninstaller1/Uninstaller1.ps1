CLS

$Software = "ABC"
$Filter = "*" + $Software + "*"
$Program = $ProgUninstall = $FileUninstaller = $FileArg = $NULL

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
        if($ProgUninstall -like "msiexec*")
        {
            $FileUninstaller = $ProgUninstall.split(" ")[0]
            $FileArg = ($($ProgUninstall).split(" ",2)[1])
        }
        else
        {
            if(($ProgUninstall -like '"*"*') -or ($ProgUninstall -like "'*'*"))
            {
                #String has quotes, don't need to do anything
            }
            else
            {
                if($NULL -ne $ProgUninstall)
                {
                    #String doesn't have quotes so we should add them
                    $ProgUninstall = '"' + ($ProgUninstall.Replace('.exe','.exe"'))                    
                }
            }

            #Let's grab the uninstaller and arguments
            $FileUninstaller = $ProgUninstall.split('"')[1]
            $FileArg = $ProgUninstall.split('"')[-1]
        }

        #Debug
        #$FileUninstaller
        #$FileArg

        #Run the Uninstaller
        Start-Process $FileUninstaller -ArgumentList $FileArg -wait -ErrorAction SilentlyContinue
    }
}