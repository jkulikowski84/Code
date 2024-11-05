CLS

#Type in the module name that you want to scan
$ModuleName = "VMware.VimAutomation.Core"

#Clear out Variables
Clear-Variable MainModule, MainPSD1File, Content, line, RequiredModule -Force -Confirm:$False -ErrorAction SilentlyContinue

#Get the main module we want to load
$MainModule = $env:HOMEDRIVE + $env:HOMEPATH + "\Documents\WindowsPowerShell\Modules\$($ModuleName)"

do
{
    $MainPSD1File = [System.IO.Directory]::EnumerateFiles($MainModule, "*.psd1", [System.IO.SearchOption]::AllDirectories) | select-object -First 1

    $Content = [System.IO.File]::ReadAllLines($MainPSD1File)

    #Open the file and find the required Module
    $line = $Content | Select-Object | Where-Object {$_ -like "*ModuleName*"}

    if(($line | measure).count -eq 1)
    {
        $RequiredModule = $line.split('"')[3]
        $RequiredModule
    }
    else
    {
        foreach($i in $line)
        {
            $i.split('"')[3]
        }
    }

    #Set our new path
    $MainModule = $env:HOMEDRIVE + $env:HOMEPATH + "\Documents\WindowsPowerShell\Modules\$($RequiredModule)"

} until(($NULL -eq $line))
