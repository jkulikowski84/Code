cls

#Set the Parameters

$ShortcutLocation = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Integration for Allscripts.lnk'

$TargetPath = 'C:\Program Files (x86)\Hyland\Integrations\Allscripts\Hyland.Applications.Allscripts.exe'

#Field Mode Arguments
$Arguments = '-headless'

#Host Mode Arguments
#$Arguments = '-headless -connected'

$WorkingDirectory = "C:\Program Files (x86)\Hyland\Integrations\Allscripts\"

#Remove the existing shortcut
if(Test-Path $ShortcutLocation){
	Remove-Item $ShortcutLocation
}

#Create the new shortcut
$Shell = New-Object -ComObject WScript.Shell; 

$Shortcut = $Shell.CreateShortcut($ShortcutLocation); 

$Shortcut.TargetPath = $TargetPath

$Shortcut.Arguments = $Arguments

$Shortcut.WorkingDirectory = $WorkingDirectory

$Shortcut.Save()
