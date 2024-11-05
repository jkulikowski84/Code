CLS

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

#If Teams is running, close it out.
Stop-Process -Name "Teams" -Force -ErrorAction SilentlyContinue

#Close out of Outlook as well because it has a teams add-in
Stop-Process -Name "Outlook" -Force  -ErrorAction SilentlyContinue

#Remove shortcut(s) is they exist
if(Test-Path("$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic.lnk")) { Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic.lnk" -Force }
if(Test-Path("$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk")) { Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk" -Force }

#Get logged on users SID
$usersid = (New-Object -ComObject Microsoft.DiskQuota).TranslateLogonNameToSID((Get-WmiObject -Class Win32_ComputerSystem).Username)
$username = ((Get-WmiObject -Class Win32_ComputerSystem).Username).split("\")[1]

#Remove all reg keys associated with Microsoft Teams
Remove-Item -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\AAB6F137689A4A549863C7A3AAAA67B0" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "HKLM:\SOFTWARE\Classes\Installer\Products\AAB6F137689A4A549863C7A3AAAA67B0" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "registry::HKEY_CLASSES_ROOT\Installer\Products\AAB6F137689A4A549863C7A3AAAA67B0" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Teams" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "registry::HKEY_USERS\$($usersid)\SOFTWARE\Microsoft\Office\Teams" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$($usersid)\Products\AAB6F137689A4A549863C7A3AAAA67B0" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null
Remove-Item -LiteralPath "registry::HKEY_USERS\$($usersid)\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams" -ErrorAction SilentlyContinue -Force -Recurse -Confirm:$false | Out-Null

#Delete teams temp files

if(Test-Path("C:\Program Files (x86)\Microsoft\TeamsMeetingAddin"))
{
    Remove-Item -Path "C:\Program Files (x86)\Microsoft\TeamsMeetingAddin" -Recurse -Force
}

if(Test-Path("C:\Program Files (x86)\Microsoft\TeamsPresenceAddin"))
{
    Remove-Item -Path "C:\Program Files (x86)\Microsoft\TeamsPresenceAddin" -Recurse -Force
}

if(Test-Path("C:\Program Files (x86)\Microsoft\Teams"))
{
    Remove-Item -Path "C:\Program Files (x86)\Microsoft\Teams" -Recurse -Force
}

if(Test-Path("C:\Program Files (x86)\Teams Installer"))
{
    Remove-Item -Path "C:\Program Files (x86)\Teams Installer" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Roaming\Teams"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Roaming\Teams" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Roaming\Microsoft Teams"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Roaming\Microsoft Teams" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Roaming\Microsoft\Teams"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Roaming\Microsoft\Teams" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Local\SquirrelTemp"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Local\SquirrelTemp" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Local\Microsoft\TeamsMeetingAddin"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Local\Microsoft\TeamsMeetingAddin" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Local\Microsoft\TeamsPresenceAddin"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Local\Microsoft\TeamsPresenceAddin" -Recurse -Force
}

if(Test-Path("C:\Users\$username\AppData\Local\Microsoft\Teams"))
{
    Remove-Item -Path "C:\Users\$username\AppData\Local\Microsoft\Teams" -Recurse -Force
}

if(Test-Path("$env:APPDATA\Teams"))
{
    Remove-Item -Path "$env:APPDATA\Teams" -Recurse -Force
}

if(Test-Path("$env:APPDATA\Microsoft Teams"))
{
    Remove-Item -Path "$env:APPDATA\Microsoft Teams" -Recurse -Force
}

if(Test-Path("$env:APPDATA\Microsoft\Teams"))
{
    Remove-Item -Path "$env:APPDATA\Microsoft\Teams" -Recurse -Force
}

if(Test-Path("$env:LOCALAPPDATA\SquirrelTemp"))
{
    Remove-Item -Path "$env:LOCALAPPDATA\SquirrelTemp" -Recurse -Force
}

if(Test-Path("$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"))
{
    Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin" -Recurse -Force
}

if(Test-Path("$env:LOCALAPPDATA\Microsoft\TeamsPresenceAddin"))
{
    Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\TeamsPresenceAddin" -Recurse -Force
}

if(Test-Path("$env:LOCALAPPDATA\Microsoft\Teams"))
{
    Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Teams" -Recurse -Force
}

if(Test-Path("$env:PUBLIC\Desktop\Microsoft Teams.lnk"))
{
    Remove-Item -Path "$env:PUBLIC\Desktop\Microsoft Teams.lnk" -Force
}

$UsersDesktop = Get-ItemPropertyValue -LiteralPath "registry::HKEY_USERS\$($usersid)\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "Desktop"

if(Test-Path("$UsersDesktop\Microsoft Teams.lnk"))
{
    Remove-Item -Path "$UsersDesktop\Microsoft Teams.lnk" -Force
}
